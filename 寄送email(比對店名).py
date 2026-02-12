import os
import re
import json
import mimetypes
import base64
import pandas as pd
import datetime as dt
import holidays
import platform
from email.message import EmailMessage
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError

# ===============================
# 基本設定
# ===============================
nas_target_path = r"\\192.168.10.253\管理中心\財會課\加盟店應收\應收拋轉(心豪)\20260210"
folder_path = nas_target_path

SCOPES = ['https://www.googleapis.com/auth/gmail.send']


# ===============================
# 店名正規化（唯一標準）
# ===============================
def normalize_store_name(name: str) -> str:
    if not isinstance(name, str) or not name:
        return ""

    name = name.strip()

    # 取 "(" 前
    name = name.split("(")[0]

    # 移除結尾「店」
    if name.endswith("店"):
        name = name[:-1]

    # 移除全形 / 半形空白
    name = name.replace("　", "").replace(" ", "")

    return name


# ===============================
# 假日 / 工作天
# ===============================
def get_tw_holidays(year_start: int, year_end: int):
    all_h = set()
    for y in range(year_start, year_end + 1):
        try:
            tw = holidays.Taiwan(years=y)
        except Exception:
            tw = holidays.country_holidays("TW", years=y)
        all_h.update(tw.keys())
    return all_h


def add_workdays(start_date, days, holidays_set=None):
    if not start_date:
        return None
    if holidays_set is None:
        holidays_set = set()

    current = start_date
    added = 0
    while added < days:
        current += dt.timedelta(days=1)
        if current.weekday() < 5 and current not in holidays_set:
            added += 1
    return current


def format_date(d: dt.date) -> str:
    if not d:
        return "N/A"
    try:
        if platform.system() == "Windows":
            return d.strftime("%#m/%#d")
        else:
            return d.strftime("%-m/%-d")
    except Exception:
        return d.strftime("%m/%d")


# ===============================
# Gmail 憑證
# ===============================
def get_credentials():
    creds = None
    token_file = r"C:\Users\Public\token.json"
    client_secret_file = r"C:\Users\Public\client_secret_google_cloud.json"

    if os.path.exists(token_file):
        try:
            creds = Credentials.from_authorized_user_file(token_file, SCOPES)
            if not creds.valid:
                if creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    raise RefreshError("Token 無法刷新")
        except Exception:
            creds = None

    if not creds:
        flow = InstalledAppFlow.from_client_secrets_file(client_secret_file, SCOPES)
        creds = flow.run_local_server(port=8080, access_type="offline", prompt="consent")
        with open(token_file, "w") as f:
            f.write(creds.to_json())

    return creds


# ===============================
# 載入店家清單
# ===============================
with open(r"\\192.168.10.253\事業中心\資訊\code\store_list.json", "r", encoding="utf-8") as f:
    stores = json.load(f)


# ===============================
# 讀取聯名周邊（戀與貨款金額.xlsx）
# ===============================
test_file = os.path.join(folder_path, "戀與貨款金額.xlsx")
test_df = pd.read_excel(test_file, header=1)

raw_period = pd.read_excel(test_file, header=None).iloc[0, 0]
period_str = str(raw_period).replace("統計時間：", "").strip()

test_df.columns = (
    test_df.columns.astype(str)
    .str.replace(" ", "", regex=False)
    .str.replace("　", "", regex=False)
    .str.strip()
)

merch_dict = {}

for _, row in test_df.iterrows():
    raw_store = row.get("門市")
    if not raw_store:
        continue

    key = normalize_store_name(raw_store)
    merch_dict[key] = {
        "period": period_str,
        "amount": int(row.get("周邊貨款小計", 0) or 0)
    }


# ===============================
# 讀取應收彙總
# ===============================
summary_file = os.path.join(folder_path, "_應收彙總.xlsx")
summary_df = pd.read_excel(summary_file)

summary_dict = {}
for _, row in summary_df.iterrows():
    raw_store = row.get("店名")
    if not raw_store:
        continue

    key = normalize_store_name(raw_store)
    summary_dict[key] = {
        "上期": int(row.get("上期", 0) or 0),
        "本期": int(row.get("本期", 0) or 0),
        "預收": int(row.get("預收", 0) or 0),
        "總計": int(row.get("本幣期末應收帳款總計(本期+上期-預收)", 0) or 0)
    }


# ===============================
# Gmail 連線
# ===============================
service = build("gmail", "v1", credentials=get_credentials())


# ===============================
# PDF 分組
# ===============================
store_files = {}

for file_name in os.listdir(folder_path):
    if not file_name.lower().endswith(".pdf"):
        continue

    file_path = os.path.join(folder_path, file_name)
    if not os.path.isfile(file_path):
        continue

    pdf_store_raw = file_name.split("_")[0]
    pdf_store_key = normalize_store_name(pdf_store_raw)

    for store_name, info in stores.items():
        if normalize_store_name(store_name) == pdf_store_key:
            store_files.setdefault(
                store_name,
                {
                    "email": info.get("email"),
                    "files": [],
                    "pay_method": info.get("pay_method", "匯款")
                }
            )["files"].append(file_path)
            break


# ===============================
# 寄信
# ===============================
week_map = ["一", "二", "三", "四", "五", "六", "日"]

for store_name, info in store_files.items():
    store_key = normalize_store_name(store_name)

    recipient_email = info["email"]
    file_paths = info["files"]
    pay_method = info["pay_method"]

    last_amt = summary_dict.get(store_key, {}).get("上期", 0)
    curr_amt = summary_dict.get(store_key, {}).get("本期", 0)
    prepay_amt = summary_dict.get(store_key, {}).get("預收", 0)
    total_amt = summary_dict.get(store_key, {}).get("總計")

    if not total_amt:
        total_amt = last_amt + curr_amt - prepay_amt

    merch_text = ""
    if store_key in merch_dict and merch_dict[store_key]["amount"] > 0:
        merch_text = (
            f"\n📣 提醒說明：\n"
            f"「統計時間：{merch_dict[store_key]['period']}」之中，"
            f"關於【戀與深空】相關聯名商品（杯墊／杯套／徽章）"
            f"之銷售對應款項統計為 {merch_dict[store_key]['amount']:,} 元整，"
            f"總公司將於 3/30（一）起開始收款。\n"
        )

    date_last, date_curr = None, None
    for fp in file_paths:
        m = re.search(r"_(\d{8})", os.path.basename(fp))
        if m:
            d = dt.datetime.strptime(m.group(1), "%Y%m%d").date()
            if "明細表" in fp:
                date_last = d
            elif "簡要表" in fp:
                date_curr = d

    if not date_last and date_curr:
        date_last = date_curr - dt.timedelta(days=1)

    date_last_str = format_date(date_last)
    date_curr_str = format_date(date_curr)

    if date_curr:
        holidays_set = get_tw_holidays(date_curr.year, date_curr.year + 1)
        repay_date = add_workdays(date_curr, 3, holidays_set)
        repay_str = f"{repay_date.month}/{repay_date.day}({week_map[repay_date.weekday()]})"
    else:
        repay_str = "N/A"

    pay_text = (
        f"明細如下，到貨後核對無誤麻煩於 {repay_str} 12:00前匯款，謝謝。"
        if pay_method == "匯款"
        else "明細如下，請依約定付款方式處理，謝謝。"
    )

    subject = f"📌 {store_name} 每週應收帳款"
    body = f"""
您好，{store_name} 負責人：

截至{date_last_str}未清帳款為 {last_amt} 元，
{date_curr_str}貨款金額為 {curr_amt} 元，

_{last_amt}_ + _{curr_amt}_ - _{prepay_amt}_ = __{total_amt}__
{merch_text}
{pay_text}
"""

    msg = EmailMessage()
    msg.set_content(body)
    # --- 清洗收件者 email ---
    recipient_email = str(recipient_email).strip()

    # 如果是 NaN 轉成空字串
    if recipient_email.lower() == "nan":
        recipient_email = ""

    if not recipient_email or "@" not in recipient_email:
        print(f"❌ 無效 Email，跳過寄送：{store_name} → {recipient_email}")
        continue
    msg["To"] = recipient_email
    msg["From"] = "teatop0048@gmail.com"
    msg["Subject"] = subject

    for fp in file_paths:
        mime, _ = mimetypes.guess_type(fp)
        maintype, subtype = mime.split("/", 1) if mime else ("application", "octet-stream")
        with open(fp, "rb") as f:
            msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=os.path.basename(fp))

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(userId="me", body={"raw": raw}).execute()

    print(f"📤 已寄出 {store_name} → {recipient_email}")
