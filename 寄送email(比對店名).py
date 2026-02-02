import os
import re
import json
import mimetypes
import base64
import pandas as pd
import datetime as dt
import holidays
from email.message import EmailMessage
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError
import platform

nas_target_path = r"\\192.168.10.253\管理中心\財會課\加盟店應收\應收拋轉(心豪)\20260202(don't remove)"

def normalize_store_name(name: str) -> str:
    if not name:
        return ""

    name = name.strip()

    # 移除結尾的「店」
    if name.endswith("店"):
        name = name[:-1]

    # 全形空白、特殊符號（可選）
    name = name.replace("　", "").replace(" ", "")

    return name


# === 新增：取得台灣國定假日（含跨年） ===
def get_tw_holidays(year_start: int, year_end: int):
    """
    回傳一個 set，內含台灣國定假日的日期（datetime.date）。
    會嘗試兩種寫法以相容不同版本的 python-holidays。
    """
    all_h = set()
    for y in range(year_start, year_end + 1):
        try:
            tw = holidays.Taiwan(years=y)
        except Exception:
            tw = holidays.country_holidays("TW", years=y)
        # tw.keys() 即為日期（datetime.date）
        all_h.update(tw.keys())
    return all_h

def add_workdays(start_date, days, holidays_set=None):
    """從 start_date 開始往後加 days 個工作天 (週一~週五，排除國定假日)"""
    if not start_date:
        return None
    if holidays_set is None:
        holidays_set = set()
    current = start_date
    added_days = 0
    while added_days < days:
        current += dt.timedelta(days=1)
        # 0=週一 ... 4=週五；且不在假日表
        if current.weekday() < 5 and current not in holidays_set:
            added_days += 1
    return current

def format_date(d: dt.date) -> str:
    """將日期轉成 m/d 格式，跨平台支援"""
    if not d:
        return "N/A"
    system = platform.system()
    try:
        if system == "Windows":
            return d.strftime("%#m/%#d")  # Windows
        else:
            return d.strftime("%-m/%-d")  # Linux / Mac
    except ValueError:
        return d.strftime("%m/%d")  # fallback

# --- 授權範圍 ---
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# --- 取得 Gmail API 憑證 ---
def get_credentials():
    creds = None
    token_file = r"C:\Users\Public\token.json"
    client_secret_file = r'C:\Users\Public\client_secret_google_cloud.json'

    if os.path.exists(token_file):
        try:
            creds = Credentials.from_authorized_user_file(token_file, SCOPES)
            if not creds.valid:
                if creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    raise RefreshError("token 無法刷新")
        except Exception as e:
            print(f"⚠️ token.json 壞掉或缺少欄位，重新 OAuth: {e}")
            creds = None

    if not creds:
        flow = InstalledAppFlow.from_client_secrets_file(client_secret_file, SCOPES)
        creds = flow.run_local_server(port=8080, access_type="offline", prompt="consent")
        with open(token_file, "w") as token:
            token.write(creds.to_json())

    return creds

# --- 載入店家 JSON ---
with open(r"\\192.168.10.253\\事業中心\\資訊\\code\\store_list.json", "r", encoding="utf-8") as f:
    stores = json.load(f)

# --- 指定要搜尋的資料夾 ---
folder_path = nas_target_path
# --- 讀取周邊聯名商品統計（test.xlsx）---
test_file = os.path.join(folder_path, "戀與貨款金額.xlsx")
test_df = pd.read_excel(test_file, header=1)
raw_period = pd.read_excel(
    test_file,
    header=None
).iloc[0, 0]

period_str = str(raw_period).replace("統計時間：", "").strip()

# 清理欄位名稱（保險）
test_df.columns = (
    test_df.columns
    .astype(str)
    .str.replace(" ", "", regex=False)
    .str.replace("　", "", regex=False)
    .str.strip()
)

merch_dict = {}

for _, row in test_df.iterrows():
    raw_store = row.get("門市")
    if not raw_store:
        continue

    store_key = normalize_store_name(raw_store)

    merch_dict[store_key] = {
        "period": period_str,
        "amount": int(row.get("周邊貨款小計", 0) or 0)
    }

print("DEBUG merch_dict keys =", list(merch_dict.keys()))



# --- 讀取應收彙總 ---
summary_file = os.path.join(folder_path, "_應收彙總.xlsx")
summary_df = pd.read_excel(summary_file)

# 轉成 dict 方便查詢：{店名: {"上期":..., "本期":...}}
summary_dict = {
    str(row["店名"]): {
        "上期": int(row.get("上期", 0) or 0),
        "本期": int(row.get("本期", 0) or 0),
        "預收": int(row.get("預收", 0) or 0),
        "總計": int(row.get("本幣期末應收帳款總計(本期+上期-預收)", 0) or 0)
    }
    for _, row in summary_df.iterrows()
}

# --- 建立 Gmail API 連線 ---
creds = get_credentials()
service = build('gmail', 'v1', credentials=creds)

# --- 分組：先把同一店家的檔案歸在一起 ---
store_files = {}
for file_name in os.listdir(folder_path):
        # 從 PDF 檔名抓「_」前的店名
    pdf_store_raw = file_name.split("_")[0]
    pdf_store_key = normalize_store_name(pdf_store_raw)  # 保險再正規化一次


    file_path = os.path.join(folder_path, file_name)

    if not os.path.isfile(file_path):
        continue
    if not file_name.lower().endswith(".pdf"):
        continue

    matched_store = None
    recipient_email = None

    for store_name, info in stores.items():
        store_key = normalize_store_name(store_name)

        if store_key == pdf_store_key:
            matched_store = store_name          # 保留原始名稱用於顯示、寄信
            recipient_email = info.get("email")
            break


    if not matched_store or not recipient_email:
        print(f"❌ 檔案 {file_name} 沒有找到對應店名，跳過。")
        continue

    if matched_store not in store_files:
        store_files[matched_store] = {
            "email": recipient_email,
            "files": [],
            "pay_method": stores[matched_store].get("pay_method", "匯款")
        }

    store_files[matched_store]["files"].append(file_path)

week_map = ["一", "二", "三", "四", "五", "六", "日"]

for store_name, info in store_files.items():
    recipient_email = info["email"]
    file_paths = info["files"]
    pay_method = info.get("pay_method", "匯款")  # 預設為匯款

    # 從彙總表抓金額
    last_amt = summary_dict.get(store_name, {}).get("上期", 0)
    curr_amt = summary_dict.get(store_name, {}).get("本期", 0)
    prepay_amt = summary_dict.get(store_name, {}).get("預收", 0)
    total_amt = summary_dict.get(store_name, {}).get("總計")
    # --- 聯名周邊提醒文字（不影響帳款）---
    merch_text = ""
    pdf_store_key = normalize_store_name(store_name)  # 保險再正規化一次
    print("DEBUG store_name raw =", repr(store_name))
    print("DEBUG pdf_store_key =", repr(pdf_store_key))
    print("DEBUG merch_dict keys =", list(merch_dict.keys())[:5])

    if pdf_store_key in merch_dict:
        merch_info = merch_dict[pdf_store_key]
        merch_amt = merch_info.get("amount", 0)
        merch_period = merch_info.get("period", "")

        if merch_amt > 0:
            merch_text = (
                f"\n"
                f"📣 提醒說明：\n"
                f"「統計時間：{merch_period}」之中，"
                f"關於【戀與深空】相關聯名商品（杯墊／杯套／徽章）"
                f"之銷售對應款項統計為 {merch_amt:,} 元整，"
                f"總公司將於 3/30（一）起開始收款。\n"
            )

    if not total_amt:  # 如果是 None 或 0，都自己算
        total_amt = last_amt + curr_amt - prepay_amt
    # --- 聯名周邊提醒（依 PDF 店名比對 test.xlsx）---



    # 從附件檔名抓日期
    date_last, date_curr = None, None
    for fp in file_paths:
        m = re.search(r"_(\d{8})", os.path.basename(fp))
        if m:
            d = dt.datetime.strptime(m.group(1), "%Y%m%d").date()
            if "明細表" in fp:
                date_last = d
            elif "簡要表" in fp:
                date_curr = d

    # 若沒找到明細表日期，就用簡要表日期 - 1 天當作上期日期
    if not date_last and date_curr:
        date_last = date_curr - dt.timedelta(days=1)

    date_last_str = format_date(date_last)
    date_curr_str = format_date(date_curr)

    # 匯款截止日：簡要表日期 + 3 個「工作天」（避開六日 + 台灣國定假日）
    if date_curr:
        # 準備跨年假日集合（保險可延伸到 +2 年）
        tw_holidays = get_tw_holidays(date_curr.year, date_curr.year + 1)
        tw_holidays.add(dt.date(2025, 9, 29))
        repay_date = add_workdays(date_curr, 3, holidays_set=tw_holidays)
        repay_str = f"{repay_date.month}/{repay_date.day}({week_map[repay_date.weekday()]})"
    else:
        repay_str = "N/A"

    # 依付款方式決定信件內容
    if pay_method == "匯款":
        pay_text = f"明細如下，到貨後核對無誤麻煩於 {repay_str} 12:00前匯款，謝謝。"
    elif pay_method == "現金":
        pay_text = "明細如下，到貨後核對無誤麻煩將貨款給司機帶回,謝謝"
    elif isinstance(pay_method, str) and pay_method.upper() == "ACH":
        pay_text = "請將款項存入ACH扣款帳戶"
    else:
        pay_text = "明細如下，請依約定付款方式處理，謝謝"

    subject = f"📌 {store_name} 每週應收帳款"
    body_text = f"""
    您好，{store_name} 負責人：

    截至{date_last_str}未清帳款為 {last_amt} 元，{date_curr_str}貨款金額為 {curr_amt} 元，
    (實際帳款以出貨單為準)

    _{last_amt}_ + _{curr_amt}_ - _{prepay_amt}_ = __{total_amt}__
    {merch_text}
    {pay_text}
    """


    message = EmailMessage()
    message.set_content(body_text)
    message['To'] = recipient_email
    message['From'] = "teatop0048@gmail.com"
    message['Subject'] = subject

    # 加入所有附件（保護 mimetype 可能為 None 的情況）
    for file_path in file_paths:
        file_name = os.path.basename(file_path)
        mime_type, _ = mimetypes.guess_type(file_path)
        if mime_type:
            maintype, mime_subtype = mime_type.split('/', 1)
        else:
            maintype, mime_subtype = "application", "octet-stream"
        with open(file_path, 'rb') as fp:
            message.add_attachment(
                fp.read(),
                maintype=maintype,
                subtype=mime_subtype,
                filename=file_name
            )

    encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    send_message = service.users().messages().send(
        userId="me",
        body={"raw": encoded_message}
    ).execute()

    print(f"📤 已寄出 {store_name} ({len(file_paths)} 個附件) → {recipient_email} (ID: {send_message['id']})")
    if merch_text == "" and pdf_store_key in merch_dict:
        print(f"⚠️ {store_name} 周邊金額為 0，未顯示提醒")
    elif merch_text == "" and pdf_store_key not in merch_dict:
        print(f"ℹ️ {store_name} 無聯名周邊資料")
    if normalize_store_name(store_name) != pdf_store_key:
        print(f"❌ 跳過：PDF 店名與系統店名不一致 → {pdf_store_raw} / {store_name}")
        continue

