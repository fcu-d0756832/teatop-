from datetime import datetime
import os
import pandas as pd
from openpyxl import load_workbook
import re
import gc
import win32com.client as win32
import matplotlib.pyplot as plt
from pandas.plotting import table
import pyautogui
import time
import subprocess
import pyperclip
import shutil
import json




file_name_1 = "會計拋轉需要_明細表_"
file_name_2 = "會計拋轉需要_簡要表_"

grouped_data = []
current_date = None
current_sum = 0
opening_balance_sum = 0  # ✅ 新增：庫存期初 / 無日期


#明細表的位置
file_name_path = r"\\192.168.2.253\\管理中心\\財會課\\加盟店應收\\應收拋轉(心豪)"
#明細表日期是簡要表減1天工作天
date1_start ="20250607"
date1_end = "20260324"
#簡要表日期2個工作天
date2_start ="20250607"
date2_end = "20260324"
# 判斷星期幾
dt = datetime.strptime(date2_end, "%Y%m%d")
weekday_index = dt.weekday()

base_path = r"C:\\Users\\楊心豪\\Downloads"
store_list_json = r"\\192.168.2.253\\事業中心\\資訊\\code\\store_list.json"
# 你的 NAS 路徑（請依實際情況修改）
nas_base_path = r"\\192.168.2.253\\管理中心\\財會課\\加盟店應收\\應收拋轉(心豪)"
output_folder = os.path.join(base_path, date2_end)
# 確保資料夾存在（不存在就建立）
os.makedirs(output_folder, exist_ok=True)
# 1. 讀入完整代號 vs 門市名稱 Excel
#df_all = pd.read_excel(r"C:\\Users\\楊心豪\Desktop\\爬蟲\\TEATOP店家通訊錄(最後更新2025.3.3).xlsx")

from datetime import datetime, timedelta

def subtract_workday(date_str, days=1):
    dt = datetime.strptime(date_str, "%Y%m%d")
    while days > 0:
        dt -= timedelta(days=1)
        if dt.weekday() < 5:  # 0~4 = 工作天
            days -= 1
    return dt

# 簡要表實際日期
summary_dt = datetime.strptime(date2_end, "%Y%m%d")

# 明細表實際日期 = 簡要表減 1 個工作天
detail_dt = summary_dt

# ======================
# 星期判斷（只算一次）
# ======================
weekday_dict = {
    0: "monday",
    1: "tuesday",
    2: "wednesday",
    3: "thursday",
    4: "friday",
    5: "saturday",
    6: "sunday",
}

weekday_index = summary_dt.weekday()
weekday_name = weekday_dict[weekday_index]

print("📅 簡要表日期：", summary_dt.strftime("%Y-%m-%d"))
print("📅 明細表日期：", detail_dt.strftime("%Y-%m-%d"))
print("📅 使用星期：", weekday_name)

# ======================
# 檔案路徑
# ======================
output_folder = os.path.join(base_path, date2_end)
os.makedirs(output_folder, exist_ok=True)

file_name_full = file_name_2 + date2_end
source_file = os.path.join(
    file_name_path,
    file_name_full + ".xlsx"
)

cleaned_file = os.path.join(
    r"C:\\Users\\楊心豪\\Downloads",
    "已清理" + file_name_full + ".xlsx"
)

# ======================
# 清理簡要表（刪第一列）
# ======================
sheet_names = pd.ExcelFile(source_file).sheet_names

with pd.ExcelWriter(cleaned_file, engine="openpyxl") as writer:
    for sheet in sheet_names:
        df = pd.read_excel(source_file, sheet_name=sheet, header=None)
        df = df.iloc[1:].reset_index(drop=True)
        df.to_excel(writer, sheet_name=sheet, index=False, header=False)

print("✅ 所有工作表的第一列已刪除，儲存為：", cleaned_file)

source_file = cleaned_file

# ======================
# 讀 store_list.json
# ======================
with open(store_list_json, "r", encoding="utf-8") as f:
    store_list = json.load(f)

# ======================
# 建立 weekday → {代號: 店名}
# ======================
dicts_by_week = {}

for store_name, info in store_list.items():
    weekday = info.get("weekday")
    code = str(info.get("code", "")).strip().upper()
    if not weekday or not code:
        continue
    dicts_by_week.setdefault(weekday, {})[code] = store_name

# ======================
# 取得當天要用的門市字典
# ======================
selected_dict = dicts_by_week.get(weekday_name, {})

print(f"📦 {weekday_name} 門市數量：", len(selected_dict))

# ======================
# 載入 Excel，準備後續處理
# ======================
wb = load_workbook(source_file)

def get_customer_id(ws):
    for row in ws.iter_rows(min_row=1, max_row=10, max_col=5):
        for cell in row:
            val = str(cell.value) if cell.value else ""
            if "客戶編號" in val:
                print(f"📌 找到原始內容：{val}")
                if "：" in val:
                    return val.split("：")[1].strip().upper()
    return ""


import os
import pandas as pd
from openpyxl import load_workbook



# 儲存符合條件的分頁資料
for sheetname in wb.sheetnames:
    ws = wb[sheetname]

    customer_id = get_customer_id(ws)
    store_name = str(ws["C6"].value).strip() if ws["C6"].value else ""
    if customer_id in selected_dict:
        file_name = f"{store_name}_簡要表_{date2_end}.xlsx"
        full_path = os.path.join(output_folder, file_name)

        # 不使用第一列當標題
        df = pd.read_excel(source_file, sheet_name=sheetname, engine="openpyxl", header=None)

        # 刪除第一列（原始格式或空白列）
        df = df.iloc[1:].reset_index(drop=True)

        # 儲存檔案（不加 index）
        df.to_excel(full_path, index=False, header=False)

        print(f"已儲存：{full_path}")
    else:
        print(f"分頁 {sheetname} 的客戶編號 {customer_id} 不在選擇字典中")

# -------- 開始讀取儲存後的檔案，抽取資料 --------
import re

result_list = []

for filename in os.listdir(output_folder):
    if filename.endswith(".xlsx"):
        filepath = os.path.join(output_folder, filename)
        df = pd.read_excel(filepath, header=None, engine="openpyxl")

        store_name = None
        date_row_idx = None
        date_col_idx = None
        unpaid_col_idx = None
        prepayment = 0  # 預收金額初始化

        # 找出公司名稱
        for i in range(df.shape[0]):
            row = df.iloc[i]
            for j in range(len(row)):
                if str(row[j]).strip() == "公司名稱：":
                    for k in range(j + 1, len(row)):
                        val = str(row[k]).strip()
                        if val and val != "0" and val.lower() != "nan":
                            store_name = val
                            break

        # 找出「日期」和「未收金額」欄位位置
        for i in range(df.shape[0]):
            row = df.iloc[i]
            for j in range(len(row)):
                if "日期" in str(row[j]):
                    date_row_idx = i
                    date_col_idx = j
                if "未收金額" in str(row[j]):
                    unpaid_col_idx = j
            if date_row_idx is not None and unpaid_col_idx is not None:
                break

        if date_row_idx is None or unpaid_col_idx is None:
            print(f"⚠️ 檔案 {filename} 找不到『日期』或『未收金額』欄位，跳過")
            continue

        opening_balance_sum = 0
        current_date = None
        current_sum = 0
        grouped_data = []

        for i in range(date_row_idx + 1, df.shape[0]):
            row = df.iloc[i]
            row_str_combined = " ".join(map(str, row.values))

            # 偵測結尾並抓取預收
            if "本幣累計預收貨款金額" in row_str_combined:
                parts = row_str_combined.split("本幣累計預收貨款金額")
                match = re.search(r"(\d[\d,.]*)", parts[-1])
                prepayment = float(match.group().replace(",", "")) if match else 0
                break 

            # 🚩 全行掃描：跳過任何包含「庫存期初」的列
            if "庫存期初" in row_str_combined:
                continue 

            # 日期與加總邏輯
            raw_date = str(row[date_col_idx]).strip()
            unpaid_val = row[unpaid_col_idx]
            parsed_date = pd.to_datetime(raw_date, errors='coerce')

            if pd.notna(parsed_date):
                if current_date is not None:
                    grouped_data.append((current_date, current_sum))
                current_date = parsed_date
                current_sum = 0

            if pd.notna(unpaid_val):
                try:
                    val = float(str(unpaid_val).replace(",", "").strip())
                    if current_date is None:
                        opening_balance_sum += val
                    else:
                        current_sum += val
                except:
                    pass

        # --- 5. 最終彙總計算 ---
        # (此處的 previous_period_amount 將是扣除庫存期初後的純淨金額)

        # --- 5. 最終結算該店金額 ---
        if current_date is not None:
            grouped_data.append((current_date, current_sum))

        target_date = pd.to_datetime(date2_end, format="%Y%m%d", errors="coerce")
        
        # 本期：剛好是 target_date 的加總
        current_period_amount = sum(v for d, v in grouped_data if d == target_date)
        
        # 上期：排除期初後的 opening_balance_sum + 其他日期的加總
        previous_period_amount = opening_balance_sum + sum(v for d, v in grouped_data if d != target_date)
        
        # 總計
        adjusted_total = current_period_amount + previous_period_amount - prepayment

        if store_name:
            result_list.append({
                "店名": store_name,
                "上期": previous_period_amount,
                "本期": current_period_amount,
                "預收": prepayment,
                "本幣期末應收帳款總計(本期+上期-預收)": adjusted_total
            })
        else:
            # ✅ 這裡的 print 也要同步修正，確保報錯時顯示的是正確的變數
            print(f"❗檔案 {filename} 缺少資訊：{store_name=}, {total_receivable=}, {prepayment=}")

# 匯出彙總表
summary_df = pd.DataFrame(result_list)

# 再次確保「預收」為數字且 NaN → 0
summary_df["預收"] = pd.to_numeric(summary_df["預收"], errors="coerce").fillna(0)
summary_df["本幣期末應收帳款總計(本期+上期-預收)"] = (
    summary_df["上期"] + summary_df["本期"] - summary_df["預收"]
)

summary_df.to_excel(os.path.join(output_folder, "_應收彙總.xlsx"), index=False)
print(summary_df)


time.sleep(2)
# 啟動 Excel
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False  # 不顯示 Excel 介面

for filename in os.listdir(output_folder):
    if filename.endswith(".xlsx") and not filename.endswith("_應收彙總.xlsx"):
        filepath = os.path.join(output_folder, filename)
        pdfpath = filepath.replace(".xlsx", ".pdf")

        try:
            wb = excel.Workbooks.Open(filepath)
            ws = wb.Sheets(1)  # 假設每個檔案都只有一個 sheet

            # 🧠 找到最後一個有內容的儲存格（自動推算範圍）
            last_row = ws.UsedRange.Rows.Count
            last_col = ws.UsedRange.Columns.Count

            # 🖨️ 設定列印區域
            ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col)).Address

            ws.UsedRange.Font.Bold = True
            ws.UsedRange.Font.Size = 12  # 可選，加大字體（10~12 都很舒服）

            # ⚙️ 設定縮放：一頁寬、高不限
            ws.PageSetup.Orientation = 2 
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = False

            # 匯出成 PDF
            wb.ExportAsFixedFormat(0, pdfpath)
            print(f"✅ 已轉換為 PDF：{pdfpath}")

            wb.Close(False)

        except Exception as e:
            print(f"❌ {filename} 發生錯誤：{e}")


f# 關閉 Excel
excel.Quit()

# 🔥 刪除除了 "_應收彙總.xlsx" 以外的所有 xlsx
for filename in os.listdir(output_folder):
    if filename.endswith(".xlsx") and not filename.endswith("_應收彙總.xlsx"):
        try:
            os.remove(os.path.join(output_folder, filename))
            print(f"🗑️ 已刪除：{filename}")
        except Exception as e:
            print(f"❌ 刪除 {filename} 發生錯誤：{e}")



# 🧾 來源 Excel：會計拋轉需要_明細表
source_file = os.path.join(file_name_path, f"會計拋轉需要_明細表_{date1_end}.xlsx")

# 📁 輸出資料夾（與簡要表同一個）
output_folder = os.path.join(base_path, date2_end)
os.makedirs(output_folder, exist_ok=True)

# 📃 取得簡要表的名單
summary_store_names = set()
for filename in os.listdir(output_folder):
    if filename.endswith("_應收彙總.xlsx"):
        df = pd.read_excel(os.path.join(output_folder, filename), engine="openpyxl")
        summary_store_names.update(df["店名"].astype(str).tolist())

time.sleep(5)

# ⚙️ 開啟 Excel 應用程式
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(source_file)
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(source_file)
time.sleep(5)
# 🔍 建立：公司名稱 → 對應工作表名清單
company_sheets = {}

for sheet in wb.Sheets:
    df = pd.read_excel(source_file, sheet_name=sheet.Name, engine="openpyxl", header=None)

    company_name = None
    for i in range(min(20, df.shape[0])):
        row = df.iloc[i]
        for val in row:
            if isinstance(val, str) and "公司名稱" in val:
                match = re.search(r"公司名稱[:：\s]*(\S+)", val)
                if match:
                    company_name = match.group(1).strip()
                break
        if company_name:
            break

    if company_name:
        company_sheets.setdefault(company_name, []).append(sheet.Name)

# --- 修正後的明細表輸出區塊 ---
for company, sheets in company_sheets.items():
    if company not in summary_store_names:
        continue

    safe_name = re.sub(r'[\\/*?:"<>|]', "_", company)
    pdf_path = os.path.join(output_folder, f"{safe_name}_明細表_{date1_end}.pdf")

    try:
        temp_wb = excel.Workbooks.Add()
        default_sheet = temp_wb.Sheets(1)
        copied = False

        for sheet_name in reversed(sheets):
            sheet = wb.Sheets(sheet_name)
            sheet.Copy(Before=temp_wb.Sheets(1))
            copied = True

        if copied:
            try:
                temp_wb.Sheets(temp_wb.Sheets.Count).Delete()
            except:
                pass

            for ws in temp_wb.Sheets:
                # 1. 🧹 刪除前三列 (正航標題列清理)
                ws.Rows("1:3").Delete()

                # 2. ✨ 字體修正：解決亂碼並確保清晰
                ws.Cells.Font.Name = "微軟正黑體"
                ws.Cells.Font.Size = 9  # 稍微縮小字體基數，讓壓縮後更清晰

                # 3. 📐 強制定義列印範圍 (A 欄到 Z 欄)
                last_row = ws.UsedRange.Rows.Count
                ws.PageSetup.PrintArea = ws.Range("A1:Z" + str(last_row)).Address

                # 4. 📏 關鍵設定：強制縮在一個 A4 寬度
                ws.PageSetup.Zoom = False            # 🚩 必須為 False，下方的 FitToPages 才會生效
                ws.PageSetup.FitToPagesWide = 1      # 🚩 強制所有欄位 (A-Z) 擠在一頁寬度內
                ws.PageSetup.FitToPagesTall = False  # 🚩 長度自然分頁，不強制壓縮高度 (避免變太小)
                
                ws.PageSetup.Orientation = 2         # 橫向列印
                ws.PageSetup.PaperSize = 9           # A4 紙張
                ws.PageSetup.CenterHorizontally = True 
                
                # 5. 🤏 邊距極小化 (爭取更多顯示空間)
                ws.PageSetup.LeftMargin = 5          
                ws.PageSetup.RightMargin = 5
                ws.PageSetup.TopMargin = 10
                ws.PageSetup.BottomMargin = 10

            # 📤 執行匯出
            temp_wb.ExportAsFixedFormat(0, pdf_path)
            print(f"✅ 已輸出最適比例 PDF (A-Z)：{pdf_path}")

        temp_wb.Close(False)
    except Exception as e:
        print(f"❌ {company} 錯誤：{e}")


# 🔚 關閉 Excel
wb.Close(False)
excel.Quit()
del excel
gc.collect()







# 組合完整目標路徑（例如：\\NAS_SERVER\共享資料\帳款報表\20240805）
nas_target_path = os.path.join(nas_base_path, date2_end)

try:
    # 如果目的資料夾已存在，先刪除（避免 shutil.move 報錯）
    if os.path.exists(nas_target_path):
        shutil.rmtree(nas_target_path)

    # 搬移整個資料夾
    shutil.move(output_folder, nas_target_path)
    print(f"📂 資料夾已成功搬移到 NAS：{nas_target_path}")

except Exception as e:
    print(f"❌ 搬移至 NAS 失敗：{e}")   


