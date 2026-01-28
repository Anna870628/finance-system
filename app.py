import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
import xlsxwriter
import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime

# ==========================================
# 頁面基本設定
# ==========================================
st.set_page_config(page_title="自動對帳系統", page_icon="📊", layout="wide")
st.title("📊 自動對帳系統 (智慧相容版)")

# 側邊欄：選擇功能
mode = st.sidebar.radio("請選擇對帳功能：", ["🚗 洗車對帳 (Code A)", "📺 LiTV 對帳 (Code B)"])

# ==========================================
# 🔴 功能 A：洗車對帳邏輯 (維持不變)
# ==========================================
def process_car_wash(file_a, file_b):
    output = io.BytesIO()
    logs = []

    try:
        sheet_name_billing = '請款'
        sheet_name_details = '累計明細'
        col_id = '訂單編號'
        col_plate = '車牌'
        col_refund = '退款時間'
        col_phone = '手機號碼'
        target_month_str = datetime.now().strftime("%Y/%m")

        logs.append(f"📂 正在讀取檔案...")
        xls_a = pd.ExcelFile(file_a)

        # 自動找標題 (洗車專用)
        df_temp = pd.read_excel(xls_a, sheet_name=sheet_name_billing, header=None, usecols="A:E", nrows=20)
        header_row_idx = 2
        for i, row in df_temp.iterrows():
            row_str = " ".join([str(x) for x in row.values])
            if '提供日期' in row_str:
                header_row_idx = i
                break
        
        df_daily = pd.read_excel(xls_a, sheet_name=sheet_name_billing, header=header_row_idx, usecols="A:E")
        
        if len(df_daily.columns) >= 5:
            val_count = pd.to_numeric(df_daily.iloc[:, 1], errors='coerce').fillna(0).sum()
            val_billing = pd.to_numeric(df_daily.iloc[:, 2], errors='coerce').fillna(0).sum()
            val_sms = pd.to_numeric(df_daily.iloc[:, 4], errors='coerce').fillna(0).sum()
            val_total = val_billing + val_sms
        else:
            val_count, val_billing, val_sms, val_total = 0, 0, 0, 0

        if not df_daily.empty:
            col_date = df_daily.columns[0]
            df_daily[col_date] = pd.to_datetime(df_daily[col_date], errors='coerce').dt.strftime('%Y-%m-%d')
            df_daily = df_daily.dropna(subset=[col_date])

        # A 表詳細
        df_details = pd.read_excel(xls_a, sheet_name=sheet_name_details)
        df_a = df_details.dropna(subset=[col_id]).copy()
        df_a[col_id] = df_a[col_id].astype(str).str.strip()
        df_a = df_a[~df_a[col_id].str.contains('合計|Total|總計', case=False, na=False)]
        if col_plate in df_a.columns:
            df_a[col_plate] = df_a[col_plate].astype(str).str.strip()
        if col_phone not in df_a.columns:
            df_a[col_phone] = ""
        else:
            df_a[col_phone] = df_a[col_phone].astype(str).str.strip()
        df_a = df_a.drop_duplicates(subset=[col_id, col_plate])

        # B 表詳細
        if hasattr(file_b, 'seek'): file_b.seek(0)
        df_b_original = pd.read_excel(file_b, sheet_name=0, header=2)
        df_b_processing = df_b_original.copy()
        df_b_refunds = pd.DataFrame()
        if col_refund in df_b_processing.columns:
            df_b_refunds = df_b_processing[df_b_processing[col_refund].notna()].copy()
            df_b_filtered = df_b_processing[df_b_processing[col_refund].isna()]
        else:
            df_b_filtered = df_b_processing
        
        df_b = df_b_filtered.dropna(subset=[col_id]).copy()
        df_b[col_id] = df_b[col_id].astype(str).str.strip()
        df_b[col_plate] = df_b[col_plate].astype(str).str.strip()
        if col_phone not in df_b.columns:
            df_b[col_phone] = ""
        else:
            df_b[col_phone] = df_b[col_phone].astype(str).str.strip()
        df_b = df_b.drop_duplicates(subset=[col_id, col_plate])

        # 合併
        cols_keep = [col_id, col_plate, col_phone]
        df_total = pd.merge(
            df_a[cols_keep], df_b[cols_keep],
            on=[col_id, col_plate], how='outer', indicator=True, suffixes=('_A', '_B')
        )

        logs.append(f"✅ 對帳完成: A表有效筆數 {int(val_count)}, B表退款筆數 {len(df_b_refunds)}")

        # 寫入 Excel
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb = writer.book
            fmt_header = wb.add_format({'bold': True, 'bg_color': '#EFEFEF', 'border': 1, 'align': 'center'})
            fmt_content = wb.add_format({'border': 1, 'align': 'center'})
            fmt_currency = wb.add_format({'num_format': '#,##0', 'border': 1, 'align': 'right'})
            fmt_blue = wb.add_format({'bg_color': '#DDEBF7'})
            fmt_pink = wb.add_format({'bg_color': '#FCE4D6'})

            ws1 = wb.add_worksheet('請款')
            writer.sheets['請款'] = ws1
            headers = ['統計月份', '轉檔筆數', '轉檔請款金額', '簡訊請款金額', '合計金額']
            values = [target_month_str, val_count, val_billing, val_sms, val_total]
            
            for col, (h, v) in enumerate(zip(headers, values)):
                ws1.write(0, col, h, fmt_header)
                ws1.write(1, col, v, fmt_currency if isinstance(v, (int, float)) else fmt_content)
            
            for col, h in enumerate(df_daily.columns):
                ws1.write(3, col, h, fmt_header)
            df_daily.to_excel(writer, sheet_name='請款', startrow=4, header=False, index=False)

            df_total.to_excel(writer, sheet_name='對帳總表', index=False)
            ws2 = writer.sheets['對帳總表']
            for i, val in enumerate(df_total['_merge']):
                if val == 'left_only': ws2.set_row(i+1, None, fmt_blue)
                elif val == 'right_only': ws2.set_row(i+1, None, fmt_pink)
            
            df_total[df_total['_merge'] == 'left_only'].drop(columns=['_merge']).to_excel(writer, sheet_name='僅A表有', index=False)
            df_total[df_total['_merge'] == 'right_only'].drop(columns=['_merge']).to_excel(writer, sheet_name='僅B表有', index=False)
            
            if not df_b_refunds.empty:
                df_b_refunds.to_excel(writer, sheet_name='B表退款排除名單', index=False)

        return output.getvalue(), logs

    except Exception as e:
        return None, [f"❌ 錯誤: {str(e)}"]

# ==========================================
# 🔵 功能 B：LiTV 對帳邏輯 (智慧容錯版)
# ==========================================
def process_litv(file_a, file_b):
    output = io.BytesIO()
    logs = []

    try:
        # --- 1. 複製 B 表作為基底 ---
        file_b_bytes = io.BytesIO(file_b.getvalue())
        wb = openpyxl.load_workbook(file_b_bytes)
        
        # --- 2. 處理報表 A (智慧讀取) ---
        logs.append("正在讀取 A 表...")
        file_a.seek(0)
        
        # [STEP 1] 先試你原本的 header=2
        try:
            df_a = pd.read_excel(file_a, header=2)
            df_a.columns = df_a.columns.str.strip()
        except:
            df_a = pd.DataFrame() # 讀取失敗就給空

        # [STEP 2] 檢查是否讀到正確欄位
        # 如果找不到 '金額' 且找不到 '方案金額'，代表 header=2 是錯的 (可能這份檔案 header 在第 0 行)
        if '金額' not in df_a.columns and '方案金額' not in df_a.columns:
            logs.append("⚠️ 原始設定 (header=2) 找不到金額欄位，嘗試切換為標準格式 (header=0)...")
            file_a.seek(0)
            df_a = pd.read_excel(file_a, header=0)
            df_a.columns = df_a.columns.str.strip()
        
        # [STEP 3] 欄位名稱校正 (把 '方案金額' 改成 '金額')
        if '方案金額' in df_a.columns:
            df_a.rename(columns={'方案金額': '金額'}, inplace=True)
            logs.append("💡 將「方案金額」視為「金額」。")
            
        # [STEP 4] 最終檢查
        if '金額' not in df_a.columns:
            # 還是找不到，報錯並列出所有欄位讓你知道發生什麼事
            return None, [f"❌ 嚴重錯誤：找不到「金額」欄位。", f"讀到的欄位有：{list(df_a.columns)}"], None, None

        # --- 以下完全是你原本的邏輯 ---
        df_a['金額'] = pd.to_numeric(df_a['金額'], errors='coerce').fillna(0)

        df_a_filtered = df_a[
            (df_a['金額'] > 0) &
            (df_a['退款時間'].isna()) &
            (df_a['手機號碼'].notna())
        ].copy()

        def fix_phone_a(val):
            if pd.isna(val): return ""
            s = str(val).split('.')[0]
            if len(s) == 9: s = '0' + s
            return s

        df_a_filtered['手機全碼'] = df_a_filtered['手機號碼'].apply(fix_phone_a)
        df_a_filtered['手機隱碼'] = df_a_filtered['手機全碼'].apply(lambda x: x[:6] + '****' if len(x) >= 10 else x)
        df_a_filtered['方案(SKU)'] = df_a_filtered['方案(SKU)'].astype(str).str.strip()
        a_lookup_set = set(zip(df_a_filtered['手機隱碼'], df_a_filtered['方案(SKU)'].str.strip()))

        # --- 3. 處理報表 B ---
        logs.append("正在處理 B 表...")
        file_b.seek(0)
        df_b_acg_full = pd.read_excel(file_b, sheet_name='ACG對帳明細')
        df_b_acg_full.columns = df_b_acg_full.columns.str.strip()

        stop_idx = None
