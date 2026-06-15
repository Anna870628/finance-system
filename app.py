import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import xlsxwriter
import openpyxl
from openpyxl.styles import PatternFill, Font
from datetime import datetime

# ==========================================
# 輔助函式：手機號碼格式化
# ==========================================
def normalize_phone(val):
    """
    將手機號碼轉為字串，去除 .0，並確保 09 開頭
    """
    if pd.isna(val) or val == "":
        return ""
    
    # 轉字串並去除前後空白
    s = str(val).strip()
    
    # 處理浮點數轉字串可能產生的 .0
    if s.endswith(".0"):
        s = s[:-2]
        
    if len(s) == 9 and s.startswith("9"):
        s = "0" + s
        
    return s

# ==========================================
# 頁面基本設定
# ==========================================
st.set_page_config(page_title="自動對帳系統 (介面優化版)", page_icon="📊", layout="wide")
st.title("📊 自動對帳系統")

mode = st.sidebar.radio("請選擇對帳功能：", ["🚗 洗車對帳 (Code A)", "📺 LiTV 對帳 (Code B)"])

# ==========================================
# 🚗 功能 A：洗車對帳邏輯 (動態工作表防呆升級版)
# ==========================================
def process_car_wash(files_supplier_upload, file_billing_upload):
    output = io.BytesIO()
    logs = []
    output_filename = "洗車對帳結果.xlsx"

    try:
        if file_billing_upload:
            base_name = os.path.splitext(file_billing_upload.name)[0]
            output_filename = f"{base_name}_CMX確認.xlsx"

        file_billing_upload.seek(0)

        col_id = '訂單編號'
        col_plate = '車牌'
        col_refund = '退款時間'
        col_phone = '手機號碼'
        target_month_str = datetime.now().strftime("%Y/%m")

        # ---------------------------------------------------------
        # 1. 處理左側檔案 (廠商報表 / A表) - 支援多檔合併
        # ---------------------------------------------------------
        logs.append(f"📂 正在讀取左側檔案 (廠商報表/A表)，共收到 {len(files_supplier_upload)} 份檔案...")
        
        df_a_list = []
        for file_supplier in files_supplier_upload:
            file_supplier.seek(0)
            df_temp = pd.read_excel(file_supplier, sheet_name=0, header=2)
            df_a_list.append(df_temp)
            logs.append(f"   ↳ 成功讀取: {file_supplier.name} ({len(df_temp)} 筆)")
            
        df_a_original = pd.concat(df_a_list, ignore_index=True)
        df_a_processing = df_a_original.copy()
        
        df_a_refunds = pd.DataFrame()
        if col_refund in df_a_processing.columns:
            df_a_refunds = df_a_processing[df_a_processing[col_refund].notna()].copy()
            df_a_filtered = df_a_processing[df_a_processing[col_refund].isna()]
        else:
            df_a_filtered = df_a_processing
        
        df_a = df_a_filtered.dropna(subset=[col_id]).copy()
        df_a[col_id] = df_a[col_id].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        
        if col_plate in df_a.columns:
            df_a[col_plate] = df_a[col_plate].astype(str).str.replace(r'[-\s]', '', regex=True).str.
