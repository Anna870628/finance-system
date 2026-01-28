import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill
import io
import os

# ==========================================
# 頁面基本設定
# ==========================================
st.set_page_config(page_title="LiTV 對帳系統 (Colab 移植版)", page_icon="📺", layout="wide")
st.title("📺 LiTV 對帳系統 (Colab 移植版)")
st.caption("完全依照原版邏輯設計：A表讀取第3行標題、B表讀取 ACG 對帳明細")

# ==========================================
# 核心邏輯 (完全復刻原版)
# ==========================================
def process_litv(file_a_upload, file_b_upload):
    # 建立一個記憶體輸出的 Buffer
    output_buffer = io.BytesIO()
    logs = []

    try:
        # --- 0. 檔案前置處理 (自動防呆：左右互換) ---
        # 為了避免使用者傳錯邊，我們先檢查 sheet name
        # 邏輯：B表必須包含 'ACG對帳明細'
        
        # 預讀 sheet names (不讀內容，速度快)
        xl_a = pd.ExcelFile(file_a_upload)
        xl_b = pd.ExcelFile(file_b_upload)
        
        file_a_target = file_a_upload
        file_b_target = file_b_upload

        # 如果 A 檔有 ACG 明細，B 檔沒有 -> 代表使用者傳反了，自動交換
        if 'ACG對帳明細' in xl_a.sheet_names and 'ACG對帳明細' not in xl_b.sheet_names:
            logs.append("💡 偵測到檔案位置相反，已自動交換 A/B 表。")
            file_a_target = file_b_upload
            file_b_target = file_a_upload
        
        # 確保指標歸零 (Streamlit 必須做這步)
        file_a_target.seek(0)
        file_b_target.seek(0)

        # --- 1. 複製 B 表作為基底 ---
        # Colab 原碼: shutil.copy(file_b_path, output_name)
        # Streamlit 改寫: 將 B 表載入到 openpyxl 物件
        logs.append("正在載入 B 表 (基底)...")
        wb = openpyxl.load_workbook(file_b_target)

        # --- 2. 處理報表 A (比對基準) ---
        logs.append("正在讀取 A 表 (header=2)...")
        
        # Colab 原碼: df_a = pd.read_excel(file_a_path, header=2)
        # Streamlit 改寫: 直接讀取上傳物件
        df_a = pd.read_excel(file_a_target, header=2)
        df_a.columns = df_a.columns.str.strip()
        
        # --- [關鍵檢查] ---
        # 如果因為任何原因讀不到金額，這里會報錯，我們加一個簡單的檢查提示使用者
        if '金額' not in df_a.columns:
             return None, [f"❌ 錯誤：A 表 (header=2) 讀不到「金額」欄位。讀到的欄位是：{list(df_a.columns)}"], None, None

        # Colab 原碼: df_a['金額'] = pd.to_numeric(...)
        df_a['金額'] = pd.to_numeric(df_a['金額'], errors='coerce').fillna(0)

        # Colab 原碼: 篩選邏輯
        df_a_filtered = df_a[
            (df_a['金額'] > 0) &
            (df_a['退款時間'].isna()) &
            (df_a['手機號碼'].notna())
        ].copy()

        # Colab 原碼: 手機號碼處理函式
        def fix_phone_a(val):
            if pd.isna(val): return ""
            s = str(val).split('.')[0]
            if len(s) == 9: s = '0' + s
            return s

        df_a_filtered['手機全碼'] = df_a_filtered['手機號碼'].apply(fix_phone_a)
        df_a_filtered['手機隱碼'] = df_a_filtered['手機全碼'].apply(lambda x: x[:6] + '****' if len(x) >= 10 else x)
        
        # Colab 原碼: 建立 lookup set
        a_lookup_set = set(zip(df_a_filtered['手機隱碼'], df_a_filtered['方案(SKU)'].str.strip()))

        # --- 3. 處理報表 B (ACG對帳明細) 與截斷邏輯 ---
        logs.append("正在處理 B 表 (ACG對帳明細)...")
        
        # 必須重置 B 表讀取指標給 pandas 用
        file_b_target.seek(0)
        
        # Colab 原碼: df_b_acg_full = pd.read_excel(...)
        df_b_acg_full = pd.read_excel(file_b_target, sheet_name='ACG對帳明細')
        df_b_acg_full.columns = df_b_acg_full.columns.str.strip()

        # Colab 原碼: 尋找「不計費」
        stop_idx = None
        for idx, val in enumerate(df_b_acg_full['編號']):
            if "不計費" in str(val):
                stop_idx = idx
                break

        # Colab 原碼: 截斷資料
        if stop_idx is not None:
            df_b_valid = df_b_acg_full.iloc[:stop_idx].copy()
        else:
            df_b_valid = df_b_acg_full.copy()

        # Colab 原碼: 清洗資料
        df_b_valid = df_b_valid.dropna(subset=['手機/虛擬帳號', '廠商對帳key1']).copy()
        df_b_valid['手機/虛擬帳號'] = df_b_valid['手機/虛擬帳號'].astype(str).str.strip()
        df_b_valid['廠商對帳key1'] = df_b_valid['廠商對帳key1'].astype(str).str.strip()
        b_lookup_set = set(zip(df_b_valid['手機/虛擬帳號'], df_b_valid['廠商對帳key1']))

        # --- 4. 對帳與收集差異數據 ---
        logs.append("正在執行比對邏輯...")
        sku_mapping = {'LiTV_LUX_1Y_OT': ['LiTV_LUX_1Y_OT', 'LiTV_LUX_F1MF_1Y_OT'], 'LiTV_LUX_1M_OT': ['LiTV_LUX_1M_OT']}
        reverse_sku_map = {'LiTV_LUX_F1MF_1Y_OT': 'LiTV_LUX_1Y_OT', 'LiTV_LUX_1Y_OT': 'LiTV_LUX_1Y_OT', 'LiTV_LUX_1M_OT': 'LiTV_LUX_1M_OT'}

        sheet1_data = []
        diff_a_not_b = []

        # Colab 原碼: A 比 B
        for _, row in df_a_filtered.iterrows():
            sku_a = str(row['方案(SKU)']).strip()
            phone_masked = row['手機隱碼']
            possible_keys = sku_mapping.get(sku_a, [sku_a])
            found_in_b = any((phone_masked, k) in b_lookup_set for k in possible_keys)

            if sku_a == 'LiTV_LUX_1M_OT':
                out_sku, out_amt, out_name = 'LiTV_LUX_1M_OT', 187, '豪華雙享餐/月繳/單次(定價$250)'
            elif sku_a == 'LiTV_LUX_1Y_OT':
                out_sku, out_amt, out_name = 'LiTV_LUX_F1MF_1Y_OT', 1717, '豪華雙享餐-首月免費/年繳/單次(定價$2,290)'
            else:
                out_sku, out_amt, out_name = sku_a, row['金額'], sku_a

            sheet1_data.append({
                '廠商方案代碼': out_sku, '廠商方案名稱': out_name, '手機/虛擬帳號': phone_masked,
                '方案金額': out_amt, 'CMX訂單編號': row['訂單編號'], 'is_diff': not found_in_b
            })

            if not found_in_b:
                diff_a_not_b.append({'手機號碼': row['手機全碼'], '方案': sku_a, '訂單編號': row['訂單編號']})

        # Colab 原碼: B 比 A
        diff_b_not_a = []
        for _, row in df_b_valid.iterrows():
            b_phone, b_key = str(row['手機/虛擬帳號']).strip(), str(row['廠商對帳key1']).strip()
            if "*" in b_phone:
                equiv_sku = reverse_sku_map.get(b_key, b_key)
                if (b_phone, equiv_sku) not in a_lookup_set:
                    diff_b_not_a.append({'手機/虛擬帳號': b_phone, '廠商對帳key1': b_key})

        # --- 6. 修改 Excel 標註 ---
        logs.append("正在寫入 Excel...")
        yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

        # A. CMX對帳明細 (新增分頁)
        if "CMX對帳明細" in wb.sheetnames: del wb["CMX對帳明細"]
        ws_new = wb.create_sheet("CMX對帳明細", 0)
        headers = ['廠商方案代碼', '廠商方案名稱', '手機/虛擬帳號', '方案金額', 'CMX訂單編號']
        ws_new.append(headers)
        for data in sheet1_data:
            ws_new.append([data[h] for h in headers])
            if data['is_diff']:
                for cell in ws_new[ws_new.max_row]: cell.fill = yellow_fill

        # B. ACG對帳明細 (標色區間受 stop_idx 限制)
        if 'ACG對帳明細' in wb.sheetnames:
            ws_acg = wb['ACG對帳明細']
            h_list = [cell.value for cell in ws_acg[1]]
            
            # 確保欄位存在
            if '手機/虛擬帳號' in h_list and '廠商對帳key1' in h_list:
                p_idx = h_list.index('手機/虛擬帳號') + 1
                k_idx = h_list.index('廠商對帳key1') + 1
                
                max_reconcile_row = (stop_idx + 1) if stop_idx is not None else ws_acg.max_row
                
                for r_idx in range(2, max_reconcile_row + 1):
                    p_val = str(ws_acg.cell(row=r_idx, column=p_idx).value).strip()
                    k_val = str(ws_acg.cell(row=r_idx, column=k_idx).value).strip()
                    if "*" in p_val:
                        equiv_sku = reverse_sku_map.get(k_val, k_val)
                        if (p_val, equiv_sku) not in a_lookup_set:
                            for cell in ws_acg[r_idx]: cell.fill = yellow_fill
        
        # 儲存到 Buffer
        wb.save(output_buffer)
        return output_buffer.getvalue(), logs, diff_a_not_b, diff_b_not_a

    except Exception as e:
        return None, [f"❌ 嚴重程式錯誤: {str(e)}"], None, None


# ==========================================
# 介面顯示區
# ==========================================

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. 上傳 A 表 (Supplier Report)")
    file_a = st.file_uploader("廠商報表 (請款明細)", type=['xlsx', 'xls'], key='a')
    st.info("💡 邏輯：讀取第 3 行作為標題 (header=2)")

with col2:
    st.subheader("2. 上傳 B 表 (ACG 對帳單)")
    file_b = st.file_uploader("車美仕對帳單 (含 ACG對帳明細)", type=['xlsx', 'xls'], key='b')
    st.info("💡 邏輯：尋找「ACG對帳明細」工作表")

if st.button("🚀 開始對帳", type="primary"):
    if file_a and file_b:
        with st.spinner("對帳中..."):
            result_bytes, logs, diff_a, diff_b = process_litv(file_a, file_b)
        
        # 顯示 Log
        with st.expander("執行紀錄 (Logs)", expanded=True):
            for log in logs:
                st.write(log)

        if result_bytes:
            st.success("✅ 對帳成功！")
            
            # 顯示差異預覽
            c1, c2 = st.columns(2)
            c1.error(f"🟥 A有B無 (共 {len(diff_a)} 筆)")
            if diff_a: c1.dataframe(pd.DataFrame(diff_a))
            
            c2.warning(f"🟨 B有A無 (共 {len(diff_b)} 筆)")
            if diff_b: c2.dataframe(pd.DataFrame(diff_b))

            st.download_button(
                label="📥 下載對帳結果 (Excel)",
                data=result_bytes,
                file_name="LiTV_CMX確認.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ 請上傳這兩個檔案！")
