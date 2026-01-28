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
st.set_page_config(page_title="自動對帳系統整合版", page_icon="📊", layout="wide")
st.title("📊 自動對帳系統 (整合版)")

# 側邊欄：選擇功能
mode = st.sidebar.radio("請選擇對帳功能：", ["🚗 洗車對帳 (Code A)", "📺 LiTV 對帳 (Code B)"])

# ==========================================
# 🔴 功能 A：洗車對帳邏輯
# ==========================================
def process_car_wash(file_a, file_b):
    output = io.BytesIO()
    logs = []

    try:
        # --- 設定區 ---
        sheet_name_billing = '請款'
        sheet_name_details = '累計明細'
        col_id = '訂單編號'
        col_plate = '車牌'
        col_refund = '退款時間'
        col_phone = '手機號碼'
        
        target_month_str = datetime.now().strftime("%Y/%m")

        logs.append(f"📂 正在讀取檔案...")
        xls_a = pd.ExcelFile(file_a)

        # 1. 讀取 A 表 (請款) - 自動找標題
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

        # 2. 準備 A 表詳細資料
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

        # 3. 準備 B 表詳細資料
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

        # 4. 合併對帳
        cols_keep = [col_id, col_plate, col_phone]
        df_total = pd.merge(
            df_a[cols_keep], df_b[cols_keep],
            on=[col_id, col_plate], how='outer', indicator=True, suffixes=('_A', '_B')
        )

        logs.append(f"✅ 對帳完成: A表有效筆數 {int(val_count)}, B表退款筆數 {len(df_b_refunds)}")

        # 5. 寫入 Excel
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb = writer.book
            fmt_header = wb.add_format({'bold': True, 'bg_color': '#EFEFEF', 'border': 1, 'align': 'center'})
            fmt_content = wb.add_format({'border': 1, 'align': 'center'})
            fmt_currency = wb.add_format({'num_format': '#,##0', 'border': 1, 'align': 'right'})
            fmt_blue = wb.add_format({'bg_color': '#DDEBF7'})
            fmt_pink = wb.add_format({'bg_color': '#FCE4D6'})

            # Sheet 1: 請款
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

            # Sheet 2: 對帳總表
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
# 🔵 功能 B：LiTV 對帳邏輯 (修正：相容「金額」與「方案金額」)
# ==========================================
def process_litv(file_a, file_b):
    output = io.BytesIO()
    logs = []

    try:
        # 1. 複製 B 表作為基底 (在記憶體中操作)
        file_b_bytes = io.BytesIO(file_b.getvalue())
        wb = openpyxl.load_workbook(file_b_bytes)
        
        # 2. 處理報表 A (自動尋找標題列，解決 KeyError 問題)
        logs.append("正在讀取 A 表並搜尋標題列...")
        
        # 預讀前 20 行找標題
        df_temp = pd.read_excel(file_a, header=None, nrows=20)
        header_idx = -1
        
        # 修正：同時尋找 '金額' 或 '方案金額'
        for i, row in df_temp.iterrows():
            row_str = " ".join([str(x).strip() for x in row.values])
            if '訂單編號' in row_str and ('金額' in row_str or '方案金額' in row_str):
                header_idx = i
                break
        
        if header_idx == -1:
            return None, ["❌ 錯誤：在 A 表中找不到「訂單編號」或「金額 / 方案金額」欄位，請確認檔案格式。"], None, None

        # 正式讀取 A 表
        df_a = pd.read_excel(file_a, header=header_idx)
        df_a.columns = df_a.columns.str.strip() # 去空白
        
        # 修正：如果欄位叫做 '方案金額'，將其改名為 '金額' 以便後續處理
        if '方案金額' in df_a.columns:
            df_a.rename(columns={'方案金額': '金額'}, inplace=True)
            logs.append("💡 提示：偵測到「方案金額」欄位，已自動視為「金額」處理。")
            
        # 再次檢查確保欄位存在
        if '金額' not in df_a.columns:
            return None, [f"❌ 錯誤：雖然找到標題列，但找不到「金額」欄位。目前欄位有：{list(df_a.columns)}"], None, None

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
        a_lookup_set = set(zip(df_a_filtered['手機隱碼'], df_a_filtered['方案(SKU)']))

        # 3. 處理報表 B (ACG對帳明細)
        logs.append("正在處理 B 表 (ACG對帳明細)...")
        # 重新從原始檔讀取資料供 pandas 使用
        df_b_acg_full = pd.read_excel(file_b, sheet_name='ACG對帳明細')
        df_b_acg_full.columns = df_b_acg_full.columns.str.strip()

        stop_idx = None
        for idx, val in enumerate(df_b_acg_full['編號']):
            if "不計費" in str(val):
                stop_idx = idx
                break
        
        if stop_idx is not None:
            df_b_valid = df_b_acg_full.iloc[:stop_idx].copy()
        else:
            df_b_valid = df_b_acg_full.copy()

        df_b_valid = df_b_valid.dropna(subset=['手機/虛擬帳號', '廠商對帳key1']).copy()
        df_b_valid['手機/虛擬帳號'] = df_b_valid['手機/虛擬帳號'].astype(str).str.strip()
        df_b_valid['廠商對帳key1'] = df_b_valid['廠商對帳key1'].astype(str).str.strip()
        b_lookup_set = set(zip(df_b_valid['手機/虛擬帳號'], df_b_valid['廠商對帳key1']))

        # 4. 對帳與收集差異
        sku_mapping = {'LiTV_LUX_1Y_OT': ['LiTV_LUX_1Y_OT', 'LiTV_LUX_F1MF_1Y_OT'], 'LiTV_LUX_1M_OT': ['LiTV_LUX_1M_OT']}
        reverse_sku_map = {'LiTV_LUX_F1MF_1Y_OT': 'LiTV_LUX_1Y_OT', 'LiTV_LUX_1Y_OT': 'LiTV_LUX_1Y_OT', 'LiTV_LUX_1M_OT': 'LiTV_LUX_1M_OT'}

        sheet1_data = []
        diff_a_not_b = []

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

        diff_b_not_a = []
        for _, row in df_b_valid.iterrows():
            b_phone, b_key = str(row['手機/虛擬帳號']).strip(), str(row['廠商對帳key1']).strip()
            if "*" in b_phone:
                equiv_sku = reverse_sku_map.get(b_key, b_key)
                if (b_phone, equiv_sku) not in a_lookup_set:
                    diff_b_not_a.append({'手機/虛擬帳號': b_phone, '廠商對帳key1': b_key})

        # 5. 修改 Excel (Openpyxl)
        yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

        # 新增分頁
        if "CMX對帳明細" in wb.sheetnames: del wb["CMX對帳明細"]
        ws_new = wb.create_sheet("CMX對帳明細", 0)
        headers = ['廠商方案代碼', '廠商方案名稱', '手機/虛擬帳號', '方案金額', 'CMX訂單編號']
        ws_new.append(headers)
        for data in sheet1_data:
            ws_new.append([data[h] for h in headers])
            if data['is_diff']:
                for cell in ws_new[ws_new.max_row]: cell.fill = yellow_fill

        # 標記 ACG對帳明細
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

        wb.save(output)
        logs.append(f"✅ 對帳完成: A有B無 {len(diff_a_not_b)} 筆，B有A無 {len(diff_b_not_a)} 筆")
        return output.getvalue(), logs, diff_a_not_b, diff_b_not_a

    except Exception as e:
        return None, [f"❌ 嚴重錯誤: {str(e)}"], None, None

# ==========================================
# 介面顯示邏輯
# ==========================================
if mode == "🚗 洗車對帳 (Code A)":
    st.header("🚗 洗車訂單對帳")
    col1, col2 = st.columns(2)
    file_a = col1.file_uploader("上傳 A 表 (請款明細)", type=['xlsx', 'xls'])
    file_b = col2.file_uploader("上傳 B 表 (廠商報表)", type=['xlsx', 'xls'])
    
    if st.button("開始對帳", type="primary"):
        if file_a and file_b:
            with st.spinner("資料處理中..."):
                result, logs = process_car_wash(file_a, file_b)
            
            st.expander("查看執行紀錄", expanded=True).write(logs)
            
            if result:
                st.success("成功！請下載結果：")
                st.download_button(
                    label="📥 下載洗車對帳結果",
                    data=result,
                    file_name=f"洗車對帳_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("⚠️ 請確認兩個檔案都已上傳。")

elif mode == "📺 LiTV 對帳 (Code B)":
    st.header("📺 LiTV 訂單對帳")
    st.info("💡 系統會自動搜尋標題列，無需手動調整。")
    
    col1, col2 = st.columns(2)
    file_a = col1.file_uploader("上傳 A 表 (report_supplier...)", type=['xlsx', 'xls'])
    file_b = col2.file_uploader("上傳 B 表 (車美仕對帳單...)", type=['xlsx', 'xls'])
    
    if st.button("開始對帳", type="primary"):
        if file_a and file_b:
            with st.spinner("比對資料中..."):
                result, logs, diff_a, diff_b = process_litv(file_a, file_b)
            
            st.expander("查看執行紀錄", expanded=True).write(logs)
            
            if result:
                st.success("成功！")
                c1, c2 = st.columns(2)
                c1.error(f"A有B無 (共 {len(diff_a) if diff_a else 0} 筆)")
                if diff_a: c1.dataframe(pd.DataFrame(diff_a))
                
                c2.warning(f"B有A無 (共 {len(diff_b) if diff_b else 0} 筆)")
                if diff_b: c2.dataframe(pd.DataFrame(diff_b))
                
                st.download_button(
                    label="📥 下載 LiTV 對帳結果",
                    data=result,
                    file_name=f"LiTV_對帳_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("⚠️ 請確認兩個檔案都已上傳。")
