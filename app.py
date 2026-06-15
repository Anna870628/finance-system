# ==========================================
# 🚗 功能 A：洗車對帳邏輯 (支援動態工作表名稱與多個 A 表)
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

        # 基礎欄位定義
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
            df_a[col_plate] = df_a[col_plate].astype(str).str.replace(r'[-\s]', '', regex=True).str.upper()
        
        if col_phone not in df_a.columns:
            df_a[col_phone] = ""
        else:
            df_a[col_phone] = df_a[col_phone].apply(normalize_phone)

        df_a = df_a.drop_duplicates(subset=[col_id, col_plate])
        logs.append(f"   ↳ 總計合併去重後，A表共有 {len(df_a)} 筆有效資料")

        # ---------------------------------------------------------
        # 2. 處理右側檔案 (請款明細 / B表) - 【動態防呆升級版】
        # ---------------------------------------------------------
        logs.append(f"📂 正在讀取右側檔案 (請款明細/B表)...")
        xls_b = pd.ExcelFile(file_billing_upload)
        
        # 獲取該 Excel 的所有工作表名稱
        available_sheets = xls_b.sheet_names
        logs.append(f"   ↳ 🔍 偵測到 B 表內包含的工作表有：{available_sheets}")

        # 🎯 動態判定「請款」工作表
        sheet_name_billing = '請款'
        if sheet_name_billing not in available_sheets:
            billing_candidates = [s for s in available_sheets if '請款' in s]
            if billing_candidates:
                sheet_name_billing = billing_candidates[0]
                logs.append(f"   ⚠️ 找不到精確的 '請款' 表，自動匹配使用：'{sheet_name_billing}'")
            else:
                sheet_name_billing = available_sheets[0]
                logs.append(f"   ⚠️ 找不到任何包含 '請款' 的表，預設使用第 1 個工作表：'{sheet_name_billing}'")

        # 🎯 動態判定「明細」工作表 (解決你遇到的主要錯誤)
        sheet_name_details = '累計明細'
        if sheet_name_details not in available_sheets:
            # 模糊搜尋名稱中含有「明細」或「detail」的工作表
            details_candidates = [s for s in available_sheets if '明細' in s or 'detail' in s.lower()]
            if details_candidates:
                sheet_name_details = details_candidates[0]
                logs.append(f"   ⚠️ 找不到精確的 '累計明細'，自動模糊匹配使用：'{sheet_name_details}'")
            else:
                # 如果連關鍵字都找不到，通常防呆預設第 2 個工作表為明細表
                if len(available_sheets) > 1:
                    sheet_name_details = available_sheets[1]
                    logs.append(f"   ⚠️ 找不到任何包含 '明細' 的工作表，自動彈性指定第 2 個工作表：'{sheet_name_details}'")
                else:
                    sheet_name_details = available_sheets[0]
                    logs.append(f"   ⚠️ 檔案內僅有 1 個工作表，強制指定使用：'{sheet_name_details}'")

        logs.append(f"   🚀 最終對帳選用工作表 ➔ 摘要請款表: '{sheet_name_billing}' | 累計明細表: '{sheet_name_details}'")

        # 開始讀取資料
        df_temp = pd.read_excel(xls_b, sheet_name=sheet_name_billing, header=None, usecols="A:E", nrows=20)
        header_row_idx = 2
        for i, row in df_temp.iterrows():
            row_str = " ".join([str(x) for x in row.values])
            if '提供日期' in row_str:
                header_row_idx = i
                break
        
        df_daily = pd.read_excel(xls_b, sheet_name=sheet_name_billing, header=header_row_idx, usecols="A:E")
        
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

        # 使用上面動態偵測出來的名稱讀取明細
        df_details = pd.read_excel(xls_b, sheet_name=sheet_name_details)
        df_b = df_details.dropna(subset=[col_id]).copy()
        
        df_b[col_id] = df_b[col_id].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        df_b = df_b[~df_b[col_id].str.contains('合計|Total|總計', case=False, na=False)]
        
        if col_plate in df_b.columns:
            df_b[col_plate] = df_b[b_plate] = df_b[col_plate].astype(str).str.replace(r'[-\s]', '', regex=True).str.upper()
            
        if col_phone not in df_b.columns:
            df_b[col_phone] = ""
        else:
            df_b[col_phone] = df_b[col_phone].apply(normalize_phone)
            
        df_b = df_b.drop_duplicates(subset=[col_id, col_plate])

        # 統計洗金寶與三合一的筆數
        sanheyi_count = 0
        wash_count = 0
        if '金額' in df_b.columns:
            sanheyi_count += df_b['金額'].astype(str).str.contains('三合一', na=False).sum()
        elif '方案(SKU)' in df_b.columns:
            sanheyi_count += df_b['方案(SKU)'].astype(str).str.contains('三合一', na=False).sum()
            
        wash_count = len(df_b) - sanheyi_count
        logs.append(f"   ↳ 📊 方案筆數統計 (依據 B表)：洗金寶 {wash_count} 筆，三合一 {sanheyi_count} 筆")

        # ---------------------------------------------------------
        # 3. 合併對帳
        # ---------------------------------------------------------
        cols_keep = [col_id, col_plate, col_phone]
        df_total = pd.merge(
            df_a[cols_keep], 
            df_b[cols_keep], 
            on=[col_id, col_plate], 
            how='outer', 
            indicator=True, 
            suffixes=('_A', '_B')
        )

        logs.append(f"✅ 對帳完成: CMX合併A表 {len(df_a)} 筆, TMS請款B表 {len(df_b)} 筆")

        # ---------------------------------------------------------
        # 4. 寫入 Excel
        # ---------------------------------------------------------
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb = writer.book
            
            base_font_size = 12
            header_font_size = 14

            fmt_header = wb.add_format({'bold': True, 'bg_color': '#EFEFEF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': header_font_size})
            fmt_content = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': base_font_size})
            fmt_currency = wb.add_format({'num_format': '#,##0', 'border': 1, 'align': 'right', 'valign': 'vcenter', 'font_size': base_font_size})
            fmt_blue = wb.add_format({'bg_color': '#DDEBF7', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': base_font_size})
            fmt_pink = wb.add_format({'bg_color': '#FCE4D6', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': base_font_size})
            fmt_text_month = wb.add_format({'num_format': '@', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': base_font_size})
            fmt_bold_total = wb.add_format({'bold': True, 'num_format': '#,##0', 'border': 1, 'bg_color': '#FFF2CC', 'align': 'right', 'valign': 'vcenter', 'font_size': base_font_size})

            ws1 = wb.add_worksheet('請款')
            top_headers = ['統計月份', '轉檔筆數', '轉檔請款金額', '簡訊請款金額', '合計金額']
            top_values = [target_month_str, val_count, val_billing, val_sms, val_total]
            
            ws1.set_row(0, 30)
            ws1.set_row(1, 25)

            for col, (header, val) in enumerate(zip(top_headers, top_values)):
                ws1.write(0, col, header, fmt_header)
                if col == 0: ws1.write(1, col, val, fmt_text_month)
                elif col == 4: ws1.write(1, col, val, fmt_bold_total)
                else:
                    if isinstance(val, (int, float)): ws1.write(1, col, val, fmt_currency)
                    else: ws1.write(1, col, val, fmt_content)
            
            for col_idx, col_name in enumerate(df_daily.columns):
                ws1.write(3, col_idx, col_name, fmt_header)
            
            for r, row in enumerate(df_daily.values):
                for c, val in enumerate(row):
                    ws1.write(r + 4, c, val, fmt_content)
            
            ws1.set_column('A:A', 25) 
            ws1.set_column('B:E', 25) 

            ws2 = wb.add_worksheet('對帳總表')
            columns = df_total.columns.tolist()
            for c_idx, col_name in enumerate(columns):
                ws2.write(0, c_idx, col_name, fmt_header)
            
            ws2.set_column(0, len(columns)-1, 25)
            ws2.set_row(0, 22)

            for r_idx, row in df_total.iterrows():
                merge_status = row['_merge']
                if merge_status == 'left_only': current_fmt = fmt_blue
                elif merge_status == 'right_only': current_fmt = fmt_pink
                else: current_fmt = fmt_content
                
                excel_row = r_idx + 1
                ws2.set_row(excel_row, 18) 

                for c_idx, val in enumerate(row):
                    write_val = "" if pd.isna(val) else val
                    ws2.write(excel_row, c_idx, write_val, current_fmt)

            df_total[df_total['_merge'] == 'left_only'].drop(columns=['_merge']).to_excel(writer, sheet_name='僅A表有', index=False)
            df_total[df_total['_merge'] == 'right_only'].drop(columns=['_merge']).to_excel(writer, sheet_name='僅B表有', index=False)
            
            if not df_a_refunds.empty:
                df_a_refunds.to_excel(writer, sheet_name='A表退款排除名單', index=False)

        return output.getvalue(), logs, output_filename

    except Exception as e:
        import traceback
        return None, [f"❌ 錯誤: {str(e)}", traceback.format_exc()], None
