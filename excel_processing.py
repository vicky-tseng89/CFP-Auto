from datetime import datetime
from docx import Document
from docx.shared import Inches
from docxtpl import DocxTemplate, InlineImage
from functools import reduce
from tkinter import filedialog, messagebox
import math
# import matplotlib
# matplotlib.use('Agg')  # 強制使用不會開視窗的 Agg 後端
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
import os
import pandas as pd
import pythoncom
import re
import time
import tkinter as tk
import win32com.client as win32
import xlsxwriter



class ExcelApp:
    def __init__(self, status_callback=None, progress_callback=None):
        self.status_callback = status_callback
        self.progress_callback = progress_callback
        self.file_path = None
        self._format_cache = {}
        # ──────────────────────────────────────────────────────────
        # 自動在程式路徑下建立「結果」資料夾
        base_dir = self.get_base_dir()
        self.output_dir = os.path.join(base_dir, "結果")    # 程式路徑
        os.makedirs(self.output_dir, exist_ok=True)
        # ──────────────────────────────────────────────────────────
    def get_base_dir(self):
        """
        如果是 PyInstaller 打包後的 single-file exe，
        sys.argv[0] 會是使用者實際「雙擊執行」的那顆 .exe 的完整路徑。
        所以把它取 dirname 就能得到 exe 所在資料夾。

        如果是開發階段直接跑 .py，
        __file__ 會是目前 .py 的檔案路徑，我們就取 .py 同層資料夾即可。
        """
        import sys
        if getattr(sys, 'frozen', False):
            # 已經被打包成 exe
            return os.path.dirname(sys.argv[0])
        else:
            # 開發環境跑 .py
            return os.path.dirname(os.path.abspath(__file__))

    def browse_file(self):
        # 瀏覽文件並設置文件路徑
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, self.file_path)

    def process_file(self, file_path=None):
        """數據處理"""
        from win32com.client import DispatchEx
        if file_path is not None:
            self.file_path = file_path
        if not self.file_path:
            messagebox.showerror("錯誤", "請選擇 Excel 文件")
            return

        try:
            self.update_progress_smooth(0, 10, step=1, delay=0.5) # 階段1：讀取 Excel 檔案與資料準備，模擬從 0% 到 10%
            # 使用 openpyxl 讀取原始的 Excel 文件，保留原始格式和樣式
            self.status_callback("讀取 Excel 文件...")
            print("讀取 Excel 文件...")
            result_workbook = openpyxl.load_workbook(self.file_path, keep_vba=False, keep_links=False)
            sheet_A_tables = self.read_multiple_tables('Raw Material', self.file_path)      #呼叫 read_multiple_tables(sheet_name, file_path) 
            sheet_C_tables = self.read_multiple_tables('Manufacturing', self.file_path)     #讀取特定數個工作表（如 Raw Material、Manufacturing 等）
            sheet_D_tables = self.read_multiple_tables('Distribution', self.file_path)      #將每個工作表中多個獨立的資料表格區段解析為 pandas DataFrame 清單
            sheet_E_tables = self.read_multiple_tables('Recycling', self.file_path)
            sheet_F_tables = self.read_multiple_tables('Usage', self.file_path)
            
            self.update_progress_smooth(10, 40, step=1, delay=0.05) # 階段2：讀取工作表B，處理工作表並計算數值，模擬進度從 10% 到 40%
            # 以 pandas 讀入另一張關鍵對照表（sheet_B，如 simapro9.3）
            sheet_B = pd.read_excel(self.file_path, sheet_name='simapro9.3', usecols=['單位對照', 'fossil(kg CO2-eq)', 'biogenic(kg CO2-eq)', 'land transformation (kg CO2-eq)', 'unit']).dropna(subset=['單位對照'])
            self.status_callback("處理工作表並獲取總值...")
            print("處理工作表並獲取總值...")
            total_A = self.process_tables(sheet_A_tables, 'Raw Material', 'W', result_workbook, sheet_B)
            total_C = self.process_tables(sheet_C_tables, 'Manufacturing', 'W', result_workbook, sheet_B)
            total_D = self.process_tables(sheet_D_tables, 'Distribution', 'U', result_workbook, sheet_B)
            total_E = self.process_tables(sheet_E_tables, 'Recycling', 'Q', result_workbook, sheet_B)
            total_F = self.process_tables(sheet_F_tables, 'Usage', 'Q', result_workbook, sheet_B)

            self.update_progress_smooth(40, 70, step=1, delay=0.02) # 階段3：更新報告模板，模擬進度從 40% 到 70%
            self.status_callback("讀取報告模板並寫入計算的數值...")
            print("讀取報告模板並寫入計算的數值...")
            base_dir = os.path.dirname(os.path.abspath(__file__))       # Temp檔路徑下
            report_path = os.path.join(base_dir, 'report_temp.xlsx')    # 將結果寫入報告範本 Excel (report_temp.xlsx) 中預定的儲存格
            report_workbook = openpyxl.load_workbook(report_path)
            # 確保選擇報告中的 'general' 工作表
            if 'general' in report_workbook.sheetnames:
                report_sheet = report_workbook['general']
            else:
                raise ValueError("報告模板中未找到名為 'general' 的工作表。")

            self.status_callback("每個工作表的加總值寫入指定的單元格...")
            print("每個工作表的加總值寫入指定的單元格...")
            # 將每個工作表的加總值寫入指定的單元格
            report_sheet['B2'], report_sheet['B3'], report_sheet['B4'] = total_A
            report_sheet['C2'], report_sheet['C3'], report_sheet['C4'] = total_C
            report_sheet['D2'], report_sheet['D3'], report_sheet['D4'] = total_D
            report_sheet['E2'], report_sheet['E3'], report_sheet['E4'] = total_E
            report_sheet['F2'], report_sheet['F3'], report_sheet['F4'] = total_F
            self.update_progress_smooth(70, 95, step=1, delay=0.05) # 階段4：儲存結果，模擬進度從 70% 到 99%
            # 獲取當前的日期和時間，用於生成檔案名稱
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.status_callback("保存更新後的報告，附上日期和時間...")
            print("保存更新後的報告，附上日期和時間...")
            # 保存更新後的報告，附上日期和時間
            self.report_file = f'report_{current_time}.xlsx'
            self.report_file = os.path.join(self.output_dir, self.report_file)
            print("路徑位置：",self.output_dir)
            report_workbook.save(self.report_file)
            
            # 另存為新文件，保留原有的表格樣式，附上日期和時間
            self.result_file = f'result_{current_time}.xlsx'
            self.result_file = os.path.join(self.output_dir, self.result_file)
            result_workbook.save(self.result_file)

        #    3. 用 Excel COM 自動修復並輸出最終結果
            pythoncom.CoInitialize()
            excel = DispatchEx("Excel.Application")
            excel.Visible = False  # 背後跑就好
            excel.DisplayAlerts = False
            # CorruptLoad=1: 自動嘗試修復任何架構問題；UpdateLinks=0: 不更新外部連結
            com_wb = excel.Workbooks.Open(
                os.path.abspath(self.result_file),
                CorruptLoad=1,
                UpdateLinks=0,
                ReadOnly=False
            )           
            com_wb.Save() 
            com_wb.Close(False)
            excel.Quit()

            if self.update_progress_smooth: # 更新進度至 100%
                self.update_progress_smooth(95, 100, step=1, delay=0.05)
            messagebox.showinfo("完成", f"完成合併並計算碳排數值，結果已保存為 {self.result_file} 和 {self.report_file}")
            return True   # 告知呼叫方：成功
        except Exception as e:
            messagebox.showerror("錯誤", f"處理文件時出錯：{e}")
            print(f"處理文件時出錯：{e}")
            return False  # 告知呼叫方：失敗``

    def read_multiple_tables(self, sheet_name, file_path):
        """
        讀取工作表（如 Raw Material、Manufacturing 等）
        將每個工作表依據辨識B欄的◎符號，分為多個獨立的資料表格區段解析為 pandas DataFrame 清單​
        """
        # 讀取整個工作表，不設定標題行
        sheet = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        tables = []
        start_idx = 0

        # 遍歷所有行，辨識B欄的◎符號所在行來定位表格起始位置
        for idx, row in sheet.iterrows():
            # 檢查該行的B欄是否包含◎符號
            if '◎' in str(row[1]):
                # 如果已經找到一個表格的開始，將其保存
                if start_idx != 0:
                    # 保存表格並使用第三行作為欄位名稱
                    header = sheet.iloc[start_idx + 2]
                    table = sheet.iloc[start_idx + 3:idx].reset_index(drop=True)
                    table.columns = header
                    tables.append((start_idx, table))
                # 更新新的表格開始位置
                start_idx = idx

        # 添加最後一個表格，並使用第三行作為欄位名稱
        header = sheet.iloc[start_idx + 2]
        table = sheet.iloc[start_idx + 3:].reset_index(drop=True)
        table.columns = header
        tables.append((start_idx, table))

        # 返回表格數據
        return tables
    
    def process_tables(self, sheet_tables, sheet_name, col_start, workbook, sheet_B):
        """
        對工作表數據進行處理與加總
        根據不同表格使用不同的欄位進行計算
        並將數據進行單位換算
        """
        total_fossil = 0
        total_biogenic = 0
        total_land_transformation = 0
        for i, (start_idx, sheet_data) in enumerate(sheet_tables):
            # 使用 merge 函數來進行類似 VLOOKUP 的合併
            merged_df = sheet_data.merge(sheet_B, left_on='name of database', right_on='單位對照', how='left', suffixes=('', '_y'))

            # 根據不同表格使用不同的欄位進行計算
            if 'total' in sheet_data.columns:
                quantity_column = 'total'
            elif 'Ton‧Km' in sheet_data.columns:
                quantity_column = 'Ton‧Km'
            elif 'consumed amount allocated to single product (energy/product unit)' in sheet_data.columns:
                quantity_column = 'consumed amount allocated to single product (energy/product unit)'
            elif sheet_name in ['Recycling', 'Usage'] and 'total amount' in sheet_data.columns:
                quantity_column = 'total amount'
            else:
                raise ValueError("表格中缺少必要的計算欄位 ('total'、'Ton‧Km'、'consumed amount allocated to single product (energy/product unit)'或 'total amount')")
            
            # 判斷工作表和工作表B的單位是否一致
            for idx, row in merged_df.iterrows():
                if pd.notna(row['Unit']) and pd.notna(row['unit']):
                    if row['Unit'] != row['unit']:
                        if row['Unit'] in ['g', 'kg', 'ton'] and row['unit'] in ['g', 'kg', 'ton']:
                            if row['Unit'] == 'g' and row['unit'] == 'kg':
                                conversion_factor = 1 / 1000
                            elif row['Unit'] == 'g' and row['unit'] == 'ton':
                                conversion_factor = 1 / 1000 / 1000 
                            elif row['Unit'] == 'ton' and row['unit'] == 'kg':
                                conversion_factor = 1 * 1000
                            elif row['Unit'] == 'kg' and row['unit'] == 'ton':
                                conversion_factor = 1 / 1000
                            elif row['Unit'] == 'kg' and row['unit'] == 'g':
                                conversion_factor = 1 * 1000
                            elif row['Unit'] == 'ton' and row['unit'] == 'g':
                                conversion_factor = 1 * 1000 * 1000
                            else:
                                conversion_factor = 1
                        else:
                            conversion_factor = 1
                    else:
                        conversion_factor = 1
                else:
                    conversion_factor = 1
                
                # 檢查數值是否為數字類型，避免類似 TypeError 的錯誤
                try:
                    quantity = float(row[quantity_column])
                    fossil_value = float(row['fossil(kg CO2-eq)_y']) 
                    biogenic_value = float(row['biogenic(kg CO2-eq)_y']) 
                    land_transformation_value = float(row['land transformation (kg CO2-eq)_y']) 
                except ValueError:
                    continue
                   
                # 計算 'fossil(kg CO2-eq)_result'
                merged_df.at[idx, 'fossil(kg CO2-eq)_result'] = quantity * fossil_value * conversion_factor
                # 計算 'biogenic(kg CO2-eq)_result'
                merged_df.at[idx, 'biogenic(kg CO2-eq)_result'] = quantity * biogenic_value * conversion_factor
                # 計算 'land transformation (kg CO2-eq)_result'
                merged_df.at[idx, 'land transformation (kg CO2-eq)_result'] = quantity * land_transformation_value * conversion_factor
            
            # 更新原始的工作表中的相關欄位，並小數點後 10 位無條件捨去
            sheet = workbook[sheet_name]
            for idx, value in enumerate(merged_df['fossil(kg CO2-eq)_result'], start=start_idx + 3):
                if pd.notna(value):
                    truncated = math.trunc(value * 10**10) / 10**10
                    sheet[f'{col_start}{idx + 1}'] = truncated
            for idx, value in enumerate(merged_df['biogenic(kg CO2-eq)_result'], start=start_idx + 3):
                if pd.notna(value):
                    truncated = math.trunc(value * 10**10) / 10**10
                    sheet[f'{chr(ord(col_start) + 1)}{idx + 1}'] = truncated
            for idx, value in enumerate(merged_df['land transformation (kg CO2-eq)_result'], start=start_idx + 3):
                if pd.notna(value):
                    truncated = math.trunc(value * 10**10) / 10**10
                    sheet[f'{chr(ord(col_start) + 2)}{idx + 1}'] = truncated

            # 假設 damage 欄位放在 fossil 欄位之後的下一欄
            num_rows = len(merged_df)
            for i in range(num_rows):
                # Excel 的列號從 1 開始，所以 row_num 需要調整
                row_num = start_idx + 3 + i + 1  
                fossil_cell = f"{col_start}{row_num}"
                biogenic_cell = f"{chr(ord(col_start) + 1)}{row_num}"
                land_cell = f"{chr(ord(col_start) + 2)}{row_num}"
                # 將 Damage Assessment 欄位設為公式
                sheet[f'{chr(ord(col_start) + 3)}{row_num}'] = f"={fossil_cell}+{biogenic_cell}+{land_cell}"
            
            self.status_callback("計算加總值並寫入每個表格的第一行...")
            # 計算加總值並寫入每個表格的第一行，並做小數點後 4 位四捨五入捨去
            fossil_total = round(merged_df['fossil(kg CO2-eq)_result'].sum(), 4)
            biogenic_total = round(merged_df['biogenic(kg CO2-eq)_result'].sum(), 4)
            land_transformation_total = round(merged_df['land transformation (kg CO2-eq)_result'].sum(), 4)

            # 計算加總值並寫入每階段的第一行
            total_fossil += fossil_total
            total_biogenic += biogenic_total
            total_land_transformation += land_transformation_total

            first_row_idx = start_idx + 1
            sheet[f'{col_start}{first_row_idx}'] = fossil_total
            sheet[f'{chr(ord(col_start) + 1)}{first_row_idx}'] = biogenic_total
            sheet[f'{chr(ord(col_start) + 2)}{first_row_idx}'] = land_transformation_total
        self.status_callback("寫入所有表格的加總值...")
        # 在每個工作表的 AB/AC/AD 欄位中寫入所有表格的加總值
        sheet[f'AB2'] = total_fossil
        sheet[f'AC2'] = total_biogenic
        sheet[f'AD2'] = total_land_transformation
        # 在每個工作表的 AE 欄位中寫入 AB/AC/AD 欄位的加總值
        sheet[f'AE2'] = total_fossil + total_biogenic + total_land_transformation

        # 返回每個工作表的加總值
        return total_fossil, total_biogenic, total_land_transformation
    
    def find_insert_positions(self, worksheet):
        """
        找出包含「◎」符號的行索引
        
        :param worksheet: xlsxwriter 工作表
        :return: 包含「◎」符號的行索引列表
        """
        insert_positions = []
        for row_num in range(1, worksheet.max_row + 1):
            for col_num in range(1, worksheet.max_column + 1):
                cell_value = worksheet.cell(row=row_num, column=col_num).value
                if cell_value and '◎' in str(cell_value):
                    insert_positions.append(row_num - 1) # 轉換成 xlsxwriter 0起始索引
                    break
        return insert_positions

    def get_format_dict(self, cell):
        """
        讀取 openpyxl cell 的字體與填充設定，並轉換為 xlsxwriter 格式字典
        """
        fmt = {}
        font = cell.font
        if font.name:
            fmt['font_name'] = font.name
        if font.sz:
            fmt['font_size'] = font.sz
        if font.bold:
            fmt['bold'] = True
        if font.italic:
            fmt['italic'] = True

        # 處理字體顏色：只接受 6 碼 hex
        if font.color and font.color.rgb:
            c = font.color.rgb
            if isinstance(c, str):
                # ARGB 轉成 RGB
                if len(c) == 8:
                    c = c[2:]
                # 檢查是否真的是 6 個 0-9A-F
                if re.fullmatch(r'[0-9A-Fa-f]{6}', c):
                    fmt['font_color'] = f'#{c}'

        # 處理背景色（solid 填充）
        fill = cell.fill
        if fill.fill_type == 'solid' and fill.fgColor and fill.fgColor.rgb:
            c = fill.fgColor.rgb
            if isinstance(c, str):
                if len(c) == 8:
                    c = c[2:]
                if re.fullmatch(r'[0-9A-Fa-f]{6}', c):
                    fmt['bg_color'] = f'#{c}'

        # 對齊
        h = cell.alignment.horizontal
        if h in ('left','center','right','fill','justify','center_across'):
            fmt['align'] = h
        v = cell.alignment.vertical
        if v in ('top','bottom','center','distributed','justify'):
            fmt['valign'] = v
        if cell.alignment.wrap_text:
            fmt['text_wrap'] = True

        # # —————— 新增：邊框設定 ——————
        # b = cell.border
        # # 判斷四邊是否同樣 style，若是就用簡單 'border'
        # styles = {b.left.style, b.right.style, b.top.style, b.bottom.style}
        # # openpyxl style 可能是 'thin','medium','dashed'…，只要非 None 就取出
        # style_map = {
        #     'thin': 1, 'medium': 2, 'thick': 4,
        #     # 如果有更多需求可擴充對應
        # }
        # # 同邊框
        # if len(styles) == 1 and None not in styles:
        #     sty = styles.pop()
        #     fmt['border'] = style_map.get(sty, 1)
        #     # 邊框顏色（若都有同色可讀出）
        #     color = b.left.color
        #     if color and color.rgb:
        #         c = color.rgb
        #         if isinstance(c, str) and len(c) in (6,8):
        #             c = c[-6:]
        #             fmt['border_color'] = f'#{c}'
        # else:
        #     # 若四邊不同，就分別處理
        #     if b.top.style:
        #         fmt['border_top'] = style_map.get(b.top.style, 1)
        #     if b.bottom.style:
        #         fmt['border_bottom'] = style_map.get(b.bottom.style, 1)
        #     if b.left.style:
        #         fmt['border_left'] = style_map.get(b.left.style, 1)
        #     if b.right.style:
        #         fmt['border_right'] = style_map.get(b.right.style, 1)
        #     # 可同時讀色
        #     if b.top.color and b.top.color.rgb:
        #         c = b.top.color.rgb[-6:]
        #         fmt['border_color'] = f'#{c}'

        return fmt

    def _get_format(self, fmt_dict, workbook):
        # 將 fmt_dict 轉成 tuple-of-tuples 作為 key（因為 dict 本身不可 hash）
        key = tuple(sorted(fmt_dict.items()))
        if key not in self._format_cache:
            self._format_cache[key] = workbook.add_format(fmt_dict)
        return self._format_cache[key]

    def transform_sheet(self):
        """
        將原始 Excel 表單轉換成盤查表單格式：
        1. 用 openpyxl 讀取模板檔案（PLCI_table_format.xlsx），取得各工作表內容。
        2. 根據模板中◎符號所在行決定插入點：
           模板中原本預留◎符號所在行及後兩列（共3列）的區塊，
           若該工作表在指定清單中，則用該工作表前四列（格式定義）替換，
           並將來源資料插入於格式定義下方。
        3. 讀取來源資料（以 pandas DataFrame 形式），統計各工作表的行數。
        4. 將來源資料插入到模板內容中，同時調整後續內容位置。
        5. 利用 xlsxwriter 將調整後的所有內容寫入新檔案中。
        """
        if not self.file_path:
            messagebox.showerror("錯誤", "請選擇 Excel 文件")
            return
        
        try:
            self._format_cache.clear()  # 清空格式快取
            self.status_callback("開始執行 Transform Sheet")
            print("開始執行 Transform Sheet")
            self.source_file_path = self.file_path
            base_dir = os.path.dirname(os.path.abspath(__file__))   # 取得目前 script 所在的資料夾/Temp檔路徑下
            target_file_path = os.path.join(base_dir, "PLCI_table_format.xlsx")
            # 用 openpyxl 讀取模板
            template_wb = openpyxl.load_workbook(target_file_path)
            print("PASS1...")
            self.status_callback("PASS1...")
            if self.update_progress_smooth:
                self.update_progress_smooth(0, 10, step=1, delay=0.02)  # 第1階段完成：10%
            # 建立格式定義字典，僅針對指定工作表
            format_definitions = {}
            for sheet_name in ['Raw Material', 'Manufacturing', 'Distribution', 'Recycling', 'Usage']:
                if sheet_name in template_wb.sheetnames:
                    sheet = template_wb[sheet_name]
                    fd = [] # 空的串列
                    # 取工作表前五列，每個儲存格都以字典形式儲存 value 與其格式設定
                    for row in sheet.iter_rows(min_row=1, max_row=5):
                        current_row = []
                        for cell in row:
                            cell_info = {
                                "format": self.get_format_dict(cell)
                            }
                            current_row.append(cell_info)
                        fd.append(current_row)
                    format_definitions[sheet_name] = fd
                    print(f"取得 {sheet_name} 的格式定義，共 {len(fd)} 列")
                else:
                    print(f"模板中找不到工作表：{sheet_name}，無法取得格式定義")
            if self.update_progress_smooth:
                self.update_progress_smooth(1, 20, step=1, delay=0.02)  # 第2階段完成：20%
            # 設定來源資料的工作表對應關係
            self.source_sheets = {
                'Raw Material': [
                    'Raw Material(Direct Material)', 
                    'Raw Material(Indirect Material)', 
                    'Raw Material(Direct Transport)', 
                    'Raw Material(Indirect Transport'
                ],
                'Manufacturing': [
                    'Manufacturing(Manufacturing)', 
                    'Manufacturing(Gas)', 
                    'Manufacturing(Electricity)', 
                    'Manufacturing(Transport)', 
                    'Manufacturing(Waste treatment)'
                ],
                'Distribution': [
                    'Distribution(Local)', 
                    'Distribution(Air)', 
                    'Distribution(Warehouse)', 
                    'Distribution(Customer)'
                ],
                'Recycling': ['Recyling(Recyling)'],
                'Usage': ['Usage']
            }
            
            # 讀取來源資料，各工作表以 DataFrame 儲存
            source_data = {}
            for target_sheet_name, source_sheet_list in self.source_sheets.items():
                for sheet_name in source_sheet_list:
                    try:
                        df = pd.read_excel(self.source_file_path, sheet_name=sheet_name)
                        # 將無限大值替換並填充空值，讓 xlsxwriter 能正確處理
                        df = df.replace([np.inf, -np.inf], 'Infinity')
                        df = df.fillna('')
                        source_data[sheet_name] = df
                        print(f"已讀取 {sheet_name} 工作表")
                        self.status_callback(f"已讀取 {sheet_name} 工作表")
                    except Exception as e:
                        print(f"警告: 無法讀取 {sheet_name} 工作表: {e}")
            print("讀取來源資料完成")
            if self.update_progress_smooth:
                self.update_progress_smooth(20, 30, step=1, delay=0.02)  # 第3階段完成：30%

            # 建立新的 xlsxwriter 工作簿
            current_datetime = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_file_name = f'merged_result_{current_datetime}.xlsx'
            new_file_path = os.path.join(self.output_dir, new_file_name)
            workbook = xlsxwriter.Workbook(new_file_path, {'nan_inf_to_errors': True})    

            if self.update_progress_smooth:
                self.update_progress_smooth(30, 80, step=1, delay=0.05) # 第4階段完成：80%
            # 處理每個目標工作表
            for target_sheet_name, source_sheet_list in self.source_sheets.items():
                print(f"處理目標工作表：{target_sheet_name}")
                worksheet = workbook.add_worksheet(target_sheet_name)
                if target_sheet_name in template_wb.sheetnames:
                    template_sheet = template_wb[target_sheet_name]
                else:
                    print(f"模板中不含 {target_sheet_name} 工作表，跳過此工作表")
                    continue
                
                template_rows = []
                for row in template_sheet.iter_rows(): 
                    current_row = []
                    for cell in row:
                        cell_info = {
                            "value": cell.value,
                            "format": self.get_format_dict(cell)  # 函式取得格式設定
                        }
                        current_row.append(cell_info)
                    template_rows.append(current_row)

                # 找出模板中含有◎符號的行索引（0-based）
                base_insert_positions = []
                for idx, row in enumerate(template_rows):
                    if any(cell is not None and '◎' in str(cell) for cell in row):
                        base_insert_positions.append(idx)
                        print(f"在 {target_sheet_name} 模板中找到插入點：第 {idx} 行")
                
                # 複製模板內容作為最終輸出，並利用 offset 追蹤因插入或替換而產生的行偏移
                new_sheet_rows = template_rows.copy()
                offset = 0


                for pos_idx, base_pos in enumerate(base_insert_positions):
                    print("開始處理插入點，pos_idx =", pos_idx)
                    self.status_callback(f"開始處理插入點，pos_idx ={pos_idx}")
                    try:
                        if pos_idx < len(source_sheet_list):
                            sheet_name = source_sheet_list[pos_idx]
                            data = source_data[sheet_name]
                            num_data_rows = data.shape[0]
                            data_rows = [list(data.iloc[i]) for i in range(num_data_rows)]
                            
                            # 原本預留◎符號所在列及其後兩列，共 3 列
                            insert_index = base_pos + 3 + offset

                            # 插入來源資料，將 data_rows 這個清單「插入」到 new_sheet_rows 的指定位置
                            new_sheet_rows[insert_index:insert_index] = data_rows   #「從第 insert_index 個元素開始，到第 insert_index 個元素結束」
                            offset += num_data_rows
                            print(f"在 {target_sheet_name} 的索引 {insert_index} 插入 {num_data_rows} 行來源資料")
                        else:
                            print(f"模板中無對應來源資料，無法插入 {sheet_name} 的資料")
                    except Exception as e:
                        print(f"無法處理 {sheet_name} 工作表：{e}")

                default_format = [cell_info["format"] for cell_info in format_definitions[target_sheet_name][4]]
                self._fallback_fmt = workbook.add_format({
                                        'border': 1,
                                        'align': 'center',
                                        'valign': 'vcenter'
                                    })

                # 將最終結果寫入 xlsxwriter 工作表
                for r, row in enumerate(new_sheet_rows):
                    for c, cell in enumerate(row):
                        # 先取出值與格式 dict
                        if isinstance(cell, dict):
                            val = cell.get("value", "")
                            fmt_dict = cell.get("format") or {}
                        else:
                            val = cell
                            fmt_dict = {}

                        if fmt_dict:
                            # 只有當 fmt_dict 裡真的有設定才建 format
                            # cell_fmt = workbook.add_format(fmt_dict, workbook)
                            cell_fmt = self._get_format(fmt_dict, workbook)
                            worksheet.write(r, c, val, cell_fmt)
                        else:
                            # 沒有自訂格式，就用 default_format
                            if isinstance(default_format, list) and c < len(default_format):
                                dfmt = default_format[c]  # 取出 column 對應的格式 dict
                                # cell_fmt = workbook.add_format(dfmt, workbook)
                                cell_fmt = self._get_format(dfmt, workbook)
                                worksheet.write(r, c, val, cell_fmt)
                            else:
                                # fallback 样式
                                cell_fmt = self._fallback_fmt
                            worksheet.write(r, c, val if val is not None else "", cell_fmt)

            if self.update_progress_smooth:
                self.update_progress_smooth(80, 95, step=1, delay=0.02) # 第5階段完成：95%
            workbook.close()
            time.sleep(0.1)
            print("靜態頁複製")
            self.status_callback("靜態頁複製")
            pythoncom.CoInitialize()
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False  # 背後跑就好
            excel.DisplayAlerts = False    # 關閉任何提示訊息
            excel.ScreenUpdating = False   # 關閉畫面更新
            # 打開新檔和範本
            if not os.path.exists(target_file_path):
                raise FileNotFoundError(f"找不到範本：{target_file_path}")
            print("============1============") 
            wb_tpl = excel.Workbooks.Open(target_file_path,
                              CorruptLoad=1,
                              ReadOnly=True,
                              IgnoreReadOnlyRecommended=True)
            wb_new = excel.Workbooks.Open(new_file_path, CorruptLoad=1)
            print("============2============")
            static_sheets = ['Instruction', 'overview', 'Process flow chart', 'simapro9.3']
            for sheet_name in static_sheets:
                try:
                    # 把範本的這張 Copy 到新檔，放在第一張動態頁前面
                    wb_tpl.Sheets(sheet_name).Copy(Before=wb_new.Sheets(1))
                except Exception as e:
                    print(f"複製「{sheet_name}」失敗：{e}")

            # 複製完所有靜態頁後，加入以下程式碼
            try:
                # 取得 wb_new 的 overview 工作表
                overview = wb_new.Sheets("overview")
                # 在 H2 設定你要的公式，例如：合計 AB2 到 AE2
                overview.Range("H2").Formula = "='Raw Material'!AE2+Manufacturing!AE2+Distribution!AE2+Recycling!AE2+Usage!AE2"
                overview.Range("V2").Formula = "=Usage!$K$5"
                # 如果要寫入本地語系公式，可改用 FormulaLocal
                # overview.Range("H2").FormulaLocal = "=SUM(AB2:AE2)"
                print("已在 overview 工作表寫入公式")
            except Exception as e:
                excel.Quit()
                print(f"設定 overview 工作表公式失敗：{e}")
                return False

            # 然後再執行 RefreshAll 並存檔
            wb_tpl.RefreshAll()
            print("靜態頁複製完成")
            self.status_callback("靜態頁複製完成")
            if self.update_progress_smooth:
                self.update_progress_smooth(95, 100, step=1, delay=0.01) # 第6階段完成：100%
            # 存檔、關檔、退出
            wb_new.Save()
            wb_tpl.Close(False)
            wb_new.Close(False)
            excel.Quit()
            # pythoncom.CoUninitialize()    

            return new_file_path
    
        except Exception as e:
            excel.Quit()
            print(f"處理 Transform Sheet 時出錯：{e}")
            messagebox.showerror("錯誤", f"處理 Transform Sheet 時出錯：{e}")
            return


    def process_all(self):
        """處理全部"""
        if not self.file_path:
            messagebox.showerror("錯誤", "請選擇 Excel 文件")
            return

        try:
            self.status_callback("開始執行 Transform Sheet")
            new_file_path = self.transform_sheet()
            if new_file_path:  # 確認返回值有效
                self.status_callback("Transform Sheet 完成，開始處理數據")
                self.process_file(file_path = new_file_path)
                self.status_callback("處理全部完成")
            return True
        
        except Exception as e:
            messagebox.showerror("錯誤", f"處理全部過程中出現錯誤：{e}")
            return False
        
    def update_excel_cache(self, result_file):
        """使用 Excel 更新公式快取值"""
        if result_file is None:
            result_file = getattr(self, "result_file", None)
        if not result_file or not os.path.exists(result_file):
            messagebox.showerror("錯誤", f"找不到檔案：{result_file}")
            return False

        excel = None
        wb = None
        try:
            pythoncom.CoInitialize()
            # 建立 Excel 應用程式實例（不顯示）
            excel = win32.DispatchEx("Excel.Application")
            excel.DisplayAlerts = False        # 不跳提示框
            wb = excel.Workbooks.Open(
                os.path.abspath(self.result_file),
                CorruptLoad=1,
                UpdateLinks=0,
                ReadOnly=False
            )           
            # 強制計算所有公式
            excel.CalculateUntilAsyncQueriesDone()
            # 儲存並關閉工作簿
            wb.Save() 
            wb.Close(SaveChanges=True)
            excel.Quit()
            pythoncom.CoUninitialize()
            return True
        
        except Exception as e:
            excel.Quit()
            print(f"更新 Excel 快取值時發生錯誤：{e}")
            messagebox.showerror("錯誤", f"更新 Excel 快取值時發生錯誤：{e}")
            return False

    def generate_report(self, template_choice):
        """
        數據處理完後產生完整報告書流程：
        1. 根據 template_choice 選擇 Word 模板
        2. 使用 self.result_file 作為數據來源，依序執行盤查表單各項函式：
            - 統整各工作表數據 (process_all_worksheets)
            - 將數據插入 Word (insert_data_to_word)
            - 生成圖表 (generate_bar_chart)
            - 針對 Raw Material、Manufacturing 等工作表進行細部處理與圖表生成
            - 前十大統整及運輸相關數據處理
        3. 最後將完整報告書存檔，檔名格式為 "智邦-產品碳足跡盤查總報告書_{today_date}.docx"
        
        """
        # 檢查是否已有數據處理過的檔案，才能進行
        if not hasattr(self, 'result_file') or not hasattr(self, 'report_file'):
            messagebox.showerror("錯誤", "請先處理檔案，再產生報告。")
            return
        # === 1. 讀取 Excel 盤查表單，並開啟 Word 模板 ===
        # 使用先前數據處理後產生的檔案名稱
        result_file = os.path.abspath(self.result_file)
        print(result_file)

        # # test code
        # result_file = r'D:\OneDrive - Accton Technology Corporation\Python\code\Excel_Vlookup_Python\結果\result_20250519_174550.xlsx'
        # template_file= r'D:\OneDrive - Accton Technology Corporation\Python\code\Excel_Vlookup_Python\智邦-產品碳足跡盤查總報告書_竹南_temp.docx'

        if result_file:
            try:
                # 在讀取前先更新公式快取值，確保公式計算後的值有被存入檔案中
                self.update_excel_cache(result_file)
            except Exception as e:
                messagebox.showerror("錯誤", f"{e}")
                return  False

        base_dir = os.path.dirname(os.path.abspath(__file__))   # 取得目前 script 所在的資料夾
        # 依據 template_choice 選擇不同模板
        if template_choice == "竹南":
            template_file = os.path.join(base_dir, "智邦-產品碳足跡盤查總報告書_竹南_temp.docx")
        elif template_choice == "竹北":
            template_file = os.path.join(base_dir, "智邦-產品碳足跡盤查總報告書_竹北_temp.docx")
        elif template_choice == "越南":
            template_file = os.path.join(base_dir, "智邦-產品碳足跡盤查總報告書_越南_temp.docx")
        else:
            messagebox.showerror("錯誤", "未知的報告模板選項")
            return
        if self.update_progress_smooth:
            self.update_progress_smooth(0, 10, step=1, delay=0.02)  # 第一階段完成：10%
        
        # 開啟選定的 Word 模板
        if not os.path.exists(template_file):
            messagebox.showerror(
                "錯誤",
                f"找不到 Word 模板檔：{template_file}"
            )
            return  False
        try:
            doc = Document(template_file)
        except Exception as e:
            print("錯誤", f"開啟 Word template 失敗：{e}")
            return  False
        
        if self.update_progress_smooth:
            self.update_progress_smooth(10, 20, step=1, delay=0.02)  # 第二階段完成：20%

        # === 2. 定義工作表名稱，讀取盤查表單存放資料至 context 清單 ===
        sheet_names = ['Raw Material', 'Manufacturing', 'Distribution', 'Usage', 'Recycling']
        transport_sheets = ['Raw Material', 'Manufacturing', 'Distribution']

        self.status_callback("讀取數據處理後產生的檔案...")
        df = pd.read_excel(result_file, sheet_name="overview") # 讀取盤查表單'overview'所需的欄位數值​

        today_date = datetime.today().strftime("%Y-%m-%d_%H%M%S")
        common_context = {'today_date': today_date,
                        'year': datetime.today().strftime("%Y"),
                        'month': datetime.today().strftime("%m")}
        if self.update_progress_smooth:
            self.update_progress_smooth(20, 30, step=1, delay=0.02)  # 第三階段完成：30%
        # 建立存放各筆資料的 context 清單
        all_contexts = []
        for _, row in df.iterrows():
            if pd.isna(row['start_date']) or pd.isna(row['end_date']):
                continue

            self.context = {
                'product_name': row['product_name'],
                'product_module': row['product_module'],
                'product_size': row['product_size'],
                'Gross_weight': row['product_weight'],
                'Net_weight': row['product_net_weight'],
                'Power': row['product_on_mode_Power'],
                'start_date': row['start_date'].strftime('%Y年%m月%d日'),
                'end_date': row['end_date'].strftime('%Y年%m月%d日'),
                # 'warranty': row['warranty'],
                'report_year': row['start_date'].strftime('%Y年'),
            }
            # 將共用參數加入每筆資料中
            self.context.update(common_context)
            all_contexts.append(self.context)
        if self.update_progress_smooth:
            self.update_progress_smooth(30, 40, step=1, delay=0.02)  # 第四階段完成：40%

        if all_contexts:
            try:
                # 建立 DocxTemplate 物件
                doc = DocxTemplate(template_file)
                # 模板中可使用 {% for item in all_contexts %} ... {% endfor %} 來逐筆列印資料
                doc.render(self.context) #使用 docxtpl 模組來套用這些資料到 Word 模板中
                full_output_name = f"智邦-產品碳足跡盤查總報告書_{today_date}.docx"   #命名output_doc
                full_output_path = os.path.join(self.output_dir, full_output_name)
                doc.save(full_output_path)
            except Exception as e:
                messagebox.showerror("錯誤", f"生成報告時發生錯誤：{e}")
                return
        else:
            messagebox.showwarning("警告", "匯入為空值，未生成 Word 文件")
            return False
        if self.update_progress_smooth:
            self.update_progress_smooth(40, 50, step=1, delay=0.02)  # 第五階段完成：50%

        # === 3. 以盤查表單作為基底，繼續處理其數據與圖表，生成完整報告書 ===
        # 呼叫各個盤查表單統整計算函式，將數據與圖表插入報告中
        self.status_callback("呼叫各個盤查表單統整計算函式，將數據與圖表生成...")
        all_results = self.process_all_worksheets(result_file, sheet_names)
        self.insert_data_to_word(all_results, sheet_names)
        self.generate_bar_chart(doc, all_results, sheet_names)
        if self.update_progress_smooth:
            self.update_progress_smooth(50, 60, step=1, delay=0.02)  # 第六階段完成：60%
        
        # Raw Material 處理與圖表生成
        self.status_callback("Raw Material 處理與圖表生成...")
        resulall_data_1, Raw_data = self.process_worksheet(result_file, 'Raw Material')
        self.process_insert_raw_data(result_file)
        self.generate_insert_raw_charts(doc, Raw_data)
        if self.update_progress_smooth:
            self.update_progress_smooth(60, 70, step=1, delay=0.02)  # 第七階段完成：70%
        
        # Manufacturing 處理與圖表生成
        print("Manufacturing 處理與圖表生成...")
        resulall_data_2, Manu_data = self.process_worksheet(result_file, 'Manufacturing')
        self.process_insert_manufacturing_data(result_file)
        self.generate_insert_manufacturing_charts(doc, Manu_data)
        self.generate_and_insert_electric_chart(doc, resulall_data_2)
        if self.update_progress_smooth:
            self.update_progress_smooth(70, 80, step=1, delay=0.02)  # 第八階段完成：80%
        
        # 前十大統整處理
        self.status_callback("前十大統整處理與圖表生成...")
        self.process_top10_data(sheet_names, result_file, doc)
        if self.update_progress_smooth:
            self.update_progress_smooth(80, 95, step=1, delay=0.02)  # 第十階段處理完畢前：95%
        
        # 運輸相關數據處理
        self.status_callback("運輸相關數據處理與圖表生成...")
        Air_all_data = self.process_transport_data(result_file, transport_sheets)
        self.analyze_and_chart_generate(Air_all_data, doc)

        # 將儲存在 self.context  的數據 & 圖表匯入
        self.status_callback("所有數據與圖表匯入報告書...")
        doc.render(self.context)    

        # === 4. 存檔完整報告書 ===
        self.status_callback("保存文件...")
        full_report_file = os.path.join(
            self.output_dir, f"智邦-產品碳足跡盤查總報告書_{today_date}.docx")
        doc.save(full_report_file)
        if self.update_progress_smooth:
            self.update_progress_smooth(95, 100, step=1, delay=0.02)  # 完全完成：100%
        print(f"【Finished】報告書匯入已完成_產品碳足跡盤查總報告書_{today_date}")

        return full_report_file

    def process_worksheet(self, file_name, sheet_name):
        """處理單個表單的數據，返回結果字典和整合數據框。"""
        df = pd.read_excel(file_name, sheet_name=sheet_name)  
        group_starts = df.index[df.iloc[:, 1].str.contains('^◎', na=False)].tolist()
        # 初始化一个空的字典，用于存储每个数据组的结果
        resulall_data = {}
        all_data = pd.DataFrame()    
        # 循环处理每个数据群组
        for j in range(len(group_starts)):
            start_idx = group_starts[j]
            end_idx = group_starts[j + 1] if j < len(group_starts) - 1 else df.shape[0]

            # 使用切片选择每个数据群组的数据
            group_data = df.iloc[start_idx:end_idx, :]

            # 删除第一列和第二列的无效数据，并将第三列作为列标题
            group_data = group_data.iloc[2:, 1:]
            group_data.columns = group_data.iloc[0, :]
            group_data = group_data.iloc[1:, :]

            num_cols = [
            'fossil(kg CO2-eq)',
            'biogenic(kg CO2-eq)',
            'land transformation (kg CO2-eq)',
            'Damage Assessment'
            ]
            # 1) 型別轉換與空值補 0
            for c in num_cols:
                group_data[c] = group_data[c].astype(float).fillna(0)
            # 2) 過濾：只保留「至少一個數值欄位非 0」的列
            mask = group_data[num_cols].sum(axis=1) != 0
            group_data = group_data.loc[mask]
            # 3) 把 Name 欄原本的空值 (NaN) 補成一個自訂標籤
            group_data['Name'] = group_data['Name'].fillna('空白群組')

            # 處理 name of database 欄位，将不同的值合并为一个字符串，使用分号分隔
            grouped_c = group_data.groupby('Name')['name of database'].apply(
                lambda x: ';'.join(sorted(set(x.dropna())))).reset_index()
            # 處理 fossil(kg CO2-eq) 欄位，将它们加总
            fossil_values = group_data.groupby('Name')['fossil(kg CO2-eq)'].sum().reset_index()
            # 處理 biogenic(kg CO2-eq) 欄位，将它们加总
            biogenic_values = group_data.groupby('Name')['biogenic(kg CO2-eq)'].sum().reset_index()
            # 處理 land transformation (kg CO2-eq) 欄位，将它们加总
            land_values = group_data.groupby('Name')['land transformation (kg CO2-eq)'].sum().reset_index()
            # 處理 Damage Assessment 欄位，将它们加总
            summed_values = group_data.groupby('Name')['Damage Assessment'].sum().reset_index()
            
            # 合并 grouped_c, fossil_values, biogenic_values, land_values, summed_values，以 'Name' 为键
            data_frames = [grouped_c, fossil_values, biogenic_values, land_values, summed_values]
            print(sheet_name, data_frames)
            merged_data = reduce(lambda left,right: pd.merge(left, right, on='Name', how='outer'), data_frames)
            merged_data = merged_data.sort_values(by='Damage Assessment', ascending=False)
            # 依據 'Damage Assessment' 列的數值大小降序排序
            
            # 将每个数据群组的结果添加到字典中
            resulall_data[f'G{j + 1}'] = merged_data
            # 將每個資料群組的整合數據添加到 all_data 中
            all_data = pd.concat([all_data, merged_data], axis=0)
            
        all_data = all_data.sort_values(by='Damage Assessment', ascending=False)    
        return resulall_data, all_data

    def process_all_worksheets(self, file_name, sheet_names):
        """處理多個表單的數據，返回所有結果。"""
        all_results = {}
        for sheet in sheet_names:
            resulall_data, all_data = self.process_worksheet(file_name, sheet)
            all_results[sheet] = {'resulall_data': resulall_data, 'all_data': all_data}
            # print(all_results)
        return all_results
        
    def insert_data_to_word(self, all_results, sheet_names):
        """
        將數據插入 Word 文件中的指定標籤位置。

        Parameters:
        - doc: Document，Word 文件對象。
        - data_mapping: dict，標籤與數據的對應字典，例如 {'[TAG_1]': 'value1', '[TAG_2]': 'value2'}。
        """
        print("【Process_2】開始將數據匯入 Word 文件")
        # 遍歷文檔中的所有段落，尋找標籤
        total_damage_assessment = 0
        sum_fossil = 0
        sum_biogenic = 0
        sum_land = 0
        for sheet in all_results.keys():
            total_damage_assessment += all_results[sheet]['all_data']['Damage Assessment'].sum()
            sum_fossil += all_results[sheet]['all_data']['fossil(kg CO2-eq)'].sum()
            sum_biogenic += all_results[sheet]['all_data']['biogenic(kg CO2-eq)'].sum()
            sum_land += all_results[sheet]['all_data']['land transformation (kg CO2-eq)'].sum()
         # 將碳排五階段統整的數值儲存至self.context
        sum_list = [] 
        for sheet in sheet_names:
            sheet_key = re.sub(r'\W+', '_', sheet).strip('_')
            df = all_results[sheet]['all_data']
            fossil      = df['fossil(kg CO2-eq)'].sum()
            biogenic    = df['biogenic(kg CO2-eq)'].sum()
            land        = df['land transformation (kg CO2-eq)'].sum()
            sum    = df['Damage Assessment'].sum()
            percentage  = sum / total_damage_assessment * 100
            names = (
                df['name of database']
                  .dropna()
                  .astype(str)
                  .unique()
                  .tolist()
            )
            sum_list.append((sheet_key, sum, names))

            self.context[f'{sheet_key}_fossil']           = round(fossil, 4)
            self.context[f'{sheet_key}_biogenic']         = round(biogenic, 4)
            self.context[f'{sheet_key}_land']             = round(land, 4)
            self.context[f'{sheet_key}_sum']              = round(sum, 4)
            self.context[f'{sheet_key}_Total_percentage'] = f"{round(percentage, 2)}%"

        sorted_sums = sorted(sum_list, key=lambda x: x[1], reverse=True)[:5]

        for idx, (sheet_key, val, names) in enumerate(sorted_sums, start=1):
            pct = val / total_damage_assessment * 100
            # 存到 self.context
            self.context[f'Carbon_percentage_{idx}'] = f"{pct:.2f}%"
            self.context[f'Carbon_stage_{idx}'] = f"{sheet_key}階段"
            self.context[f'Carbon_name_{idx}'] = ";".join(names)
        self.context['sum_percentage_1'] = f"{round(sum_fossil    / total_damage_assessment * 100, 2)}%"
        self.context['sum_percentage_2'] = f"{round(sum_biogenic  / total_damage_assessment * 100, 2)}%"
        self.context['sum_percentage_3'] = f"{round(sum_land      / total_damage_assessment * 100, 2)}%"
        self.context['Total'] = f"{round(total_damage_assessment,4)} kg CO2e"

        print("【Process_2】已匯入全階段統計表格數值") 

        # 初始化用於儲存 [Total_percentage_i] 的DataFrame
        total_percentage_df = pd.DataFrame(columns=['Sheet', 'Total_Percentage'])
        # 假設 all_results 已經被填充了數據
        for sheet in sheet_names:
            # 假設你已經有了每個工作表的 total_damage_assessment 值
            total_percentage = all_results[sheet]['all_data']['Damage Assessment'].sum() / total_damage_assessment * 100
            # 將數據添加到DataFrame中
            total_percentage_df = pd.concat([total_percentage_df, pd.DataFrame({'Sheet': [sheet], 'Total_Percentage': [total_percentage]})], ignore_index=True)
        # 根據 Total_Percentage 降序排序
        total_percentage_df.sort_values(by='Total_Percentage', ascending=False, inplace=True)
        print("total_percentage_df:", total_percentage_df)

        for j, row in total_percentage_df.iterrows():
            self.context[f'Sheet_name_{j+1}']       = row['Sheet']
            self.context[f'Total_percentage_{j+1}'] = f"{round(row['Total_Percentage'],2)}%"

    def generate_bar_chart(self, doc, all_results, sheet_names):
        """
        生成全階段的長條圖、堆疊長條圖，並將圖表插入 Word 文件對應標籤的位置。

        Parameters:
        - doc: Document
            Word 文件的 Document 物件。
        - all_results: dict
            每個表單處理結果的字典，例如：
            {
                'Raw Material': {
                    'all_data': <DataFrame>,
                    'resulall_data': <dict_of_dataframes>
                },
                'Manufacturing': {...},
                ...
            }
        - sheet_names: list
            記錄各個工作表名稱的清單，如 ['Raw Material', 'Manufacturing', 'Distribution', ...]
        """
        print("【Process_3】開始生成長條圖")
        # ------------------- 1. 計算各 Sheet 的 Damage Assessment 百分比長條圖 (bar_chart_1) -------------------
        # 先計算 total_damage_assessment
        total_damage_assessment = sum(
            all_results[sheet]['all_data']['Damage Assessment'].sum() 
            for sheet in all_results
        )
        # 创建一个颜色列表，包含前十项的颜色和一个总和项的颜色
        colors = ['#FF9D47', '#F03535', '#027671', '#0033AA', '#04DCCE', 'grey']
        # 计算每个工作表的百分比
        percentages = []
        sheet_labels = []
        for sheet in all_results:
            sheet_sum = all_results[sheet]['all_data']['Damage Assessment'].sum()
            percentage = (sheet_sum / total_damage_assessment) * 100
            percentages.append(percentage)
            sheet_labels.append(sheet)
        # 繪製 bar_chart_1
        plt.figure(figsize=(10, 6))  # 設定圖表大小
        bars = plt.bar(sheet_labels, percentages, color=colors, width=0.2)  # 創建條形圖
        # 添加 X/Y 軸標籤與標題
        plt.xlabel('Sheet Name')
        plt.ylabel('Percentage of Total Damage Assessment')
        plt.title('Percentage of Damage Assessment by Sheet')
        # 在每個 bar 上方添加數值標籤
        for i, bar in enumerate(bars):
            bar.set_label(sheet_names[i])  # 如果您想在 legend 中顯示 sheet_names[i]
            yval = bar.get_height()
            plt.text(
                bar.get_x() + bar.get_width() / 2,
                yval,
                f'{round(yval, 2)}%',
                va='bottom',
                ha='center'
            )
        plt.xticks(rotation=0)
        plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
        plt.legend(labels=list(sheet_names), loc='upper right')
        plt.tight_layout()  # 確保標籤、標題不重疊
        plt.savefig('bar_chart_1.png', bbox_inches='tight')
        # plt.show()
        # ------------------- 2. 產生各 Sheet 在 fossil/biogenic/land transformation 三項佔比 (bar_chart_2) -------------------
        categories = ['fossil(kg CO2-eq)', 'biogenic(kg CO2-eq)', 'land transformation (kg CO2-eq)']
        category_data = {category: [] for category in categories}
        sheet_labels = list(all_results.keys())  # 重新整理 labels

        # 計算 percentage
        for category in categories:
            for sheet in sheet_labels:
                category_value = all_results[sheet]['all_data'][category].sum()
                percentage = (category_value / total_damage_assessment) * 100 if total_damage_assessment > 0 else 0
                category_data[category].append(percentage)

        bar_width = 0.15  # 每個 bar 的寬度
        #category_spacing = 0.8  # 类别间的额外空间
        index = np.arange(len(categories))  # X 軸位置

        plt.figure(figsize=(10, 6))
        bars_all = []  # 用於存放所有條形的物件引用

        for i, sheet in enumerate(sheet_labels):
            bar_positions = index + i * bar_width
            bar = plt.bar(
                bar_positions,
                [category_data[cat][i] for cat in categories],
                bar_width,
                label=sheet,
                color=colors[i % len(colors)]
            )
            bars_all.append(bar)

        # 為每個 bar 添加數值標籤
        for bar_group in bars_all:
            for bar in bar_group:
                height = bar.get_height()
                plt.text(
                    bar.get_x() + bar.get_width() / 2,
                    height,
                    f'{height:.2f}%',
                    ha='center',
                    va='bottom'
                )
        # 添加图表元素
        plt.xlabel('Category')
        plt.ylabel('Percentage')
        plt.title('Values by Category and Sheet')
        plt.xticks(index + bar_width * len(sheet_labels) / 2, categories)
        plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
        plt.legend(title='Sheet Name')
        plt.tight_layout()
        plt.savefig('bar_chart_2.png', bbox_inches='tight')
        # plt.show()

        # ------------------- 3. 將繪製好的圖儲存至self.context -------------------
        chart_1 = InlineImage(doc,
                        'bar_chart_1.png',
                        width=Inches(5.83),
                        height=Inches(3.81))
        chart_2 = InlineImage(doc,
                        'bar_chart_2.png',
                        width=Inches(5.83),
                        height=Inches(3.81))

        self.context['Chart_1'] = chart_1
        self.context['Chart_2'] = chart_2

    def process_insert_raw_data(self, file_name):
        """
        讀取、統整 'Raw Material' 工作表，並將統整結果插入至 Word 的對應標籤位置。
        
        Parameters
        ----------
        file_name : str
            Excel 檔案名稱 (如 'P_lci表單_tset.xlsx')。
        doc : docx.document.Document
            已讀取的 Word 檔案 Document 物件。
        
        Returns
        -------
        Raw_data : pandas.DataFrame
            統整後的 Raw Material 資料表。(即 all_data)
        """
        print("【Process_4】開始處理原材料數據")
        # (A) 改用通用的 process_worksheet
        # resulall_data_1 可以保留在需要的話使用，但主要我們只需要 all_data
        resulall_data_1, Raw_data = self.process_worksheet(file_name, 'Raw Material')

        # (B) 開始將 Raw_data 插入 Word
        raw_sum = Raw_data['Damage Assessment'].sum()
        self.context['Raw_total'] = round(raw_sum, 4)

        for idx, row in Raw_data.head(10).reset_index(drop=True).iterrows():
            i = idx + 1  # 1-based index
            self.context[f'Raw_Name_{i}']              = row['Name']
            self.context[f'Raw_name_of_database_{i}']  = row['name of database']
            self.context[f'Raw_Damage_Assessment_{i}'] = round(row['Damage Assessment'], 4)
            # 百分比
            pct = row['Damage Assessment'] / raw_sum * 100
            self.context[f'Raw_percentage_{i}']        = f"{round(pct, 2)}%"

        # （如果少於十筆，也可選擇把沒有資料的 key 先設成空字串）
        for i in range(len(Raw_data)+1, 11):
            self.context[f'Raw_Name_{i}']              = ""
            self.context[f'Raw_name_of_database_{i}']  = ""
            self.context[f'Raw_Damage_Assessment_{i}'] = ""
            self.context[f'Raw_percentage_{i}']        = ""

        # 將統整好的前十大Raw Material數值儲存至self.context
        remaining_val = Raw_data['Damage Assessment'][10:].sum()
        self.context['Remaining_processes_1'] = f"{remaining_val:.4f}"
        total_dmg = Raw_data['Damage Assessment'].sum()
        if total_dmg > 0:
            pct = remaining_val / total_dmg * 100
        else:
            pct = 0
        self.context['Remaining_percentage_1'] = f"{round(pct, 2)}%"
        
        
        return Raw_data

    def generate_insert_raw_charts(self, doc, Raw_data):
        """
        繪製 Raw Material 的前十大 Damage Assessment 長條圖與圓餅圖，
        並將產生的圖片插入 Word 中指定的標籤位置。

        Parameters
        ----------
        doc : docx.document.Document
            Word 文件的 Document 物件。
        Raw_data : pandas.DataFrame
            包含 'Name' 與 'Damage Assessment' 欄位的資料表。

        Returns
        -------
        None
            直接在函式內完成繪圖、儲存圖片與插入 Word 不返回任何值。
        """
        print("【Process_5】開始生成並插入原材料圖表")
        # ------------------ 1. 準備繪圖資料 ------------------
        name_values = Raw_data['Name'].head(10).fillna(0)
        damage_values = Raw_data['Damage Assessment'].head(10)

        remaining_name = 'Remaining processes'
        remaining_value = Raw_data['Damage Assessment'][10:].sum()

        # 如果剩餘值是 NaN，則改成 0
        if pd.isna(remaining_value):
            remaining_value = 0
        # ------------------ 2. 繪製長條圖 (bar_chart_3.png) ------------------
        colors = [
            '#e0e462', '#d9ed92', '#b5e48c', '#99d98c', '#76c893', 
            '#52b69a', '#34a0a4', '#168aad', '#1a759f', '#184e77', 'grey'
        ]

        plt.figure(figsize=(10, 6))
        bars = plt.bar(name_values, damage_values, color=colors[:-1])
        plt.bar(remaining_name, remaining_value, color=colors[-1])  # 顯示剩餘部分

        # 添加圖表標籤/標題
        plt.xlabel('Name')
        plt.ylabel('Damage Assessment')
        plt.title('Damage Assessment by Name')

        # 在每個 bar 上方顯示對應數值
        for i, bar in enumerate(bars):
            bar.set_label(name_values.iloc[i])
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width() / 2, yval, round(yval, 4), 
                    ha='center', va='bottom')

        # 美化與保存
        plt.xticks(rotation=90)
        plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
        plt.legend(labels=list(name_values) + [remaining_name], loc='upper right')
        plt.tight_layout()
        plt.savefig('bar_chart_3.png', bbox_inches='tight')
        # plt.show()

        # ------------------ 3. 繪製圓餅圖 (pie_chart_4.png) ------------------
        if len(name_values) < 10:
            labels = list(name_values)
            sizes = list(damage_values)
            # explode 陣列根據資料數量設定（第一塊稍微突起）
            explode = [0.01] + [0] * (len(name_values) - 1)
        else:
            labels = list(name_values) + [remaining_name]
            sizes = list(damage_values) + [remaining_value]
            explode = (0.01, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)  # 突出第一塊

        sizes = [0 if np.isnan(x) else x for x in sizes]

        # 過濾掉大小為 0 的項目，同時移除對應的 labels 與 explode
        filtered = [(lab, size, exp) for lab, size, exp in zip(labels, sizes, explode) if size != 0]
        if filtered:
            labels, sizes, explode = zip(*filtered)
            labels, sizes, explode = list(labels), list(sizes), list(explode)
        else:
            # 如果全部資料都為 0 或 NaN，可依需求處理，例如設定預設值
            labels = ['No Data']
            sizes = [1]
            explode = [0]
            
        # 檢查總和是否為 0，避免後續除法錯誤
        if sum(sizes) == 0:
            labels = ['No Data']
            sizes = [1]
            explode = [0]

        # 當沒有有效數據時，避免 `annotate()` 出錯
        only_no_data = (labels == ['No Data'])

        #繪製圓餅圖
        plt.figure(figsize=(5.83, 3.81))
        if len(sizes) == 1:
            # 只有一筆有效資料，直接用 autopct 標示圓心百分比
            wedges, texts, autotexts = plt.pie(
                sizes,
                explode=explode,
                labels=labels,
                colors=colors[:len(labels)],
                autopct=lambda pct: f"{pct:1.1f}%",
                startangle=180,
                wedgeprops={'width': 0.3, 'edgecolor': 'w', 'linewidth': 2}
            )
        else:
            wedges, texts, autotexts = plt.pie(
                sizes, 
                explode=explode, 
                colors=colors, 
                autopct='',  # 不在此使用autotext，由我們手動加上
                startangle=180,
                wedgeprops={'width': 0.3, 'edgecolor': 'w', 'linewidth': 2}
            )

        # 在每個 wedge 上加百分比標籤（帶箭頭）
        if not only_no_data:
            for i, wedge in enumerate(wedges):
                ang = (wedge.theta2 - wedge.theta1) / 2 + wedge.theta1
                x = wedge.r * 0.85 * np.cos(np.deg2rad(ang))
                y = wedge.r * 0.85 * np.sin(np.deg2rad(ang))

                percentage = f"{100 * sizes[i] / sum(sizes):1.1f}%"
                connectionstyle = f"angle,angleA=0,angleB={ang}"
                kw = dict(
                    arrowprops=dict(arrowstyle="->", connectionstyle=connectionstyle),
                    zorder=0, va="center"
                )
                plt.annotate(
                    percentage,
                    xy=(x, y),
                    xytext=(1.35 * np.sign(x), 1.4 * y),
                    textcoords='data',
                    horizontalalignment='center',
                    **kw
                )

        plt.axis('equal')  # 使圓餅圖保持為圓形
        plt.subplots_adjust(left=0.3, right=0.7)
        plt.title('Damage Assessment by Name (Pie Chart)')

        legend = plt.legend(labels, loc='upper right', bbox_to_anchor=(1.5, 1))
        if labels == ['No Data']:
            plt.title('No Data Available')  # 設定標題，避免 `tight_layout()` 崩潰
        else:
            plt.tight_layout()
        plt.savefig('pie_chart_4.png', bbox_inches='tight')
        # plt.show()


        # ------------------ 4. 將繪製好的圖儲存至self.context ------------------

        chart_3 = InlineImage(doc,
                        'bar_chart_3.png',
                        width=Inches(5.83),
                        height=Inches(3.81))
        chart_4 = InlineImage(doc,
                        'pie_chart_4.png',
                        width=Inches(5.83),
                        height=Inches(3.81))

        self.context['Chart_3'] = chart_3
        self.context['Chart_4'] = chart_4

        print("【Process_5】Raw Material已匯入至報告書")

    def process_insert_manufacturing_data(self, file_name):
        """
        使用通用的 process_worksheet 函式處理 'Manufacturing' 表單，
        並將處理結果插入 Word 文件(doc)中的指定標籤。

        Parameters
        ----------
        file_name : str
            Excel 檔案名稱 (如 'P_lci表單_tset.xlsx')。
        doc : docx.document.Document
            已讀取的 Word 文件 Document 物件。

        Returns
        -------
        resulall_data_2 : dict
            以 {'G1': <DataFrame>, 'G2': <DataFrame>, ...} 形式存放的群組資料。
        Manu_data : pandas.DataFrame
            綜合所有群組的彙整資料 (Damage Assessment 降冪排序)。
        """
        print("【Process_6】開始處理製造數據")
        # 1. 呼叫通用函式 process_worksheet
        resulall_data_2, Manu_data = self.process_worksheet(file_name, 'Manufacturing')

        # 2. 用 Manu_data 插入 Word (表格標籤)

        Manu_sum = Manu_data['Damage Assessment'].sum()
        self.context['Manufacturing_total'] = round(Manu_sum, 4)

        for idx, row in Manu_data.head(10).reset_index(drop=True).iterrows():
            i = idx + 1  # 1-based index
            self.context[f'Manufacturing_Name_{i}']              = row['Name']
            self.context[f'Manufacturing_name_of_database_{i}']  = row['name of database']
            self.context[f'Manufacturing_Damage_Assessment_{i}'] = round(row['Damage Assessment'], 4)
            # 百分比
            pct = row['Damage Assessment'] / Manu_sum * 100
            self.context[f'Manufacturing_percentage_{i}']        = f"{round(pct, 2)}%"

        # （如果少於十筆，也可選擇把沒有資料的 key 先設成空字串）
        for i in range(len(Manu_data)+1, 11):
            self.context[f'Manufacturing_Name_{i}']              = ""
            self.context[f'Manufacturing_name_of_database_{i}']  = ""
            self.context[f'Manufacturing_Damage_Assessment_{i}'] = ""
            self.context[f'Manufacturing_percentage_{i}']        = ""


        remaining_val = Manu_data['Damage Assessment'][10:].sum()
        self.context['Remaining_processes_2'] = remaining_val
        total_dmg = Manu_data['Damage Assessment'].sum()
        if total_dmg > 0:
            pct = remaining_val / total_dmg * 100
        else:
            pct = 0
        self.context['Remaining_percentage_2'] = f"{round(pct, 2)}%"

        print("【Process_6】已匯入Manufacturing表格資料")


        # 4. 回傳結果，若外部還需使用
        return resulall_data_2, Manu_data

    def generate_insert_manufacturing_charts(self, doc, Manu_data):
        """將Manufacturing的Manu_data數據繪製長條圖並匯入至Word"""
        print("【Process_7】開始生成並插入製造圖表")
        # 取得要繪製的資料，若缺值就以預設值替代
        name_values = Manu_data['Name'].head(10).fillna(0)
        damage_values = Manu_data['Damage Assessment'].head(10)

        remaining_name = 'Remaining processes'
        remaining_value = Manu_data['Damage Assessment'][10:].sum()
        # 如果剩餘值是 NaN，則改成 0
        if pd.isna(remaining_value):
            remaining_value = 0

        # 创建一个颜色列表，包含前十项的颜色和一个总和项的颜色
        colors = ['#e0e462', '#d9ed92', '#b5e48c', '#99d98c', '#76c893', '#52b69a', '#34a0a4', '#168aad', '#1a759f', '#184e77', 'grey']

        # 创建一个条形图
        plt.figure(figsize=(10, 6))  # 设置图表的大小
        bars = plt.bar(name_values, damage_values, color=colors)  # 创建条形图
        plt.bar(remaining_name, remaining_value, color='grey')  # 创建条形图
        # 添加标签和标题
        plt.xlabel('Name')  # x轴标签
        plt.ylabel('Damage Assessment')  # y轴标签
        plt.title('Damage Assessment by Name')  # 图表标题
        for i, bar in enumerate(bars):
            bar.set_label(name_values.iloc[i])
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2, yval, round(yval, 4), ha='center', va='bottom')
        # 旋转x轴标签，以避免重叠
        plt.xticks(rotation=90)
        plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
        plt.legend(labels=list(name_values) + ['Remaining processes'], loc='upper right')
        # 显示图表
        plt.tight_layout()# 调整布局，确保标签和标题不重叠
        plt.savefig('bar_chart_5.png', bbox_inches='tight')
        # plt.show()



        # 创建一个圓餅图
        if len(name_values) < 10:
            labels = list(name_values)
            sizes = list(damage_values)
            # explode 陣列根據資料數量設定（第一塊稍微突起）
            explode = [0.01] + [0] * (len(name_values) - 1)
        else:
            labels = list(name_values) + ['Remaining processes']
            sizes = list(damage_values) + [remaining_value]
            explode = [0] * len(name_values) + [0.2]  # 如果你想要突出某个块，可以设置它的值大于0
        
        sizes = [0 if np.isnan(x) else x for x in sizes]

        # 過濾掉大小為 0 的項目，同時移除對應的 labels 與 explode
        filtered = [(lab, size, exp) for lab, size, exp in zip(labels, sizes, explode) if size != 0]
        if filtered:
            labels, sizes, explode = zip(*filtered)
            labels, sizes, explode = list(labels), list(sizes), list(explode)
        else:
            # 如果全部資料都為 0 或 NaN，可依需求處理，例如設定預設值
            labels = ['No Data']
            sizes = [1]
            explode = [0]

        # 檢查總和是否為 0，避免後續除法錯誤
        if sum(sizes) == 0:
            # 如果所有值都是 0，可以給一個預設值，或跳出錯誤處理
            labels = ['No Data']
            sizes = [1]
            explode = [0]
            
        only_no_data = (labels == ['No Data']) # 當沒有有效數據時，避免 `annotate()` 出錯
        
        #繪製圓餅圖
        plt.figure(figsize=(8, 6))
        
        if len(sizes) == 1:
            # 只有一筆有效資料，直接用 autopct 標示圓心百分比
            wedges, texts, autotexts = plt.pie(
                sizes,
                explode=explode,
                labels=labels,
                colors=colors[:len(labels)],
                autopct=lambda pct: f"{pct:1.1f}%",
                startangle=180,
                wedgeprops={'width': 0.3, 'edgecolor': 'w', 'linewidth': 2}
            )
        else:
            wedges, texts, autotexts = plt.pie(
                sizes, 
                explode=explode, 
                colors=colors, 
                autopct='', 
                startangle=180, 
                wedgeprops={'width': 0.3, 'edgecolor': 'w', 'linewidth': 2}
            )
        if not only_no_data:
        # 為每個區塊添加注釋（包裝在 try/except 中以防個別失敗）
            for i, wedge in enumerate(wedges):
                ang = (wedge.theta2 - wedge.theta1) / 2 + wedge.theta1
                x = wedge.r * 0.85 * np.cos(np.deg2rad(ang))
                y = wedge.r * 0.85 * np.sin(np.deg2rad(ang))
                percentage = f"{100 * sizes[i] / sum(sizes):1.1f}%"
                connectionstyle = f"angle,angleA=0,angleB={ang}"# 设置指针样式
                kw = dict(
                    arrowprops=dict(arrowstyle="->", connectionstyle=connectionstyle),
                    zorder=0, va="center"
                )
                
                # 添加注释
                plt.annotate(
                    percentage,
                    xy=(x, y),
                    xytext=(1.35*np.sign(x), 1.4*y),
                    textcoords='data',
                    horizontalalignment='center',  # 水平居中对齐
                    **kw
                )

        plt.axis('equal')  # 使得圆饼图是正圆的
        plt.subplots_adjust(left=0.3, right=0.7)
        plt.title('Damage Assessment by Name (Pie Chart)')# 添加标题
        legend = plt.legend(labels, loc='upper right', bbox_to_anchor=(1.5, 1))# 添加图例

        # 显示图表
        if labels == ['No Data']:
            plt.title('No Data Available')  # 設定標題，避免 `tight_layout()` 崩潰
        else:
            plt.tight_layout()
        plt.savefig('pie_chart_6.png') 
        # plt.show()
        print("【Process_7】已完成製造圖表生成與插入")  
        #--------------------------6. 將繪製好的圖儲存至self.context---------------------------

        chart_5 = InlineImage(doc,
                        'bar_chart_5.png',
                        width=Inches(5.83),
                        height=Inches(3.81))
        chart_6 = InlineImage(doc,
                        'pie_chart_6.png',
                        width=Inches(5.83),
                        height=Inches(3.81))

        self.context['Chart_5'] = chart_5
        self.context['Chart_6'] = chart_6

        print("【Process_7】Manufacturing已匯入至報告書")

    def generate_and_insert_electric_chart(self, doc, resulall_data_2):
        """
        從 resulall_data_2['G3'] 取得電力數據，繪製水平方向的長條圖，並將圖片插入 Word 的 [Chart_8] 標籤處。

        Parameters
        ----------
        doc : docx.document.Document
            Word 文件的 Document 物件。
        resulall_data_2 : dict
            包含多個群組資料的字典，必須存在 'G3' 這個鍵。 
            例如: resulall_data_2['G3'] => 需要包含 'Name' 與 'Damage Assessment' 欄位的 DataFrame。
        
        Returns
        -------
        None
            直接在函式內完成繪圖並插入圖片，不回傳任何值。
        """
        print("【Process_8】開始生成並插入電力圖表")
        # 1. 取得電力資料 (G3 群組)
        if 'G3' not in resulall_data_2:
            raise KeyError("resulall_data_2 中沒有 'G3' 群組，無法繪製電力數據圖表。")

        elec_data = resulall_data_2['G3'].copy()
        elec_data = elec_data.sort_values(by='Damage Assessment', ascending=False)

        # 若需要檢查 grouped_d，可視需求加上
        grouped_d = elec_data.groupby('name of database')['Name'].apply(' ; '.join).reset_index()
        
        print("電力群組資料 (G3) 結構：")
        print(elec_data.head())  # 可自行檢查前幾筆
        
        # 2. 繪製長條圖 (bar_chart_8.png)
        name_values = elec_data['Name'].head(10).fillna(0)
        damage_values = elec_data['Damage Assessment'].head(10)
        
        colors = [
            '#e0e462', '#d9ed92', '#b5e48c', '#99d98c', '#76c893',
            '#52b69a', '#34a0a4', '#168aad', '#1a759f', '#184e77', 'grey'
        ]
        
        plt.figure(figsize=(10, 6))
        bars = plt.barh(name_values, damage_values, color=colors[:len(name_values)])
        plt.xlabel('Name')
        plt.ylabel('Damage Assessment')
        plt.title('電力碳排')

        # 在每個長條顯示數值
        for i, bar in enumerate(bars):
            val = bar.get_width()  # bar.get_width() 對應 x 軸長度(因為是 barh)
            plt.text(val, bar.get_y() + bar.get_height() / 2,
                    f"{val:.2f}",
                    va='center')

        plt.xticks(rotation=90)
        plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
        plt.legend(labels=list(name_values), loc='upper right')
        plt.tight_layout()
        plt.savefig('bar_chart_8.png', bbox_inches='tight')
        # plt.show()

        
        chart_8 = InlineImage(doc,
            'bar_chart_8.png',
            width=Inches(5.83),
            height=Inches(3.81))

        self.context['Chart_8'] = chart_8

        print("【Process_8】已完成電力圖表生成與插入")

    def process_top10_data(self, sheet_names, input_file, doc):
        """
        前十大數值統整並匯入 Word 文檔的函數。

        Parameters:
        - sheet_names: list, 所有工作表的名稱。
        - input_file: str, Excel 文件名稱。
        - doc: Document, Word 文件對象。

        Returns:
        - combined_all_data: DataFrame, 統整的前十大數值數據。
        """
        print("【Process_9】開始生成前十大數據長條圖")
        combined_all_data = pd.DataFrame()
        all_results = {}

        # 處理每個工作表數據
        for sheet in sheet_names:
            resulall_data, all_data = self.process_worksheet(input_file, sheet)
            all_results[sheet] = {'resulall_data': resulall_data, 'all_data': all_data}

        # 合併所有工作表的數據
        for sheet, data in all_results.items():
            combined_all_data = pd.concat([combined_all_data, data['all_data']], axis=0)

        # 按照 'Damage Assessment' 列進行排序
        combined_all_data = combined_all_data.sort_values(by='Damage Assessment', ascending=False)
        print(combined_all_data.head())

        # 匯入前十大數據到 Word 表格
        self.insert_top10_to_word(combined_all_data)

        # 繪製長條圖並匯入 Word
        self.top10_bar_chart(combined_all_data, doc)

        return combined_all_data

    def insert_top10_to_word(self, combined_all_data):
        """
        將前十大數值匯入到 Word 文檔中的表格和段落。

        Parameters:
        - doc: Document, Word 文件對象。
        - combined_all_data: DataFrame, 統整的前十大數據。
        - all_results: dict, 全階段的處理數據。
        """
        print("【Process_10】開始將前十大數據匯入 Word 文件")
        
        # 1. 前十大 Name, name_of_database, Damage_Assessment 與 percentage
        for j in range(1, 11):
            idx = j - 1
            row = combined_all_data.iloc[idx]
            # 名稱
            self.context[f"Top10_Name_{j}"] = row["Name"]
            # 對應的 database 字串
            self.context[f"Top10_name_of_database_{j}"] = row["name of database"]
            # Damage Assessment 四位小數
            self.context[f"Top10_Damage_Assessment_{j}"] = f"{row['Damage Assessment']:.4f}"
            # 百分比：該筆 / 總和 *100，保留兩位小數
            pct = row["Damage Assessment"] / combined_all_data["Damage Assessment"].sum() * 100
            self.context[f"Top10_percentage_{j}"] = f"{pct:.2f}%"


        # 2. 剩餘製程合計與百分比（從第 11 筆開始到最後）
        remaining_sum = combined_all_data["Damage Assessment"].iloc[10:].sum()
        remaining_pct = remaining_sum / combined_all_data["Damage Assessment"].sum() * 100
        self.context["Remaining_processes_3"]   = f"{remaining_sum:.4f}"
        self.context["Remaining_percentage_3"] = f"{remaining_pct:.2f}%"

        print("【Process_10】已匯入前十大統計表格數值")

    def top10_bar_chart(self, combined_all_data, doc):
        """
        繪製前十大數據的長條圖並匯入 Word 文檔。

        Parameters:
        - combined_all_data: DataFrame, 統整的前十大數據。
        - doc: Document, Word 文件對象。
        """
        print("【Process_11】開始生成前十大數據長條圖")
        name_values = combined_all_data['Name'].head(10)
        damage_values = combined_all_data['Damage Assessment'].head(10)

        remaining_name = 'Remaining processes'
        remaining_value = combined_all_data['Damage Assessment'][10:].sum()

        colors = ['#e0e462', '#d9ed92', '#b5e48c', '#99d98c', '#76c893', '#52b69a', '#34a0a4', '#168aad', '#1a759f', '#184e77', 'grey']

        plt.figure(figsize=(10, 6))
        bars = plt.bar(name_values, damage_values, color=colors[:len(name_values)])
        plt.bar(remaining_name, remaining_value, color='grey')

        plt.xlabel('Name')
        plt.ylabel('Damage Assessment')
        plt.title('Damage Assessment by Name')
        for i, bar in enumerate(bars):
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width() / 2, yval, round(yval, 4), ha='center', va='bottom')

        plt.xticks(rotation=90)
        plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
        plt.legend(labels=list(name_values) + ['Remaining processes'], loc='upper right')
        plt.tight_layout()

        bar_chart_path = 'bar_chart_7.png'
        plt.savefig(bar_chart_path, bbox_inches='tight')
        plt.close()
        # 將繪製好的圖儲存至self.context
        chart_7 = InlineImage(doc,
                        'bar_chart_7.png',
                        width=Inches(5.83),
                        height=Inches(3.81))

        self.context['Chart_7'] = chart_7

        print("【Process_11】已完成前十大數據長條圖生成")

    def process_transport_data(self, file_name, transport_sheets):
        """
        處理運輸相關的數據，整合多個工作表並進行分析。

        Parameters:
            file_name (str): Excel 檔案名稱。
            transport_sheets (list): 包含工作表名稱的列表。

        Returns:
            dict: 每個工作表的分組結果。
            DataFrame: 合併後的所有數據。
        """
        print("【Process_12】開始處理運輸數據")
        # transport_all_results = {}
        Air_all_data = pd.DataFrame()

        for sheet_name in transport_sheets:
            sheet_df = pd.read_excel(file_name, sheet_name=sheet_name)
            group_starts = sheet_df.index[sheet_df.iloc[:, 1].str.contains('^◎', na=False)].tolist()
            # resulall_data_3 = {}

            for j in range(len(group_starts)):
                start_idx = group_starts[j]
                end_idx = group_starts[j + 1] if j < len(group_starts) - 1 else sheet_df.shape[0]
                sub_df = sheet_df.iloc[start_idx:end_idx, :]

                # 清理數據
                sub_df = sub_df.iloc[2:, 1:]
                sub_df.columns = sub_df.iloc[0, :]
                sub_df = sub_df.iloc[1:, :]
                
                if 'type of transport' not in sub_df.columns:
                    continue
                df_air = sub_df[sub_df['type of transport'].isin(['Air','AIR'])]
                if df_air.empty:
                    continue    

                # 分組和統計
                transport_grouped = df_air.groupby(['type of transport', 'Name'])
                summed_values = transport_grouped['Damage Assessment'].sum().reset_index(name='Damage Assessment')
                dbnames = transport_grouped['name of database'] \
                        .agg(lambda x: ';'.join(sorted(set(x.dropna())))) \
                        .reset_index(name='name of database')
                # data_frames = [grouped_c, fossil_values, biogenic_values, land_values, damage_values, database_names]
                merged = pd.merge(summed_values, dbnames,
                            on=['type of transport','Name'],
                            how='outer')
                merged= merged.sort_values(by='Damage Assessment', ascending=False)
                # resulall_data_3[f'G{j + 1}'] = merged
                Air_all_data = pd.concat([Air_all_data, merged], axis=0)

        Air_all_data = Air_all_data.sort_values(by='Damage Assessment', ascending=False)

        # ---- 新增：清空以前的 Air_* keys（若有的話） ----
        for k in list(self.context):
            if k.startswith('Air_'):
                del self.context[k]
        # ---- 新增：把 merged 的每一列放到 self.context  ----
        total = Air_all_data['Damage Assessment'].sum()
        for idx, row in enumerate(Air_all_data.itertuples(index=False), start=1):
            # row.Name, row._3 (對應 name of database), row._2（Damage Assessment）依實際欄位順序與屬性名稱調整
            self.context[f'Air_Name_{idx}']              = row.Name
            self.context[f'Air_name_of_database_{idx}']  = row._3
            self.context[f'Air_Damage_Assessment_{idx}'] = round(row._2, 4)
            # 百分比四捨五入到小數點 2 位
            pct = (row._2 / total * 100) if total else 0
            self.context[f'Air_percentage_{idx}']        = f"{pct:.2f}%"

        # 2. 剩餘製程合計與百分比（從第 11 筆開始到最後）
        remaining_sum = Air_all_data["Damage Assessment"].iloc[10:].sum()
        # remaining_pct = remaining_sum / Air_all_data["Damage Assessment"].sum() * 100

        if total == 0:
            remaining_pct = 0.0
        else:
            remaining_pct = remaining_sum / total * 100

        self.context["Remaining_processes_4"]   = f"{remaining_sum:.4f}"
        self.context["Remaining_percentage_4"] = f"{remaining_pct:.2f}%"

        print("【Process_12】已完成運輸數據處理")
        return Air_all_data

    def analyze_and_chart_generate(self, Air_all_data, doc):
        """
        分析合併後的運輸數據，生成報告並插入 Word 文件。

        Parameters:
            transport_all_results (dict): 每個工作表的數據結果。
            doc (Document): Word 文件對象。
            output_image (str): 長條圖保存的文件名。
            output_doc (str): Word 文件保存的文件名。

        Returns:
            None
        """
        print("【Process_13】開始生成運輸相關圖表並插入 Word 文件")
        # 分析運輸數據
        name_values = Air_all_data['Name'].head(10)
        damage_values = Air_all_data['Damage Assessment'].head(10)

        remaining_name = 'Remaining processes'
        remaining_value = Air_all_data['Damage Assessment'][10:].sum()

        if Air_all_data.empty:
            print("No air transport data available.")
            return

        # 生成長條圖
        name_values = Air_all_data['Name'].head(10).fillna(0)
        damage_values = Air_all_data['Damage Assessment'].head(10)

        remaining_name = 'Remaining processes'
        remaining_value = Air_all_data['Damage Assessment'][10:].sum()

        # 如果剩餘值是 NaN，則改成 0
        if pd.isna(remaining_value):
            remaining_value = 0

        # 创建一个颜色列表，包含前十项的颜色和一个总和项的颜色
        colors = [
            '#e0e462', '#d9ed92', '#b5e48c', '#99d98c', '#76c893', 
            '#52b69a', '#34a0a4', '#168aad', '#1a759f', '#184e77', 'grey'
        ]

        # 创建一个条形图
        plt.figure(figsize=(10, 6))  # 设置图表的大小
        bars = plt.bar(name_values, damage_values, color=colors)  # 创建条形图
        plt.bar(remaining_name, remaining_value, color='grey')  # 创建条形图

        # 添加标签和标题
        plt.xlabel('Name')  # x轴标签
        plt.ylabel('Damage Assessment')  # y轴标签
        plt.title('運輸碳排')  # 图表标题

        for i, bar in enumerate(bars):
            bar.set_label(name_values.iloc[i])
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width() / 2, yval, round(yval, 4), ha='center', va='bottom')

        # 旋转x轴标签，以避免重叠
        plt.xticks(rotation=45)
        plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
        plt.legend(labels=list(name_values) + ['Remaining processes'], loc='upper right')

        # 显示图表
        plt.tight_layout()  # 调整布局，确保标签和标题不重叠
        plt.savefig('bar_chart_9.png', bbox_inches='tight')
        # plt.show()
        # 將繪製好的圖儲存至self.context
        chart_9 = InlineImage(doc,
                        'bar_chart_9.png',
                        width=Inches(5.83),
                        height=Inches(3.81))

        self.context['Chart_9'] = chart_9

        print("【Process_13】已完成運輸相關圖表生成與插入")

    def update_progress_smooth(self, start, end, step=1, delay=0.05):
        """
        從 start 到 end 平滑更新進度，
        每次增加 step，延遲 delay 秒（單位秒）。
        """
        if self.progress_callback:
            # 確保整數更新
            for value in range(start, end + 1, step):
                self.progress_callback(value)
                time.sleep(delay)

