from excel_processing import ExcelApp
from tkcalendar import DateEntry
from tkinter import filedialog, messagebox, ttk
from tkinter import ttk
import excel_processing
import importlib
import openpyxl
import os
import pythoncom
import sys
import threading
import time
import tkinter as tk
import win32com.client as win32

importlib.reload(excel_processing) # 調用 excel_processing 模組

class ProgressBarWindow:
    def __init__(self, master, maximum=100):
        self.excel = ExcelApp()
        self.top = tk.Toplevel(master)
        self.top.title("處理進度")
        icon_path = os.path.join(sys._MEIPASS, '7106320_graph_infographic_data_element_icon.ico')
        self.top.iconbitmap(icon_path)
        # 新增 Label 顯示「LOADING（點點點）」
        self.Loading_label = tk.Label(self.top, text="LOADING．", font=("Arial", 12))
        self.Loading_label.pack(padx=20, pady=10)
        # 用一個變數記錄目前有幾個「．」，初始為 1 個
        self._loading_dot_count = 1
        # 啟動動畫
        self._animate_loading()

        # 先建立一個 frame，裡面水平排列進度條和百分比 Label
        bar_frame = tk.Frame(self.top)
        bar_frame.pack(padx=20, pady=(0, 10), fill="x")  # fill="x" 讓 frame 撐滿寬度
        # 新增 Progressbar 顯示進度條
        self.progress = ttk.Progressbar(bar_frame, orient="horizontal", length=400, mode="determinate")
        self.progress["maximum"] = maximum
        self.progress.pack(side=tk.LEFT, padx=(10, 0), fill="x", expand=True)
        # 新增 Label 顯示百分比
        self.progress_label = tk.Label(bar_frame, text="0%", font=("Arial", 14, "bold"))
        self.progress_label.pack(side=tk.LEFT, padx=(10, 0)) 

        # 新增一個 Label 顯示執行秒數
        self.elapsed_label = tk.Label(self.top, text="已執行：0 秒")
        self.elapsed_label.pack(padx=20, pady=10)
        self.start_time = time.time()
        self.update_elapsed_time()  # 每秒更新一次

    def update_elapsed_time(self):
        elapsed = time.time() - self.start_time
        minutes = int(elapsed // 60)
        seconds = elapsed - minutes * 60
        self.elapsed_label.config(text=f"已執行：{minutes}m{seconds:.1f}s")
        # 如果進度還沒到 100%，才繼續排下一次更新
        if self.progress["value"] < self.progress["maximum"]:
            self.top.after(1000, self.update_elapsed_time)

    def update_progress(self, value):
        try:
            # 利用 after() 安排在主線程更新進度條
            self.top.after(0, lambda: self.progress.config(value=value))
            self.top.after(0, lambda: self.progress_label.config(text=f"{value}%"))
        except tk.TclError:
            pass
        # 當進度達到或超過 100%，自動關閉進度視窗
        # if value >= 100:
        #     self.top.after(0, self.close)
        
    def close(self):
        self.top.destroy()

    def open_transform_progress(self):
        # 建立用於 transform_sheet 的進度條視窗
        self.transform_progress_window = ProgressBarWindow(self.root, maximum=100)
        self.root.update_idletasks()

    def open_process_progress(self):
        # 建立用於 process_file 的進度條視窗
        self.process_progress_window = ProgressBarWindow(self.root, maximum=100)
        self.root.update_idletasks()
        
    def update_status(self, status):
        # 利用 after() 確保在主線程更新視窗標題
        try:
            self.top.after(0, lambda: self.top.title(status))
        except tk.TclError:
            pass

    def _animate_loading(self):
        """
        這個函式每 500ms 被呼叫一次，  
        self._loading_dot_count 會在 1→2→3→1… 之間循環，  
        然後更新 Label 文字。
        """
        # 先計算下一輪要顯示幾個．（1、2、3 循環）
        self._loading_dot_count = (self._loading_dot_count % 3) + 1

        # 產生對應數量的全形點 (U+FF0E)，或依照你原始的「．」字元
        dots = "．" * self._loading_dot_count
        new_text = f"LOADING {dots}"

        # 更新 Label 文字
        self.Loading_label.config(text=new_text)

        # 600 毫秒後再呼叫自己一次，形成無限迴圈
        self.top.after(600, self._animate_loading)


class GUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Processing GUI")
        self.root.geometry("750x500")

        self.file_path = None
        self.excel = ExcelApp(status_callback = self.update_status, progress_callback = self.update_progress)
        self.excel.progress_callback = None
        self.progress_window = None # 進度條視窗屬性
        self.enable_refresh = tk.BooleanVar(value=False)  # 新增變數控制是否執行重新整理

        # 創建 Notebook（分頁）
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill='both')

        # 創建四個分頁
        self.tab_transform = ttk.Frame(self.notebook)
        self.tab_process = ttk.Frame(self.notebook)
        self.tab_all = ttk.Frame(self.notebook)
        self.tab_report = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab_transform, text="轉換格式")
        self.notebook.add(self.tab_process, text="處理數據")
        self.notebook.add(self.tab_all, text="完整處理")
        self.notebook.add(self.tab_report, text="完整報告書生成")
        
        # 宣告三個欄位的共用變數：公司名稱、報告類型、日期
        self.company_var = tk.StringVar()
        self.start_date_var = tk.StringVar()
        self.end_date_var = tk.StringVar()

        # 初始化分頁內容
        self.create_transform_tab()
        self.create_process_tab()
        self.create_all_tab()
        self.create_report_tab()

    def create_transform_tab(self):
        frame = self.tab_transform
        ttk.Label(frame, text="選擇 Accton Excel 檔案：").grid(row=0, column=0, padx=10, pady=10)
        
        self.transform_file_entry = ttk.Entry(frame, width=50)
        self.transform_file_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Button(frame, text="瀏覽", command=self.browse_file).grid(row=0, column=2, padx=10, pady=10)
        
        # 新增三個欄位
        ttk.Label(frame, text="產品F階機種：").grid(row=1, column=0, padx=10, pady=10)
        self.product_f_entry = ttk.Entry(frame, textvariable=self.company_var, width=50)
        self.product_f_entry.grid(row=1, column=1, padx=10, pady=10)
        
        ttk.Label(frame, text="碳足跡蒐集起始時間 (YYYY/MM/DD)：").grid(row=2, column=0, padx=10, pady=10)
        self.start_date_entry = DateEntry(
            frame, 
            textvariable=self.start_date_var,    
            date_pattern='yyyy/MM/dd',   # 顯示格式
            showweeknumbers=False,       # 不顯示週次
            width=12
            )
        self.start_date_entry.delete(0, tk.END)
        self.start_date_entry.grid(row=2, column=1, sticky='w', padx=10, pady=10)
        
        ttk.Label(frame, text="碳足跡蒐集結束時間 (YYYY/MM/DD)：").grid(row=3, column=0, padx=10, pady=10)
        self.end_date_entry = DateEntry(
            frame, 
            textvariable=self.end_date_var,    
            date_pattern='yyyy/MM/dd',   # 顯示格式
            showweeknumbers=False,       # 不顯示週次
            width=12
            )
        self.end_date_entry.delete(0, tk.END)
        self.end_date_entry.grid(row=3, column=1, sticky='w', padx=10, pady=10)


        # 新增重新整理功能的勾選框
        ttk.Checkbutton(frame, 
                        text="啟用重新整理功能",
                        variable=self.enable_refresh,
                        command=self.toggle_refresh_fields
                        ).grid(row=4, column=0, columnspan=2, padx=5, pady=5)
        
        ttk.Button(frame, text="開始轉換", command=self.transform_sheet).grid(row=4, column=1, pady=10)
        self.add_status_label(frame)
        ttk.Button(frame, text="Excel ✕", command=lambda: os.system("taskkill /f /im excel.exe")).grid(row=4, column=2, padx=10, pady=10)

        self.toggle_refresh_fields()

    def create_process_tab(self):
        frame = self.tab_process
        ttk.Label(frame, text="選擇 Excel 檔案：").grid(row=0, column=0, padx=10, pady=10)
        
        self.process_file_entry = ttk.Entry(frame, width=50)
        self.process_file_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Button(frame, text="瀏覽", command=self.browse_file).grid(row=0, column=2, padx=10, pady=10)
        
        ttk.Button(frame, text="開始處理", command=self.process_file).grid(row=1, column=1, pady=10)
        self.add_status_label(frame)

    def create_all_tab(self):
        frame = self.tab_all
        ttk.Label(frame, text="選擇 Accton Excel 檔案：").grid(row=0, column=0, padx=10, pady=10)
        
        self.process_all_file_entry = ttk.Entry(frame, width=50)
        self.process_all_file_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Button(frame, text="瀏覽", command=self.browse_file).grid(row=0, column=2, padx=10, pady=10)
        
        # 新增三個欄位
        ttk.Label(frame, text="產品F階機種：").grid(row=1, column=0, padx=10, pady=10)
        self.product_f_entry = ttk.Entry(frame, textvariable=self.company_var, width=50)
        self.product_f_entry.grid(row=1, column=1, padx=10, pady=10)
        
        ttk.Label(frame, text="碳足跡蒐集起始時間 (YYYY/MM/DD)：").grid(row=2, column=0, padx=10, pady=10)
        self.start_date_entry = DateEntry(
            frame, 
            textvariable=self.start_date_var,    
            date_pattern='yyyy/MM/dd',   # 顯示格式
            showweeknumbers=False,       # 不顯示週次
            width=12
            )
        self.start_date_entry.delete(0, tk.END)
        self.start_date_entry.grid(row=2, column=1, sticky='w', padx=10, pady=10)
        
        ttk.Label(frame, text="碳足跡蒐集結束時間 (YYYY/MM/DD)：").grid(row=3, column=0, padx=10, pady=10)
        self.end_date_entry = DateEntry(
            frame, 
            textvariable=self.end_date_var,    
            date_pattern='yyyy/MM/dd',   # 顯示格式
            showweeknumbers=False,       # 不顯示週次
            width=12
            )
        self.end_date_entry.delete(0, tk.END)
        self.end_date_entry.grid(row=3, column=1, sticky='w', padx=10, pady=10)

        # 新增重新整理功能的勾選框
        ttk.Checkbutton(frame, 
                        text="啟用重新整理功能",
                        variable=self.enable_refresh,
                        command=self.toggle_refresh_fields
                        ).grid(row=4, column=0, columnspan=2, padx=5, pady=5)

        ttk.Button(frame, text="處理全部", command=self.process_all).grid(row=4, column=1, pady=10)
        self.add_status_label(frame)
        ttk.Button(frame, text="Excel ✕", command=lambda: os.system("taskkill /f /im excel.exe")).grid(row=4, column=2, padx=10, pady=10)

        self.toggle_refresh_fields()
        
    def create_report_tab(self):
        frame = self.tab_report
        # 標題
        ttk.Label(frame, text="完整報告書生成", font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=2, padx=10, pady=10)
        
        # 下拉選單標籤
        ttk.Label(frame, text="請選擇區域：").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        # 建立下拉選單
        self.report_area = ttk.Combobox(frame, values=["竹南", "竹北", "越南"], state="readonly", width=20)
        self.report_area.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        self.report_area.current(0)  # 預設選擇第一個選項
        # 生成報告的按鈕
        ttk.Button(frame, text="生成報告書", command=self.generate_report).grid(row=2, column=0, columnspan=2, pady=10)

    def add_status_label(self, frame):
        ttk.Label(frame, text="狀態：").grid(row=5, column=0, padx=10, pady=10)
        self.status_label = ttk.Label(frame, text="等待操作", font=("Arial", 10))
        self.status_label.grid(row=5, column=1, padx=10, pady=10)
    
    def toggle_refresh_fields(self):
        """根據 self.enable_refresh 是否為 True，決定欄位要不要鎖住（disabled）"""
        if self.enable_refresh.get():
            state = 'normal'
        else:
            state = 'disabled'

        # 將三個欄位整組鎖起來或解鎖
        self.product_f_entry.config(state=state)
        self.start_date_entry.config(state=state)
        self.end_date_entry.config(state=state)

    def browse_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if self.file_path:
            self.transform_file_entry.delete(0, tk.END)
            self.process_file_entry.delete(0, tk.END)
            self.process_all_file_entry.delete(0, tk.END)
            self.transform_file_entry.insert(0, self.file_path)
            self.process_file_entry.insert(0, self.file_path)
            self.process_all_file_entry.insert(0, self.file_path)
    
    def transform_sheet(self):
        if not self.file_path:
            messagebox.showerror("錯誤", "請選擇 Excel 文件")
            return
        data = [
                self.product_f_entry.get(),
                self.start_date_entry.get(),
                self.end_date_entry.get()
            ]
        
        self.open_progress_window()
        # self.root.update()
        # 將進度更新 callback 傳入主要資料處理程式
        self.excel.progress_callback = self.update_progress
        # 使用執行緒來執行長時間運算，避免 GUI 畫面凍結
        t = threading.Thread(target=self.run_transform, args=(data,), daemon=True)
        t.start()

    def process_file(self, file_path=None):
        if file_path is not None:
            self.file_path = file_path
        if not self.file_path:
            messagebox.showerror("錯誤", "請選擇 Excel 文件")
            return
        
        self.open_progress_window()
        self.root.update()
        # 將進度更新 callback 傳入主要資料處理程式
        self.excel.progress_callback = self.update_progress
        # 使用執行緒來執行長時間運算，避免 GUI 畫面凍結
        t = threading.Thread(target=self.run_process, daemon=True)
        t.start()

    def process_all(self):
        if not self.file_path:
            messagebox.showerror("錯誤", "請選擇 Excel 文件")
            return
        data = [
            self.product_f_entry.get(),
            self.start_date_entry.get(),
            self.end_date_entry.get()
        ]
        self.open_progress_window()
        self.root.update()
        # 將進度更新 callback 傳入主要資料處理程式
        self.excel.progress_callback = self.update_progress
        # 使用執行緒來執行長時間運算，避免 GUI 畫面凍結
        t = threading.Thread(target=self.run_process_all, args=(data,), daemon=True)
        t.start()

    def generate_report(self):
        # 開始完整處理前先開啟進度條視窗
        self.open_progress_window()
        self.root.update()  #更新「主執行緒」上的 UI 事件

        # 從下拉選單取得使用者選擇的區域（例如 "竹南"、"竹北"、"越南"）
        selected_area = self.report_area.get()

        # 將進度更新 callback 傳入主要資料處理程式
        self.excel.progress_callback = self.update_progress
        # 使用執行緒來執行長時間運算，避免 GUI 畫面凍結
        t = threading.Thread(target=self.run_report, args=(selected_area,), daemon=True)
        t.start()

    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()  # 立即更新顯示
        
    def update_progress(self, value):
        if self.progress_window:
            self.progress_window.update_progress(value)
    
    def open_progress_window(self):
        self.progress_window = ProgressBarWindow(self.root, maximum=100)

    def open_transform_progress(self):
        # 建立用於 Transform 進度的視窗
        self.transform_progress_window = ProgressBarWindow(self.root, maximum=100)
        self.progress_window = self.transform_progress_window  # 若你只使用一個進度條，也可以這樣設定
        self.root.update_idletasks()

    def update_transform_progress(self, value):
        # 呼叫進度條視窗的更新函式
        if self.transform_progress_window:
            self.transform_progress_window.update_progress(value)

    def open_process_progress(self):
        # 建立用於 Transform 進度的視窗
        self.process_progress_window = ProgressBarWindow(self.root, maximum=100)
        self.progress_window = self.process_progress_window  # 若你只使用一個進度條，也可以這樣設定
        self.root.update_idletasks()

    def update_process_progress(self, value):
        # 呼叫進度條視窗的更新函式
        if self.process_progress_window:
            self.process_progress_window.update_progress(value)
            self.top.after(0, lambda: self.progress_label.config(text=f"{value}%"))
    
    def run_transform(self, data):
        """新的執行緒，作為背景線程運行"""
        self.excel.status_callback = self.progress_window.update_status
        if self.enable_refresh.get(): # 如果啟用了重新整理功能
            confirm = True
            if any(val == "" for val in data):
                confirm = messagebox.askyesno("提醒", 
                            "有部分欄位資料為空，請確認是否需完整填寫？\n若繼續執行，空值將保留原資料。是否繼續？")
            # —— 偵測到空格就自動去除 —— 
            prod_f_val = self.company_var.get()
            if ' ' in prod_f_val:
                # 移除所有空格
                cleaned = prod_f_val.replace(' ', '')
                # 把清理過的字串設回去
                self.company_var.set(cleaned)
                prod_f_val = cleaned
            # 如果使用者有輸入，且長度超過13
            if prod_f_val and len(self.company_var.get()) > 13:
                messagebox.showerror("輸入錯誤", "產品F階機種欄位最多 13 碼！")
                return
            if not confirm:
                self.root.after(0, lambda: self.progress_window.close())
                return

            if not self.update_input_sheet(self.file_path):
                self.root.after(0, lambda: self.progress_window.close())
                return           

        self.excel.file_path = self.file_path
        self.check_excel_Product()

        # 呼叫主要處理流程
        result = self.excel.transform_sheet()

        if result:
            messagebox.showinfo("完成", f"轉換完成：{result}")
        else:
            self.root.after(0, lambda: self.progress_window.close())
        
        # 整個流程完成後，再關閉進度視窗
        self.root.after(0, self.progress_window.close)
    
    def run_process(self):
        """新的執行緒，作為背景線程運行"""
        self.excel.status_callback = self.progress_window.update_status
        self.excel.file_path = self.file_path
        try:
            # 呼叫主要處理流程
            result = self.excel.process_file()

            if result:
                messagebox.showinfo("完成", "數據處理成功")
            # 整個流程完成後，再關閉進度視窗
            self.root.after(0, self.progress_window.close)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"數據處理過程中出現錯誤：{e}")
            self.root.after(0, lambda: self.progress_window.close())  # 關閉進度視窗
            print(e)
            return False
        finally:
            # 不管成不成功，都關掉進度視窗
            pythoncom.CoUninitialize() 
            if self.progress_window:
                self.progress_window.close()

    def run_process_all(self, data):
        """新的執行緒，作為背景線程運行"""
        self.excel.status_callback = self.progress_window.update_status # 確保進度視窗存在，並設定狀態回呼
        if self.enable_refresh.get(): # 如果啟用了重新整理功能
            confirm = True
            if any(val == "" for val in data):
                confirm = messagebox.askyesno("提醒", 
                            "有部分欄位資料為空，請確認是否需完整填寫？\n若繼續執行，空值將保留原資料。是否繼續？")
            # —— 偵測到空格就自動去除 —— 
            prod_f_val = self.company_var.get()
            if ' ' in prod_f_val:
                # 移除所有空格
                cleaned = prod_f_val.replace(' ', '')
                # 把清理過的字串設回去
                self.company_var.set(cleaned)
                prod_f_val = cleaned
            # 如果使用者有輸入，且長度超過13
            if prod_f_val and len(self.company_var.get()) > 13:
                messagebox.showerror("輸入錯誤", "產品F階機種欄位最多 13 碼！")
                return
            
            if not confirm:
                self.root.after(0, lambda: self.progress_window.close())
                return
            
            if not self.update_input_sheet(self.file_path):
                self.root.after(0, lambda: self.progress_window.close())
                return           

        # 設定檔案路徑與檢查
        self.excel.file_path = self.file_path
        self.check_excel_Product()
        try:
            # 第一階段：呼叫 transform_sheet (新檔案產生)
            self.excel.progress_callback = self.update_progress # 指派 transform 進度回呼
            new_file_path = self.excel.transform_sheet()
            if new_file_path:  # 確認返回值有效
                self.root.after(0, lambda: self.progress_window.update_progress(0)) # 重置進度條為 0
                self.progress_window.close()
                
                # 第二階段：呼叫 process_file (以新檔案處理)
                self.open_process_progress()
                self.excel.progress_callback = self.update_progress
                self.excel.process_file(file_path = new_file_path)
                # 【Finish】 完成
                self.root.after(0, lambda: self.progress_window.close())# 整個流程完成後，再關閉進度視窗
            if not new_file_path:
                self.root.after(0, lambda: self.progress_window.close())
        except Exception as e:
            messagebox.showerror("錯誤", f"處理全部過程中出現錯誤：{e}")
            self.root.after(0, lambda: self.progress_window.close())  # 關閉進度視窗
            print(e)
        finally:
            # 不管成不成功，都關掉進度視窗
            if self.progress_window:
                self.progress_window.close()
            del wb_tpl, wb_new, excel
            pythoncom.CoUninitialize()  

    def run_report(self, selected_area):
        """新的執行緒，作為背景線程運行"""
        self.excel.status_callback = self.progress_window.update_status # 確保進度視窗存在，並設定狀態回呼
        output_doc = None
        try:
        # 呼叫 ExcelApp 的 generate_report 方法，並將選擇的區域傳入
            output_doc = self.excel.generate_report(selected_area)
            if output_doc:
                self.root.after(0, lambda: messagebox.showinfo("完整報告書生成", 
                    f"報告書生成完成，檔案為：{output_doc}"))
            if self.progress_window:
                self.progress_window.close()
        except Exception as e:
            # 方法一：用 lambda 的預設參數把 e 綁進去
            self.root.after(0, lambda err=e: messagebox.showerror(
                "報告生成錯誤",
                f"生成報告時發生錯誤：{err}"
            ))
            print("報告生成錯誤：",e)
            output_doc = None

        finally:
            # 不管成不成功，都關掉進度視窗
            if self.progress_window:
                self.progress_window.close()
            if excel.Workbooks.Count == 0:
                excel.Quit()
            del wb, excel
            pythoncom.CoUninitialize() 

        return output_doc                
    
    def update_input_sheet(self, file_path):
        """
        將 GUI 上的三個欄位數據寫入 INPUT 工作表的 B 欄
        並重新整理Excel上的連線資料庫
        """
        try:
            self.excel.status_callback("開始更新Excel資料")

            if self.excel.update_progress_smooth:
                self.excel.update_progress_smooth(0, 10, step=1, delay=0.02)  # 第1階段完成：10%
            # 從 GUI 取得欄位資料，若無資料則以空字串處理
            product = self.product_f_entry.get() or ""
            start_date = self.start_date_entry.get() or ""
            end_date = self.end_date_entry.get() or ""
            print("取得欄位資料：", product, start_date, end_date)
            self.excel.status_callback("取得欄位資料")
            if self.excel.update_progress_smooth:
                self.excel.update_progress_smooth(10, 20, step=1, delay=0.02)  # 第2階段完成：20%
            # 建立 Excel COM 物件
            pythoncom.CoInitialize() 
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = True  # 顯示 Excel 視窗
            excel.EnableEvents = False # 暫時關閉事件，防止因為儲存格變更而觸發其他連線的自動刷新
            excel.DisplayAlerts = False
            self.excel.status_callback("開始更新 INPUT 工作表")
            if self.excel.update_progress_smooth:
                self.excel.update_progress_smooth(20, 30, step=1, delay=0.02)  # 第3階段完成：30%
            # 開啟工作簿
            workbook = excel.Workbooks.Open(file_path, ReadOnly=False)
            
            # 停用支援自動刷新的連線（針對 OLEDB 連線）
            for conn in workbook.Connections:
                try:
                    conn_name = conn.Name
                    conn_type = conn.Type  # 1 = QueryTable, 2 = OLEDB 連線
                    print(f"正在處理連線: {conn_name} (類型 {conn_type})")

                    # 確保連線有 Type 屬性，並且是 QueryTable 連線 (1 代表 QueryTable)
                    if hasattr(conn, "Type") and conn.Type == 1:
                        print(f"正在處理連線: {conn.Name}")

                        # 嘗試關閉「開啟檔案時自動刷新」
                        if hasattr(conn.OLEDBConnection, "RefreshOnFileOpen"):
                            try:
                                conn.OLEDBConnection.RefreshOnFileOpen = False
                                print(f"✅ {conn.Name}: 已關閉 RefreshOnFileOpen")
                            except Exception as e1:
                                print(f"❌ {conn.Name}: 無法關閉 RefreshOnFileOpen - {e1}")

                        # 嘗試關閉「允許刷新」
                        if hasattr(conn.OLEDBConnection, "EnableRefresh"):
                            try:
                                conn.OLEDBConnection.EnableRefresh = False
                                print(f"✅ {conn.Name}: 已關閉 EnableRefresh")
                            except Exception as e2:
                                print(f"❌ {conn.Name}: 無法關閉 EnableRefresh - {e2}")

                except Exception as e:
                    print(f"⚠️ 無法處理連線 {conn.Name}: {e}")

            # 取得 "INPUT" 工作表
            ws = workbook.Worksheets("INPUT")
            if self.excel.update_progress_smooth:
                self.excel.update_progress_smooth(30, 40, step=1, delay=0.02)  # 第4階段完成：40%
            print("更新中...1")
            self.excel.status_callback("更新中...1")

            # 分別寫入資料：第一欄資料寫入 B1，第二欄寫入 B2，t_end_date_entry 無論如何都寫入 B3
            if product.strip():   #條件式設定不為空白字串才進行寫入
                ws.Cells(1, 2).Value = product
            if start_date.strip():    
                ws.Cells(2, 2).Value = start_date
            if end_date.strip():
                ws.Cells(3, 2).Value = end_date
            print("欄位值已寫入...2") 
            if self.excel.update_progress_smooth:
                self.excel.update_progress_smooth(40, 50, step=1, delay=0.02)  # 第5階段完成：50%
            self.excel.status_callback("重新刷新Excel資料...2") 
            self.safe_save_workbook(workbook) #儲存檔案(若資料還在刷新則會重新嘗試執行)
            workbook.RefreshAll()#重新刷新Excel資料
            print("重新整理中...3") 
            if self.excel.update_progress_smooth:
                self.excel.update_progress_smooth(50, 60, step=1, delay=0.02)  # 第6階段完成：60%
            self.excel.status_callback("更新中...3")
            time.sleep(30)
            max_wait = 120
            start_time = time.time()
            refresh_done = False
            while time.time() - start_time < max_wait:
                # 檢查所有 QueryTable 是否都不在刷新中
                all_done = True
                for sh in workbook.Worksheets:
                    for qt in sh.QueryTables:
                        # 若有屬性 Refreshing 並且為 True，表示尚未完成
                        if hasattr(qt, "Refreshing") and qt.Refreshing:
                            all_done = False
                            break
                    if not all_done:
                        break
                if all_done:
                    refresh_done = True
                    break
                time.sleep(1)
            if not refresh_done:
                messagebox.showwarning("警告", "部分連線刷新可能未完全完成，將進行後續操作")
            print("更新中...4")
            if self.excel.update_progress_smooth:
                self.excel.update_progress_smooth(60, 70, step=1, delay=0.02)  # 第7階段完成：70%
            self.excel.status_callback("更新中...4")
            self.safe_save_workbook(workbook)
            workbook.Close(False)
            excel.Quit() 
            print("完成更新Excel資料")
            if self.excel.update_progress_smooth:
                self.excel.update_progress_smooth(70, 100, step=1, delay=0.02)  # 第6階段完成：60%
            self.excel.status_callback("完成更新Excel資料")
            # 整個流程完成後
            self.root.after(0, lambda: self.progress_window.update_progress(0)) # 重置進度條為 0
            pythoncom.CoUninitialize() 
            return True
        except Exception as e:
            messagebox.showerror("錯誤", f"更新 Accton 表單時發生錯誤: {e}")
            print(e)
            return False

    def safe_save_workbook(self, workbook, retry_count=10, wait_time=5):
        for i in range(retry_count):
            try:
                workbook.Save()
                return True
            except Exception as e:
                if hasattr(e, 'args') and e.args and e.args[0] == -2147418111:
                    time.sleep(wait_time)
                else:
                    raise e
        return False

    def check_excel_Product(self):
        """檢查 Excel 表中 'INPUT' 工作表的 B1 是否有數值"""
        try:
            wb = openpyxl.load_workbook(self.file_path, read_only=True)
            ws = wb["INPUT"]
            cell_value = ws["B1"].value
            wb.close()
            if cell_value is None or str(cell_value).strip() == "":
                return False
            return True
        except Exception as e:
            messagebox.showerror("錯誤", f"檢查 Excel B1 時發生錯誤: {e}")
            return False

if __name__ == "__main__":
    os.system("taskkill /f /im excel.exe >nul 2>&1")        #將Excel檔案清除
    python = sys.executable #測試
    root = tk.Tk()
    app = GUI(root)
    # app = ProgressBarWindow(root, maximum=100)
    root.mainloop()