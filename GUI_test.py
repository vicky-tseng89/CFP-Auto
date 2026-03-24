from excel_processing import ExcelApp
from tkcalendar import DateEntry
from tkinter import filedialog, messagebox, ttk
from tkinter import ttk
import csv
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
#
importlib.reload(excel_processing) # 調用 excel_processing 模組

class ProgressBarWindow:
    def __init__(self, master, maximum=100, on_user_close=None):
        self.excel = ExcelApp()
        self.top = tk.Toplevel(master)
        self._after_jobs = set()
        self._closed = False
        self._on_user_close = on_user_close
        self.top.protocol("WM_DELETE_WINDOW", lambda: self.close(request_cancel=True))

        self.top.title("處理進度")

        # 新增Icon在進度條上
        # icon_path = os.path.join(sys._MEIPASS, '7106320_graph_infographic_data_element_icon.ico')
        # self.top.iconbitmap(icon_path)

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
            self.after(1000, self.update_elapsed_time)

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
        
    def after(self, ms, func, *args):
        job = self.top.after(ms, func, *args)
        self._after_jobs.add(job)
        return job

    def cancel_afters(self):
        for job in list(self._after_jobs):
            try:
                self.top.after_cancel(job)
            except Exception:
                pass
        self._after_jobs.clear()

    def close(self, request_cancel=False):
        if self._closed:
            return
        self._closed = True
        if request_cancel and callable(self._on_user_close):
            try:
                self._on_user_close()
            except Exception:
                pass
        try:
            self.cancel_afters()
        finally:
            try:
                if self.top and self.top.winfo_exists():
                    self.top.destroy()
            except Exception:
                pass

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
        self.after(600, self._animate_loading)


class GUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Processing GUI")
        self.root.geometry("900x640")

        self.file_path = None
        self.file_paths = []
        self.current_file_path = None
        self.batch_file_listbox = None
        self.batch_results = []
        self.excel = ExcelApp(status_callback = self.update_status, progress_callback = self.update_progress)
        self.excel.progress_callback = None
        self.cancel_event = threading.Event()
        self.excel.set_cancel_event(self.cancel_event)
        self._cancel_dialog_shown = False
        self.progress_window = None # 進度條視窗屬性
        self.transform_progress_window = None
        self.process_progress_window = None
        self.enable_refresh = tk.BooleanVar(value=False)  # 新增變數控制是否執行重新整理
        self.is_running = False
        self.run_buttons = []

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
        self.factory_site_var = tk.StringVar(value="竹南")

        # 初始化分頁內容
        self.create_transform_tab()
        self.create_process_tab()
        self.create_all_tab()
        self.create_report_tab()
        self.create_batch_file_panel()

    def create_transform_tab(self):
        frame = self.tab_transform
        ttk.Label(frame, text="選擇 Accton Excel 檔案：").grid(row=0, column=0, padx=10, pady=10)
        
        self.transform_file_entry = ttk.Entry(frame, width=50)
        self.transform_file_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Button(frame, text="瀏覽", command=self.browse_file).grid(row=0, column=2, padx=10, pady=10)
        
        # 新增三個欄位


        ttk.Label(frame, text="盤查廠區").grid(row=1, column=0, padx=10, pady=10)
        self.factory_site_combo = ttk.Combobox(
            frame,
            values=["","竹南廠", "竹北廠"],
            textvariable=self.factory_site_var,
            state="readonly",
            width=20
        )
        self.factory_site_combo.grid(row=1, column=1, sticky='w', padx=10, pady=10)

        ttk.Label(frame, text="產品F階機種：").grid(row=2, column=0, padx=10, pady=10)
        self.product_f_entry = ttk.Entry(frame, textvariable=self.company_var, width=50)
        self.product_f_entry.grid(row=2, column=1, padx=10, pady=10)
        
        ttk.Label(frame, text="碳足跡蒐集起始時間 (YYYY/MM/DD)：").grid(row=3, column=0, padx=10, pady=10)
        self.start_date_entry = DateEntry(
            frame, 
            textvariable=self.start_date_var,    
            date_pattern='yyyy/MM/dd',   # 顯示格式
            showweeknumbers=False,       # 不顯示週次
            width=12
            )
        self.start_date_entry.delete(0, tk.END)
        self.start_date_entry.grid(row=3, column=1, sticky='w', padx=10, pady=10)
        
        ttk.Label(frame, text="碳足跡蒐集結束時間 (YYYY/MM/DD)：").grid(row=4, column=0, padx=10, pady=10)
        self.end_date_entry = DateEntry(
            frame, 
            textvariable=self.end_date_var,    
            date_pattern='yyyy/MM/dd',   # 顯示格式
            showweeknumbers=False,       # 不顯示週次
            width=12
            )
        self.end_date_entry.delete(0, tk.END)
        self.end_date_entry.grid(row=4, column=1, sticky='w', padx=10, pady=10)


        # 新增重新整理功能的勾選框
        ttk.Checkbutton(frame, 
                        text="啟用重新整理功能",
                        variable=self.enable_refresh,
                        command=self.toggle_refresh_fields
                        ).grid(row=5, column=0, columnspan=2, padx=5, pady=5)
        
        self.transform_button = ttk.Button(frame, text="開始轉換", command=self.transform_sheet)
        self.transform_button.grid(row=5, column=1, pady=10)
        self.run_buttons.append(self.transform_button)
        self.add_status_label(frame, row=6)
        ttk.Button(frame, text="Excel ✕", command=lambda: os.system("taskkill /f /im excel.exe")).grid(row=5, column=2, padx=10, pady=10)

        self.toggle_refresh_fields()

    def create_process_tab(self):
        frame = self.tab_process
        ttk.Label(frame, text="選擇 Excel 檔案：").grid(row=0, column=0, padx=10, pady=10)
        
        self.process_file_entry = ttk.Entry(frame, width=50)
        self.process_file_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Button(frame, text="瀏覽", command=self.browse_file).grid(row=0, column=2, padx=10, pady=10)
        
        self.process_button = ttk.Button(frame, text="開始處理", command=self.process_file)
        self.process_button.grid(row=1, column=1, pady=10)
        self.run_buttons.append(self.process_button)
        self.add_status_label(frame)

    def create_all_tab(self):
        frame = self.tab_all
        ttk.Label(frame, text="選擇 Accton Excel 檔案：").grid(row=0, column=0, padx=10, pady=10)
        
        self.process_all_file_entry = ttk.Entry(frame, width=50)
        self.process_all_file_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Button(frame, text="瀏覽", command=self.browse_file).grid(row=0, column=2, padx=10, pady=10)
        
        # 新增三個欄位


        ttk.Label(frame, text="盤查廠區").grid(row=1, column=0, padx=10, pady=10)
        self.factory_site_combo = ttk.Combobox(
            frame,
            values=["","竹南廠", "竹北廠"],
            textvariable=self.factory_site_var,
            state="readonly",
            width=20
        )
        self.factory_site_combo.grid(row=1, column=1, sticky='w', padx=10, pady=10)

        ttk.Label(frame, text="產品F階機種：").grid(row=2, column=0, padx=10, pady=10)
        self.product_f_entry = ttk.Entry(frame, textvariable=self.company_var, width=50)
        self.product_f_entry.grid(row=2, column=1, padx=10, pady=10)
        
        ttk.Label(frame, text="碳足跡蒐集起始時間 (YYYY/MM/DD)：").grid(row=3, column=0, padx=10, pady=10)
        self.start_date_entry = DateEntry(
            frame, 
            textvariable=self.start_date_var,    
            date_pattern='yyyy/MM/dd',   # 顯示格式
            showweeknumbers=False,       # 不顯示週次
            width=12
            )
        self.start_date_entry.delete(0, tk.END)
        self.start_date_entry.grid(row=3, column=1, sticky='w', padx=10, pady=10)
        
        ttk.Label(frame, text="碳足跡蒐集結束時間 (YYYY/MM/DD)：").grid(row=4, column=0, padx=10, pady=10)
        self.end_date_entry = DateEntry(
            frame, 
            textvariable=self.end_date_var,    
            date_pattern='yyyy/MM/dd',   # 顯示格式
            showweeknumbers=False,       # 不顯示週次
            width=12
            )
        self.end_date_entry.delete(0, tk.END)
        self.end_date_entry.grid(row=4, column=1, sticky='w', padx=10, pady=10)

        # 新增重新整理功能的勾選框
        ttk.Checkbutton(frame, 
                        text="啟用重新整理功能",
                        variable=self.enable_refresh,
                        command=self.toggle_refresh_fields
                        ).grid(row=5, column=0, columnspan=2, padx=5, pady=5)

        self.process_all_button = ttk.Button(frame, text="處理全部", command=self.process_all)
        self.process_all_button.grid(row=5, column=1, pady=10)
        self.run_buttons.append(self.process_all_button)
        self.add_status_label(frame, row=6)
        ttk.Button(frame, text="Excel ✕", command=lambda: os.system("taskkill /f /im excel.exe")).grid(row=5, column=2, padx=10, pady=10)

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
        self.report_button = ttk.Button(frame, text="生成報告書", command=self.generate_report)
        self.report_button.grid(row=2, column=0, columnspan=2, pady=10)
        self.run_buttons.append(self.report_button)

    def create_batch_file_panel(self):
        panel = ttk.LabelFrame(self.root, text="批次匯入檔案")
        panel.pack(fill="both", expand=False, padx=10, pady=(0, 10))

        button_row = ttk.Frame(panel)
        button_row.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Button(button_row, text="重新選擇", command=self.browse_file).pack(side="left", padx=(0, 6))
        ttk.Button(button_row, text="加入檔案", command=self.append_files).pack(side="left", padx=(0, 6))
        ttk.Button(button_row, text="移除選取", command=self.remove_selected_files).pack(side="left", padx=(0, 6))
        ttk.Button(button_row, text="清空清單", command=self.clear_selected_files).pack(side="left")

        list_row = ttk.Frame(panel)
        list_row.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.batch_file_listbox = tk.Listbox(list_row, height=6, selectmode=tk.EXTENDED)
        self.batch_file_listbox.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(list_row, orient="vertical", command=self.batch_file_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.batch_file_listbox.configure(yscrollcommand=scrollbar.set)

    def _normalize_file_paths(self, paths):
        unique_paths = []
        seen = set()
        for raw_path in paths:
            path = os.path.abspath(str(raw_path))
            if path in seen:
                continue
            seen.add(path)
            unique_paths.append(path)
        return unique_paths

    def _set_selected_files(self, paths, replace=True):
        normalized_paths = self._normalize_file_paths(paths)
        if replace:
            self.file_paths = normalized_paths
        else:
            existing = list(self.file_paths)
            self.file_paths = self._normalize_file_paths(existing + normalized_paths)
        self.file_path = self.file_paths[0] if self.file_paths else None
        self.current_file_path = self.file_path
        self._refresh_file_entries()
        self._refresh_file_listbox()

    def _refresh_file_listbox(self):
        if not self.batch_file_listbox:
            return
        self.batch_file_listbox.delete(0, tk.END)
        for idx, file_path in enumerate(self.file_paths, start=1):
            self.batch_file_listbox.insert(tk.END, f"{idx:03d}. {file_path}")

    def _build_entry_text(self):
        if not self.file_paths:
            return ""
        if len(self.file_paths) == 1:
            return self.file_paths[0]
        return f"{self.file_paths[0]} (+{len(self.file_paths) - 1} files)"

    def _refresh_file_entries(self):
        entry_text = self._build_entry_text()
        for entry in (
            getattr(self, "transform_file_entry", None),
            getattr(self, "process_file_entry", None),
            getattr(self, "process_all_file_entry", None),
        ):
            if entry is None:
                continue
            entry.delete(0, tk.END)
            if entry_text:
                entry.insert(0, entry_text)

    def append_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not file_paths:
            return
        self._set_selected_files(file_paths, replace=False)

    def remove_selected_files(self):
        if not self.batch_file_listbox:
            return
        selected_indices = list(self.batch_file_listbox.curselection())
        if not selected_indices:
            return
        selected_set = set(selected_indices)
        self.file_paths = [path for idx, path in enumerate(self.file_paths) if idx not in selected_set]
        self.file_path = self.file_paths[0] if self.file_paths else None
        self.current_file_path = self.file_path
        self._refresh_file_entries()
        self._refresh_file_listbox()

    def clear_selected_files(self):
        self.file_paths = []
        self.file_path = None
        self.current_file_path = None
        self._refresh_file_entries()
        self._refresh_file_listbox()

    def _get_selected_files(self):
        return list(self.file_paths)

    def add_status_label(self, frame, row=5):
        ttk.Label(frame, text="狀態：").grid(row=row, column=0, padx=10, pady=10)
        self.status_label = ttk.Label(frame, text="等待操作", font=("Arial", 10))
        self.status_label.grid(row=row, column=1, padx=10, pady=10)
    
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
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not file_paths:
            return
        self._set_selected_files(file_paths, replace=True)
    
    def sync_factory_site(self):
        self.excel.factory_site = (self.factory_site_var.get() or "").strip()

    def transform_sheet(self):
        if not self.file_paths:
            messagebox.showerror("錯誤", "請至少匯入 1 個 Excel 文件")
            return
        if not self.begin_task():
            return
        self.reset_cancel_state()
        data = [
                self.product_f_entry.get(),
                self.start_date_entry.get(),
                self.end_date_entry.get()
            ]
        try:
            self.open_progress_window()
            # self.root.update()
            # 將進度更新 callback 傳入主要資料處理程式
            self.excel.progress_callback = self.update_progress
            # 使用執行緒來執行長時間運算，避免 GUI 畫面凍結
            t = threading.Thread(target=self.run_transform, args=(data,), daemon=True)
            t.start()
        except Exception as e:
            self.finish_task()
            self.show_error("錯誤", f"啟動轉換流程失敗：{e}")

    def process_file(self, file_path=None):
        if file_path is not None:
            self._set_selected_files([file_path], replace=True)
        if not self.file_paths:
            messagebox.showerror("錯誤", "請至少匯入 1 個 Excel 文件")
            return
        if not self.begin_task():
            return
        self.reset_cancel_state()
        try:
            self.open_progress_window()
            self.root.update()
            # 將進度更新 callback 傳入主要資料處理程式
            self.excel.progress_callback = self.update_progress
            # 使用執行緒來執行長時間運算，避免 GUI 畫面凍結
            t = threading.Thread(target=self.run_process, daemon=True)
            t.start()
        except Exception as e:
            self.finish_task()
            self.show_error("錯誤", f"啟動處理流程失敗：{e}")

    def process_all(self):
        if not self.file_paths:
            messagebox.showerror("錯誤", "請至少匯入 1 個 Excel 文件")
            return
        if not self.begin_task():
            return
        self.reset_cancel_state()
        data = [
            self.product_f_entry.get(),
            self.start_date_entry.get(),
            self.end_date_entry.get()
        ]
        try:
            self.open_progress_window()
            self.root.update()
            # 將進度更新 callback 傳入主要資料處理程式
            self.excel.progress_callback = self.update_progress
            # 使用執行緒來執行長時間運算，避免 GUI 畫面凍結
            t = threading.Thread(target=self.run_process_all, args=(data,), daemon=True)
            t.start()
        except Exception as e:
            self.finish_task()
            self.show_error("錯誤", f"啟動完整流程失敗：{e}")

    def generate_report(self):
        if not self.begin_task():
            return
        self.reset_cancel_state()
        try:
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
        except Exception as e:
            self.finish_task()
            self.show_error("錯誤", f"啟動報告流程失敗：{e}")

    def update_status(self, message):
        def _update():
            self.status_label.config(text=message)
            self.root.update_idletasks()  # 立即更新顯示
        if threading.current_thread() is threading.main_thread():
            _update()
        else:
            self.run_on_main(_update, wait=False)
        
    def update_progress(self, value):
        if self.progress_window:
            self.progress_window.update_progress(value)

    def reset_cancel_state(self):
        self.cancel_event.clear()
        self.excel.clear_cancel()
        self._cancel_dialog_shown = False

    @staticmethod
    def _cancel_message():
        return "Operation cancelled by user.\n已取消作業，未寫入任何檔案。"

    def request_cancel(self):
        if self.cancel_event.is_set():
            return
        self.cancel_event.set()
        self.excel.request_cancel()
        self.update_status("取消中...")

    def _raise_if_cancelled(self):
        if self.cancel_event.is_set():
            self.excel.request_cancel()
            self.excel.was_cancelled = True
            raise excel_processing.UserCancelledError(self._cancel_message())

    def _handle_user_cancelled(self, exc=None):
        self.excel.request_cancel()
        self.excel.was_cancelled = True
        self.update_status("作業已取消")
        if self._cancel_dialog_shown:
            return
        self._cancel_dialog_shown = True
        message = str(exc) if exc else self._cancel_message()
        self.show_error("UserCancelledError", message)

    def _make_threadsafe_status_callback(self):
        def _status(status):
            self.root.after(0, lambda: self.progress_window.update_status(status) if self.progress_window else None)
        return _status

    def run_on_main(self, func, wait=True):
        if threading.current_thread() is threading.main_thread():
            func()
            return
        if not wait:
            self.root.after(0, func)
            return
        done = threading.Event()
        def wrapper():
            try:
                func()
            finally:
                done.set()
        self.root.after(0, wrapper)
        done.wait()

    def show_info(self, title, message, wait=True, close_progress=False):
        def _show():
            if close_progress:
                self.close_progress_window(wait=True)
            messagebox.showinfo(title, message)
        self.run_on_main(_show, wait=wait)

    def show_warning(self, title, message, wait=False):
        self.run_on_main(lambda: messagebox.showwarning(title, message), wait=wait)

    def show_error(self, title, message, wait=False):
        self.run_on_main(lambda: messagebox.showerror(title, message), wait=wait)

    def _extract_process_error(self, result):
        if isinstance(result, excel_processing.TaskResult):
            return result.message
        if isinstance(result, dict):
            err = result.get("error")
            if err:
                return err
        last_err = getattr(self.excel, "last_error", None)
        if last_err:
            return last_err
        if result is None:
            return "process_file() 沒有回傳結果，請檢查 process_file 成功路徑的 return。"
        return f"非預期回傳內容：{result!r}"

    def ask_yes_no(self, title, message):
        response = {"value": False}
        def ask():
            response["value"] = messagebox.askyesno(title, message)
        self.run_on_main(ask, wait=True)
        return response["value"]

    def _set_run_buttons_state(self, state):
        for button in self.run_buttons:
            try:
                button.config(state=state)
            except Exception:
                pass

    def begin_task(self):
        if self.is_running:
            self.show_warning("提醒", "目前已有作業執行中，請等待目前作業完成。", wait=False)
            return False
        self.is_running = True
        self._set_run_buttons_state("disabled")
        return True

    def finish_task(self):
        self.is_running = False
        self._set_run_buttons_state("normal")

    def close_progress_window(self, wait=True, request_cancel=False):
        def _close():
            if self.progress_window:
                window = self.progress_window
                try:
                    window.close(request_cancel=request_cancel)
                finally:
                    if window is getattr(self, "transform_progress_window", None):
                        self.transform_progress_window = None
                    if window is getattr(self, "process_progress_window", None):
                        self.process_progress_window = None
                    self.progress_window = None
        self.run_on_main(_close, wait=wait)

    def _on_progress_window_user_close(self):
        self.request_cancel()
        self.close_progress_window(wait=False)

    def open_progress_window(self):
        self.progress_window = ProgressBarWindow(self.root, maximum=100, on_user_close=self._on_progress_window_user_close)

    def open_transform_progress(self):
        # 建立用於 Transform 進度的視窗
        self.transform_progress_window = ProgressBarWindow(self.root, maximum=100, on_user_close=self._on_progress_window_user_close)
        self.progress_window = self.transform_progress_window  # 若你只使用一個進度條，也可以這樣設定
        self.root.update_idletasks()

    def update_transform_progress(self, value):
        # 呼叫進度條視窗的更新函式
        if self.transform_progress_window:
            self.transform_progress_window.update_progress(value)

    def open_process_progress(self):
        # 建立用於 Transform 進度的視窗
        self.process_progress_window = ProgressBarWindow(self.root, maximum=100, on_user_close=self._on_progress_window_user_close)
        self.progress_window = self.process_progress_window  # 若你只使用一個進度條，也可以這樣設定
        self.root.update_idletasks()

    def update_process_progress(self, value):
        # 呼叫進度條視窗的更新函式
        if self.process_progress_window:
            self.process_progress_window.update_progress(value)
            self.process_progress_window.top.after(0, lambda: self.process_progress_window.progress_label.config(text=f"{value}%"))

    def safe_save_workbook(self, workbook, retry_count=10, wait_time=5):
        for i in range(retry_count):
            try:
                self._raise_if_cancelled()
                workbook.Save()
                return True
            except Exception as e:
                if hasattr(e, 'args') and e.args and e.args[0] == -2147418111:
                    self._raise_if_cancelled()
                    time.sleep(wait_time)
                else:
                    raise e
        return False

    def check_excel_Product(self, file_path=None):
        """檢查 Excel 表中 'INPUT' 工作表的 B1 是否有數值"""
        target_path = file_path or self.file_path
        if not target_path:
            return False
        try:
            wb = openpyxl.load_workbook(target_path, read_only=True)
            ws = wb["INPUT"]
            cell_value = ws["B1"].value
            wb.close()
            if cell_value is None or str(cell_value).strip() == "":
                return False
            return True
        except Exception as e:
            messagebox.showerror("錯誤", f"檢查 Excel B1 時發生錯誤: {e}")
            return False

    def _make_batch_status_callback(self, task_name, index, total, file_path):
        short_name = os.path.basename(file_path)

        def _status(status):
            prefix = f"[{index + 1}/{total}] {task_name} | {short_name}"
            merged = f"{prefix} | {status}" if status else prefix
            self.root.after(
                0,
                lambda: self.progress_window.update_status(merged) if self.progress_window else None,
            )
            self.root.after(0, lambda: self.update_status(merged))

        return _status

    def _make_batch_progress_callback(self, index, total, stage_start=0.0, stage_span=1.0):
        stage_start = max(0.0, min(1.0, stage_start))
        stage_span = max(0.0, min(1.0 - stage_start, stage_span))

        def _progress(value):
            try:
                value = float(value)
            except Exception:
                value = 0.0
            value = max(0.0, min(100.0, value))
            file_progress = stage_start + stage_span * (value / 100.0)
            overall = ((index + file_progress) / max(1, total)) * 100.0
            self.update_progress(int(round(overall)))

        return _progress

    def _task_result_to_record(self, task_name, file_path, result):
        artifacts = {}
        error_code = ""
        message = ""
        ok = False
        if isinstance(result, excel_processing.TaskResult):
            ok = bool(result.ok)
            error_code = result.error_code or ""
            message = result.message or ""
            artifacts = dict(result.artifacts or {})
        elif isinstance(result, dict):
            ok = bool(result.get("ok"))
            error_code = str(result.get("error_code") or "")
            if result.get("cancelled"):
                error_code = "USER_CANCELLED"
            message = str(result.get("message") or result.get("error") or "")
            artifacts = dict(result)
        else:
            message = str(result)

        if ok:
            status = "success"
        elif error_code == "USER_CANCELLED":
            status = "cancelled"
        else:
            status = "failed"

        return {
            "task": task_name,
            "input_file": file_path,
            "status": status,
            "ok": ok,
            "error_code": error_code,
            "message": message,
            "merged_file": artifacts.get("path") or artifacts.get("merged_file") or "",
            "result_file": artifacts.get("result_file") or "",
            "report_file": artifacts.get("report_file") or artifacts.get("report_doc") or "",
            "run_id": artifacts.get("run_id") or "",
        }

    def _append_skipped_records(self, records, task_name, file_paths, start_index):
        for idx in range(start_index, len(file_paths)):
            records.append(
                {
                    "task": task_name,
                    "input_file": file_paths[idx],
                    "status": "skipped",
                    "ok": False,
                    "error_code": "SKIPPED_AFTER_CANCEL",
                    "message": "Skipped because batch was cancelled.",
                    "merged_file": "",
                    "result_file": "",
                    "report_file": "",
                    "run_id": "",
                }
            )

    def _write_batch_summary_csv(self, task_name, records):
        if not records:
            return ""
        if len(records) <= 1:
            return ""
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        output_dir = getattr(self.excel, "result_dir", os.getcwd())
        os.makedirs(output_dir, exist_ok=True)
        summary_path = os.path.join(output_dir, f"batch_summary_{task_name}_{timestamp}.csv")
        columns = [
            "task",
            "input_file",
            "status",
            "ok",
            "error_code",
            "message",
            "merged_file",
            "result_file",
            "report_file",
            "run_id",
        ]
        with open(summary_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=columns)
            writer.writeheader()
            for row in records:
                writer.writerow({col: row.get(col, "") for col in columns})
        return summary_path

    def _show_batch_summary(self, task_label, records, summary_path):
        total = len(records)
        success = sum(1 for item in records if item.get("status") == "success")
        failed = sum(1 for item in records if item.get("status") == "failed")
        cancelled = sum(1 for item in records if item.get("status") == "cancelled")
        skipped = sum(1 for item in records if item.get("status") == "skipped")
        message_lines = [
            f"{task_label} batch complete.",
            f"Total: {total}",
            f"Success: {success}",
            f"Failed: {failed}",
            f"Cancelled: {cancelled}",
            f"Skipped: {skipped}",
        ]
        if summary_path:
            message_lines.append(f"Summary CSV: {summary_path}")
        if failed > 0:
            message_lines.append("")
            message_lines.append("Failed files:")
            for item in records:
                if item.get("status") != "failed":
                    continue
                short_name = os.path.basename(item.get("input_file", ""))
                reason = item.get("message") or item.get("error_code") or "Unknown error"
                message_lines.append(f"- {short_name}: {reason}")
        self.show_info("Batch Result", "\n".join(message_lines), wait=False, close_progress=True)


    # Overrides: use TaskResult and delegate COM lifecycle to engine.
    def run_transform(self, data):
        file_paths = self._get_selected_files()
        total = len(file_paths)
        records = []
        try:
            self.sync_factory_site()
            for idx, file_path in enumerate(file_paths):
                self._raise_if_cancelled()
                self.current_file_path = file_path
                self.file_path = file_path
                self.excel.file_path = file_path
                self.excel.status_callback = self._make_batch_status_callback("Transform", idx, total, file_path)
                self.excel.progress_callback = self._make_batch_progress_callback(idx, total)
                result = self.excel.transform_sheet()
                record = self._task_result_to_record("transform", file_path, result)
                records.append(record)
                self.update_progress(int(round(((idx + 1) / max(1, total)) * 100)))
                if record["status"] == "cancelled":
                    self._append_skipped_records(records, "transform", file_paths, idx + 1)
                    break
            summary_path = self._write_batch_summary_csv("transform", records)
            self._show_batch_summary("Transform", records, summary_path)
        except excel_processing.UserCancelledError as e:
            if records:
                processed = len(records)
                self._append_skipped_records(records, "transform", file_paths, processed)
                summary_path = self._write_batch_summary_csv("transform", records)
                self._show_batch_summary("Transform", records, summary_path)
            else:
                self._handle_user_cancelled(e)
        except Exception as e:
            self.show_error("Error", f"Transform execution error: {e}")
        finally:
            self.close_progress_window(wait=False)
            self.run_on_main(self.finish_task, wait=False)

    def run_process(self):
        file_paths = self._get_selected_files()
        total = len(file_paths)
        records = []
        try:
            self.sync_factory_site()
            for idx, file_path in enumerate(file_paths):
                self._raise_if_cancelled()
                self.current_file_path = file_path
                self.file_path = file_path
                self.excel.file_path = file_path
                self.excel.status_callback = self._make_batch_status_callback("Process", idx, total, file_path)
                self.excel.progress_callback = self._make_batch_progress_callback(idx, total)
                result = self.excel.process_file(file_path=file_path)
                record = self._task_result_to_record("process", file_path, result)
                records.append(record)
                self.update_progress(int(round(((idx + 1) / max(1, total)) * 100)))
                if record["status"] == "cancelled":
                    self._append_skipped_records(records, "process", file_paths, idx + 1)
                    break
            summary_path = self._write_batch_summary_csv("process", records)
            self._show_batch_summary("Process", records, summary_path)
        except excel_processing.UserCancelledError as e:
            if records:
                processed = len(records)
                self._append_skipped_records(records, "process", file_paths, processed)
                summary_path = self._write_batch_summary_csv("process", records)
                self._show_batch_summary("Process", records, summary_path)
            else:
                self._handle_user_cancelled(e)
        except Exception as e:
            self.show_error("Error", f"Processing execution error: {e}")
        finally:
            self.close_progress_window(wait=False)
            self.run_on_main(self.finish_task, wait=False)

    def run_process_all(self, data):
        file_paths = self._get_selected_files()
        total = len(file_paths)
        records = []
        try:
            self.sync_factory_site()
            for idx, file_path in enumerate(file_paths):
                self._raise_if_cancelled()
                self.current_file_path = file_path
                self.file_path = file_path
                self.excel.file_path = file_path

                if self.enable_refresh.get():
                    self.excel.status_callback = self._make_batch_status_callback("Refresh INPUT", idx, total, file_path)
                    if not self.update_input_sheet(file_path, reset_progress=False):
                        if self.excel.was_cancelled:
                            records.append(
                                {
                                    "task": "process_all",
                                    "input_file": file_path,
                                    "status": "cancelled",
                                    "ok": False,
                                    "error_code": "USER_CANCELLED",
                                    "message": "Cancelled during INPUT refresh.",
                                    "merged_file": "",
                                    "result_file": "",
                                    "report_file": "",
                                    "run_id": "",
                                }
                            )
                            self._append_skipped_records(records, "process_all", file_paths, idx + 1)
                            break
                        records.append(
                            {
                                "task": "process_all",
                                "input_file": file_path,
                                "status": "failed",
                                "ok": False,
                                "error_code": "UPDATE_INPUT_FAILED",
                                "message": getattr(self.excel, "last_error", "Update INPUT failed."),
                                "merged_file": "",
                                "result_file": "",
                                "report_file": "",
                                "run_id": "",
                            }
                        )
                        continue

                self.excel.status_callback = self._make_batch_status_callback("Transform", idx, total, file_path)
                self.excel.progress_callback = self._make_batch_progress_callback(idx, total, stage_start=0.0, stage_span=0.5)
                transform_result = self.excel.transform_sheet()
                transform_record = self._task_result_to_record("process_all_transform", file_path, transform_result)
                if transform_record["status"] != "success":
                    records.append(transform_record)
                    if transform_record["status"] == "cancelled":
                        self._append_skipped_records(records, "process_all", file_paths, idx + 1)
                        break
                    continue

                merged_file = transform_record["merged_file"]
                if not merged_file:
                    records.append(
                        {
                            "task": "process_all",
                            "input_file": file_path,
                            "status": "failed",
                            "ok": False,
                            "error_code": "MISSING_TRANSFORM_OUTPUT",
                            "message": "Transform completed but merged file path is missing.",
                            "merged_file": "",
                            "result_file": "",
                            "report_file": "",
                            "run_id": transform_record.get("run_id", ""),
                        }
                    )
                    continue
                self.excel.status_callback = self._make_batch_status_callback("Process", idx, total, file_path)
                self.excel.progress_callback = self._make_batch_progress_callback(idx, total, stage_start=0.5, stage_span=0.5)
                process_result = self.excel.process_file(file_path=merged_file)
                process_record = self._task_result_to_record("process_all", file_path, process_result)
                process_record["merged_file"] = merged_file
                records.append(process_record)
                self.update_progress(int(round(((idx + 1) / max(1, total)) * 100)))
                if process_record["status"] == "cancelled":
                    self._append_skipped_records(records, "process_all", file_paths, idx + 1)
                    break

            summary_path = self._write_batch_summary_csv("process_all", records)
            self._show_batch_summary("Process All", records, summary_path)
        except excel_processing.UserCancelledError as e:
            if records:
                processed = len(records)
                self._append_skipped_records(records, "process_all", file_paths, processed)
                summary_path = self._write_batch_summary_csv("process_all", records)
                self._show_batch_summary("Process All", records, summary_path)
            else:
                self._handle_user_cancelled(e)
        except Exception as e:
            self.show_error("Error", f"Full flow execution error: {e}")
        finally:
            self.close_progress_window(wait=False)
            self.run_on_main(self.finish_task, wait=False)

    def run_report(self, selected_area):
        self.excel.status_callback = self._make_threadsafe_status_callback()
        try:
            self._raise_if_cancelled()
            result = self.excel.generate_report(selected_area)
            if result.ok:
                output_doc = result.artifacts.get("report_doc") or result.artifacts.get("path") or ""
                self.show_info("Done", f"Report complete:\n{output_doc}", wait=False, close_progress=True)
            elif result.error_code == "USER_CANCELLED" or self.excel.was_cancelled:
                self._handle_user_cancelled()
            else:
                run_id = result.artifacts.get("run_id", "")
                suffix = f"\nrun_id: {run_id}" if run_id else ""
                self.show_error("Error", f"Report failed: {result.message}{suffix}")
        except excel_processing.UserCancelledError as e:
            self._handle_user_cancelled(e)
        except Exception as e:
            self.show_error("Error", f"Report execution error: {e}")
        finally:
            if self.progress_window:
                self.close_progress_window(wait=False)
            self.run_on_main(self.finish_task, wait=False)

    def update_input_sheet(self, file_path, reset_progress=True):
        result = self.excel.update_input_sheet(
            file_path=file_path,
            product=self.product_f_entry.get() or "",
            start_date=self.start_date_entry.get() or "",
            end_date=self.end_date_entry.get() or "",
        )
        if result.ok:
            if reset_progress:
                self.run_on_main(lambda: self.progress_window.update_progress(0) if self.progress_window else None, wait=False)
            return True
        if result.error_code == "USER_CANCELLED":
            self.excel.was_cancelled = True
            return False
        run_id = result.artifacts.get("run_id", "")
        suffix = f"\nrun_id: {run_id}" if run_id else ""
        self.excel.last_error = f"{result.message}{suffix}"
        return False

if __name__ == "__main__":
    # os.system("taskkill /f /im excel.exe >nul 2>&1")        #將Excel檔案清除
    python = sys.executable #測試
    root = tk.Tk()
    app = GUI(root)
    root.mainloop()
















