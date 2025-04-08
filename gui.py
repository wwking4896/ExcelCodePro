import tkinter as tk
from tkinter import filedialog, ttk, messagebox, simpledialog
from tkinter.scrolledtext import ScrolledText
import threading
import os
import pandas as pd
import time
import re
from excel_handler import ExcelHandler
from code_generator import CodeGenerator
from utils import excel_notation_to_index, save_config, load_config, get_templates_directory, get_resource_path
from version import VERSION, check_for_updates

class ExcelToCodeApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"ExcelCode Pro v{VERSION}")

        # 設置應用程式圖示
        try:
            icon_path = get_resource_path("icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            self.log(f"設置圖示時出錯: {str(e)}")

        # 檢查更新
        update_info = check_for_updates()
        if update_info:
            result = messagebox.askyesno(
                "有新版本可用", 
                f"發現新版本 {update_info['version']}，是否前往下載？"
            )
            if result:
                import webbrowser
                webbrowser.open(update_info['download_url'])

        self.root.geometry("1050x700")  # 增加預設寬度
        
        # 增加視窗大小變化的追蹤
        self.root.bind("<Configure>", self.on_window_resize)
        
        self.excel_files = []  # 修改为列表，存储多个文件
        self.dfs = {}  # 使用字典存储多个数据框
        self.selected_sheet = None
        self.selected_range = None
        self.selected_ranges = []  # 新增多選範圍列表
        self.code_template = None
        self.config_loading_completed = True  # 預設為已完成狀態
        self.loading_window = None  # 初始化 loading_window 為 None
        self.is_loading = False  # 追蹤載入狀態
        
        # 初始化處理器
        self.excel_handler = ExcelHandler(self)
        self.code_generator = CodeGenerator(self)
        
        self.create_widgets()
        
        # 添加視窗關閉事件處理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 載入最近的檔案記錄
        self.load_recent_files()

    def on_closing(self):
        """處理視窗關閉事件"""
        self.remember_recent_files()  # 儲存最近的檔案記錄
        self.root.destroy()

    def on_window_resize(self, event):
        """視窗大小變化時的處理函數"""
        # 只關注來自root視窗的事件
        if event.widget == self.root:
            # 日誌記錄可在開發時啟用，正式版可註釋
            # self.log(f"視窗大小變更: {event.width}x{event.height}")
            
            # 這裡可以添加對UI元素的動態調整
            pass
        
    def create_widgets(self):
        """建立應用程式的所有視覺元件，優化布局使程式碼區域更寬"""
        # 添加選單列
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 檔案選單
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="檔案", menu=file_menu)
        file_menu.add_command(label="開啟檔案...", command=self.browse_file)
        file_menu.add_command(label="開啟多個檔案...", command=self.browse_multiple_files)
        
        # 添加最近檔案的子選單
        self.recent_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="最近使用的檔案", menu=self.recent_menu)
        
        # 其他選單項目
        file_menu.add_separator()
        file_menu.add_command(label="儲存程式碼", command=self.save_code)
        file_menu.add_command(label="儲存設定", command=self.save_config_to_file)
        file_menu.add_command(label="載入設定", command=self.load_config_from_file)
        file_menu.add_separator()
        file_menu.add_command(label="結束", command=self.on_closing)
        
        # 幫助選單
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="幫助", menu=help_menu)
        help_menu.add_command(label="關於", command=self.show_about)
        
        # 建立主要框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 建立上半部區域 (控制區+程式碼區)
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill="both", expand=True, padx=0, pady=0)
        
        # 使用PanedWindow來允許使用者調整左右區域的比例
        self.paned_window = ttk.PanedWindow(top_frame, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill="both", expand=True, padx=0, pady=0)
        
        # 折疊面板狀態追蹤
        self.is_control_collapsed = False
        
        # 左側控制區容器
        self.control_container = ttk.Frame(self.paned_window)
        
        # 折疊按鈕和控制區布局
        control_wrapper = ttk.Frame(self.control_container)
        control_wrapper.pack(fill="both", expand=True)
        
        self.collapse_button = ttk.Button(control_wrapper, text="<<", width=2, command=self.toggle_control_panel)
        self.collapse_button.pack(side="left", fill="y")
        
        # 左側控制區 - 設定較小的固定寬度
        self.control_frame = ttk.Frame(control_wrapper, width=230)
        self.control_frame.pack(side="left", fill="both", expand=True)
        self.control_frame.pack_propagate(False)  # 防止子元件改變控制區的尺寸
        
        # 右側程式碼顯示區域 - 保存為實例變數
        self.code_frame = ttk.LabelFrame(self.paned_window, text="生成的程式碼")
        
        # 將左右區域添加到PanedWindow
        self.paned_window.add(self.control_container, weight=1)
        self.paned_window.add(self.code_frame, weight=3)  # 給予程式碼區域更多比重
        
        # ===== 控制區域內的元件 =====
        
        # 1. 檔案選擇區域 - 精簡布局
        file_frame = ttk.LabelFrame(self.control_frame, text="1. 選擇 Excel 檔案")
        file_frame.pack(fill="x", pady=(0, 5))
        
        file_buttons_frame = ttk.Frame(file_frame)
        file_buttons_frame.pack(side="top", fill="x", padx=5, pady=2)
        
        self.file_label = ttk.Label(file_frame, text="尚未選擇檔案")
        self.file_label.pack(side="left", padx=5, pady=5, fill="x", expand=True)
        
        # 將按鈕放在同一行，並使用更小的尺寸
        self.browse_button = ttk.Button(file_buttons_frame, text="瀏覽...", width=8, command=self.browse_file)
        self.browse_button.pack(side="left", padx=2)
        
        self.browse_multiple_button = ttk.Button(file_buttons_frame, text="多選...", width=8, command=self.browse_multiple_files)
        self.browse_multiple_button.pack(side="left", padx=2)
        
        # 檔案列表顯示 - 使用較小的高度
        self.files_listbox_frame = ttk.LabelFrame(self.control_frame, text="已選擇的檔案")
        self.files_listbox_frame.pack(fill="x", pady=(0, 5))
        
        self.files_listbox = tk.Listbox(self.files_listbox_frame, height=2)  # 減少高度
        self.files_listbox.pack(side="left", fill="both", expand=True, padx=5, pady=3)
        
        # 滾動條
        scrollbar = ttk.Scrollbar(self.files_listbox_frame, orient="vertical", command=self.files_listbox.yview)
        scrollbar.pack(side="right", fill="y", pady=3)
        self.files_listbox.config(yscrollcommand=scrollbar.set)
        
        # 2. 工作表選擇區域 - 更緊湊的布局
        sheet_frame = ttk.LabelFrame(self.control_frame, text="2. 選擇工作表")
        sheet_frame.pack(fill="x", pady=(0, 5))
        
        self.sheet_combobox = ttk.Combobox(sheet_frame, state="disabled")
        self.sheet_combobox.pack(fill="x", padx=5, pady=3)
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.on_sheet_selected)
        
        # 3. 範圍選擇區域 - 更緊湊的布局
        range_frame = ttk.LabelFrame(self.control_frame, text="3. 選擇資料範圍")
        range_frame.pack(fill="x", pady=(0, 5))
        
        range_btns_frame = ttk.Frame(range_frame)
        range_btns_frame.pack(fill="x", padx=5, pady=3)
        
        self.range_label = ttk.Label(range_frame, text="尚未選擇範圍")
        self.range_label.pack(fill="x", padx=5, pady=3)
        
        self.select_range_btn = ttk.Button(range_btns_frame, text="選擇範圍", width=10,
                                        command=self.select_multiple_ranges, state="disabled")
        self.select_range_btn.pack(side="left", padx=(0, 5))
        
        self.view_data_btn = ttk.Button(range_btns_frame, text="預覽資料", width=10,
                                        command=self.preview_data, state="disabled")
        self.view_data_btn.pack(side="right")
        
        # 4. 程式碼樣板 - 更緊湊的布局
        template_frame = ttk.LabelFrame(self.control_frame, text="4. 設定程式碼樣板")
        template_frame.pack(fill="x", pady=(0, 5))
        
        # 使用更緊湊的布局和更小的按鈕
        template_select_frame = ttk.Frame(template_frame)
        template_select_frame.pack(fill="x", padx=5, pady=3)
        
        template_btns_frame = ttk.Frame(template_select_frame)
        template_btns_frame.pack(fill="x", side="top", pady=2)
        
        # 自訂樣板按鈕
        self.template_btn = ttk.Button(template_btns_frame, text="設定自訂樣板", 
                                    command=self.set_template, state="disabled",
                                    width=14)
        self.template_btn.pack(side="left", padx=(0, 2))
        
        # 匯入樣板檔案按鈕
        self.import_template_btn = ttk.Button(template_btns_frame, text="匯入樣板", 
                                            command=self.import_template_file, state="disabled",
                                            width=8)
        self.import_template_btn.pack(side="left", padx=(0, 2))
        
        # 管理樣板按鈕
        self.manage_templates_btn = ttk.Button(template_btns_frame, text="管理樣板", 
                                            command=self.manage_templates,
                                            state="disabled", width=8)
        self.manage_templates_btn.pack(side="right", padx=0)
        
        # 預設樣板選擇區
        preset_frame = ttk.Frame(template_select_frame)
        preset_frame.pack(fill="x", side="top", pady=2)
        
        ttk.Label(preset_frame, text="或選擇預設:").pack(side="left", padx=(0, 5))
        self.template_combo = ttk.Combobox(preset_frame, state="disabled", width=20)
        self.template_combo.pack(side="left", fill="x", expand=True, padx=0)
        self.template_combo['values'] = ("陣列初始化", "二維陣列", "三維陣列", 
                                        "四維陣列 (範圍優先)", "四維陣列 (檔案優先)", 
                                        "三維多範圍陣列", "權重表設定", "多範圍處理")
        self.template_combo.bind("<<ComboboxSelected>>", self.on_template_selected)
        
        self.template_preview = ttk.Label(template_frame, text="尚未設定樣板")
        self.template_preview.pack(padx=5, pady=3, fill="x")
        
        # 生成按鈕
        self.generate_button = ttk.Button(self.control_frame, text="生成程式碼", 
                                        command=self.generate_code, state="disabled")
        self.generate_button.pack(fill="x", pady=3)
        
        # 程式碼字型大小控制區
        font_frame = ttk.Frame(self.control_frame)
        font_frame.pack(fill="x", pady=(3, 0))
        
        ttk.Label(font_frame, text="程式碼字型大小:").pack(side="left", padx=5)
        
        self.font_size = tk.StringVar(value="10")
        font_sizes = ["8", "9", "10", "11", "12", "14", "16"]
        self.font_size_combo = ttk.Combobox(font_frame, values=font_sizes, textvariable=self.font_size, width=3, state="readonly")
        self.font_size_combo.pack(side="left", padx=5)
        self.font_size_combo.bind("<<ComboboxSelected>>", self.on_font_size_changed)
        
        # 全螢幕模式按鈕
        self.fullscreen_var = tk.BooleanVar(value=False)
        self.fullscreen_btn = ttk.Checkbutton(font_frame, text="全螢幕", variable=self.fullscreen_var, 
                                            command=self.toggle_fullscreen)
        self.fullscreen_btn.pack(side="right", padx=5)
        
        # 程式碼顯示區
        self.code_text = ScrolledText(self.code_frame, wrap=tk.WORD, font=("Courier New", 10))
        self.code_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 新增底部Log顯示區域 - 使更緊湊
        log_frame = ttk.LabelFrame(main_frame, text="執行日誌")
        log_frame.pack(fill="x", expand=False, pady=(5, 0), ipady=3)
        
        self.log_text = ScrolledText(log_frame, wrap=tk.WORD, height=6, font=("Courier New", 9))
        self.log_text.pack(fill="both", expand=True, padx=5, pady=3)
        self.log_text.config(state="disabled")  # 初始設定為唯讀
        
        # 底部按鈕區域
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill="x", padx=10, pady=5)
        
        # 設定檔按鈕
        self.save_config_button = ttk.Button(bottom_frame, text="儲存設定", 
                                    command=self.save_config_to_file)
        self.save_config_button.pack(side="left", padx=5)
        
        self.load_config_button = ttk.Button(bottom_frame, text="載入設定", 
                                    command=self.load_config_from_file)
        self.load_config_button.pack(side="left", padx=5)
        
        # 原有的按鈕
        self.save_button = ttk.Button(bottom_frame, text="儲存程式碼", 
                                    command=self.save_code, state="disabled")
        self.save_button.pack(side="right", padx=5)
        
        self.copy_button = ttk.Button(bottom_frame, text="複製到剪貼簿", 
                                    command=self.copy_to_clipboard, state="disabled")
        self.copy_button.pack(side="right", padx=5)
    
    def log(self, message):
        """在日誌區域添加訊息"""
        # 獲取當前時間
        current_time = time.strftime("%H:%M:%S", time.localtime())
        
        # 設定為可編輯狀態
        self.log_text.config(state="normal")
        
        # 添加時間戳記和訊息
        self.log_text.insert(tk.END, f"[{current_time}] {message}\n")
        
        # 滾動到最底部
        self.log_text.see(tk.END)
        
        # 恢復唯讀狀態
        self.log_text.config(state="disabled")
        
        # 同時在控制台輸出，方便開發者查看
        # self.log(f"[{current_time}] {message}")

    def toggle_control_panel(self):
        """切換控制區的顯示/隱藏狀態"""
        if self.is_control_collapsed:
            # 展開控制區
            self.collapse_button.config(text="<<")
            self.control_frame.pack(side="left", fill="both", expand=True)
        else:
            # 摺疊控制區
            self.collapse_button.config(text=">>")
            self.control_frame.pack_forget()
        
        self.is_control_collapsed = not self.is_control_collapsed
        
        # 重新配置 PanedWindow 的權重來調整 UI
        if self.is_control_collapsed:
            # 當控制區被折疊時，給程式碼區域更多空間
            self.paned_window.forget(self.control_container)
            self.paned_window.forget(self.code_frame)
            self.paned_window.add(self.control_container, weight=0)
            self.paned_window.add(self.code_frame, weight=1)
        else:
            # 當控制區展開時，恢復原始配置
            self.paned_window.forget(self.control_container)
            self.paned_window.forget(self.code_frame)
            self.paned_window.add(self.control_container, weight=1)
            self.paned_window.add(self.code_frame, weight=3)

    def on_font_size_changed(self, event):
        """當選擇字型大小時的處理函數"""
        size = int(self.font_size.get())
        self.code_text.config(font=("Courier New", size))

    def toggle_fullscreen(self):
        """切換全螢幕模式"""
        if self.fullscreen_var.get():
            # 進入全螢幕模式
            self.root.attributes("-fullscreen", True)
        else:
            # 退出全螢幕模式
            self.root.attributes("-fullscreen", False)

    def on_template_selected(self, event):
        """處理預設樣板選擇"""
        template_name = self.template_combo.get()
        
        # 當選擇舊名稱「四維陣列」時，自動轉換為新名稱「四維陣列 (範圍優先)」以保持兼容性
        if template_name == "四維陣列":
            template_name = "四維陣列 (範圍優先)"
            self.template_combo.set(template_name)
        
        self.code_template = self.code_generator.get_default_template(self.template_combo.get())
        self.template_preview.config(text=f"已選擇: {self.template_combo.get()} 樣板")
        self.generate_button.config(state="normal")
    
    def show_loading_screen(self, message="正在處理中，請稍候..."):
        """顯示等待畫面"""
        # 如果已經有 loading 視窗，就更新訊息
        if self.is_loading and hasattr(self, 'loading_window') and self.loading_window:
            self.update_loading_message(message)
            return
            
        # 先關閉任何可能存在的之前的 loading 視窗
        self.hide_loading_screen()
        
        self.is_loading = True
        self.loading_window = tk.Toplevel(self.root)
        self.loading_window.title("處理中")
        
        # 取得主視窗的位置和大小
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 150
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 50
        self.loading_window.geometry(f"300x120+{x}+{y}")
        
        self.loading_window.transient(self.root)
        self.loading_window.grab_set()
        self.loading_window.resizable(False, False)
        
        # 進度訊息
        self.loading_message = ttk.Label(self.loading_window, text=message)
        self.loading_message.pack(pady=(15, 10))
        
        # 進度條
        self.progress = ttk.Progressbar(self.loading_window, mode="indeterminate", length=250)
        self.progress.pack(padx=25)
        self.progress.start(10)  # 開始動畫
        
        # 更新UI
        self.loading_window.update()
        
        # 新增關閉按鈕 (如果需要手動關閉)
        self.force_close_btn = ttk.Button(
            self.loading_window, 
            text="強制關閉", 
            command=self.hide_loading_screen
        )
        self.force_close_btn.pack(pady=(5, 10))
        
        # 設置協議處理器，確保用戶無法直接關閉這個窗口
        def on_close():
            self.log("使用者嘗試關閉載入視窗，但被攔截")
            pass  # 不做任何事情，防止使用者關閉
        
        self.loading_window.protocol("WM_DELETE_WINDOW", on_close)

    def update_loading_message(self, message):
        """更新載入訊息"""
        if self.is_loading and hasattr(self, 'loading_message') and self.loading_message:
            try:
                self.loading_message.config(text=message)
                if hasattr(self, 'loading_window') and self.loading_window:
                    self.loading_window.update()
            except Exception as e:
                self.log(f"更新載入訊息時出錯: {str(e)}")

    def hide_loading_screen(self):
        """關閉等待畫面"""
        if not hasattr(self, 'loading_window') or not self.loading_window:
            self.is_loading = False
            return
            
        try:
            self.log("正在關閉載入視窗...")
            if hasattr(self, 'progress'):
                self.progress.stop()
            self.loading_window.destroy()
            self.log("載入視窗已關閉")
        except Exception as e:
            self.log(f"關閉處理中視窗時出錯: {str(e)}")
        finally:
            self.loading_window = None
            self.is_loading = False
            # 再次確認視窗已關閉
            if hasattr(self, 'loading_window') and self.loading_window:
                try:
                    self.loading_window.destroy()
                except:
                    pass
        
    def browse_file(self):
        """選擇單一 Excel 檔案"""
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 檔案",
            filetypes=[("Excel 檔案", "*.xlsx *.xls")]
        )
        
        if file_path:
            self.excel_files = [file_path]  # 單文件模式，清除先前的文件列表
            self.file_label.config(text=os.path.basename(file_path))
            self.update_files_listbox()
            self.excel_handler.load_sheets(self.excel_files)
            
            # 記住這個文件
            self.remember_recent_files()
            
            # 更新最近文件選單
            self.load_recent_files()
    
    def browse_multiple_files(self):
        """選擇多個 Excel 檔案"""
        file_paths = filedialog.askopenfilenames(
            title="選擇多個 Excel 檔案",
            filetypes=[("Excel 檔案", "*.xlsx *.xls")]
        )
        
        if file_paths:
            self.show_loading_screen("正在處理選擇的檔案...")
            
            def process_files():
                try:
                    self.excel_files = list(file_paths)  # 更新文件列表
                    # 在主線程中更新UI
                    self.root.after(0, lambda: self.file_label.config(text=f"已選擇 {len(self.excel_files)} 個檔案"))
                    self.root.after(0, self.update_files_listbox)
                    self.root.after(0, lambda: self.excel_handler.load_sheets(self.excel_files))
                    
                    # 記住這些檔案
                    self.root.after(0, self.remember_recent_files)
                    
                    # 更新最近檔案選單
                    self.root.after(0, self.load_recent_files)
                finally:
                    self.root.after(0, self.hide_loading_screen)
            
            thread = threading.Thread(target=process_files)
            thread.daemon = True
            thread.start()
    
    def update_files_listbox(self):
        """更新文件列表显示"""
        self.files_listbox.delete(0, tk.END)
        for file_path in self.excel_files:
            self.files_listbox.insert(tk.END, os.path.basename(file_path))
    
    def on_sheet_selected(self, event):
        self.selected_sheet = self.sheet_combobox.get()
        if self.selected_sheet:
            self.excel_handler.on_sheet_selected(event)
    
    def select_multiple_ranges(self):
        self.excel_handler.select_multiple_ranges()
    
    def preview_data(self):
        self.excel_handler.preview_data()
    
    def set_template(self):
        self.code_generator.set_template()
    
    def import_template_file(self):
        """從檔案匯入樣板"""
        file_path = filedialog.askopenfilename(
            title="匯入樣板檔案",
            filetypes=[("文字檔案", "*.txt"), ("C/C++檔案", "*.c;*.cpp;*.h"), ("所有檔案", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    self.code_template = f.read()
                
                self.template_preview.config(text=f"已從檔案匯入樣板: {os.path.basename(file_path)}")
                self.template_combo.set("")  # 清空預設模板選擇
                self.generate_button.config(state="normal")
            except Exception as e:
                messagebox.showerror("錯誤", f"無法讀取樣板檔案: {str(e)}")
    
    def manage_templates(self):
        """管理已儲存的樣板"""
        templates_dir = get_templates_directory()
        
        # 建立管理窗口
        manager_dialog = tk.Toplevel(self.root)
        manager_dialog.title("樣板管理")
        manager_dialog.geometry("600x400")
        manager_dialog.grab_set()
        
        # 顯示已有的樣板
        templates_frame = ttk.LabelFrame(manager_dialog, text="已儲存的樣板")
        templates_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 建立樣板列表
        templates_list = tk.Listbox(templates_frame)
        templates_list.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        
        # 滾動條
        scrollbar = ttk.Scrollbar(templates_frame, orient="vertical", command=templates_list.yview)
        scrollbar.pack(side="right", fill="y")
        templates_list.config(yscrollcommand=scrollbar.set)
        
        # 載入樣板列表
        template_files = [f for f in os.listdir(templates_dir) if f.endswith('.txt') or f.endswith('.c') or f.endswith('.cpp') or f.endswith('.h')]
        for template_file in template_files:
            templates_list.insert(tk.END, template_file)
        
        # 按鈕區域
        btn_frame = ttk.Frame(manager_dialog)
        btn_frame.pack(fill="x", padx=10, pady=10)
        
        # 載入選中的樣板
        def load_selected_template():
            selection = templates_list.curselection()
            if selection:
                template_file = templates_list.get(selection[0])
                template_path = os.path.join(templates_dir, template_file)
                
                try:
                    with open(template_path, 'r', encoding='utf-8') as f:
                        self.code_template = f.read()
                    
                    self.template_preview.config(text=f"已載入樣板: {template_file}")
                    self.template_combo.set("")  # 清空預設模板選擇
                    self.generate_button.config(state="normal")
                    manager_dialog.destroy()
                except Exception as e:
                    messagebox.showerror("錯誤", f"無法讀取樣板: {str(e)}", parent=manager_dialog)
        
        # 刪除選中的樣板
        def delete_selected_template():
            selection = templates_list.curselection()
            if selection:
                template_file = templates_list.get(selection[0])
                template_path = os.path.join(templates_dir, template_file)
                
                if messagebox.askyesno("確認刪除", f"確定要刪除樣板 {template_file} 嗎？", parent=manager_dialog):
                    try:
                        os.remove(template_path)
                        templates_list.delete(selection[0])
                    except Exception as e:
                        messagebox.showerror("錯誤", f"無法刪除樣板: {str(e)}", parent=manager_dialog)
        
        # 儲存當前樣板
        def save_current_template():
            if not self.code_template:
                messagebox.showwarning("警告", "目前沒有設定樣板", parent=manager_dialog)
                return
                
            template_name = simpledialog.askstring("儲存樣板", "請輸入樣板名稱:", parent=manager_dialog)
            if template_name:
                if not (template_name.endswith('.txt') or template_name.endswith('.c') or template_name.endswith('.cpp') or template_name.endswith('.h')):
                    template_name += '.txt'
                    
                template_path = os.path.join(templates_dir, template_name)
                
                # 檢查檔案是否已存在
                if os.path.exists(template_path):
                    if not messagebox.askyesno("覆蓋確認", f"樣板 {template_name} 已存在，是否覆蓋？", parent=manager_dialog):
                        return
                
                try:
                    with open(template_path, 'w', encoding='utf-8') as f:
                        f.write(self.code_template)
                    
                    # 更新列表
                    if template_name not in templates_list.get(0, tk.END):
                        templates_list.insert(tk.END, template_name)
                    messagebox.showinfo("成功", f"樣板已儲存為 {template_name}", parent=manager_dialog)
                except Exception as e:
                    messagebox.showerror("錯誤", f"無法儲存樣板: {str(e)}", parent=manager_dialog)
        
        # 預覽選中的樣板
        def preview_selected_template():
            selection = templates_list.curselection()
            if selection:
                template_file = templates_list.get(selection[0])
                template_path = os.path.join(templates_dir, template_file)
                
                try:
                    with open(template_path, 'r', encoding='utf-8') as f:
                        template_content = f.read()
                    
                    # 建立預覽視窗
                    preview_window = tk.Toplevel(manager_dialog)
                    preview_window.title(f"預覽樣板: {template_file}")
                    preview_window.geometry("700x500")
                    
                    preview_text = ScrolledText(preview_window, wrap=tk.WORD, font=("Courier New", 10))
                    preview_text.pack(fill="both", expand=True, padx=10, pady=10)
                    preview_text.insert("1.0", template_content)
                    preview_text.config(state="disabled")
                    
                    ttk.Button(preview_window, text="關閉", command=preview_window.destroy).pack(pady=10)
                    
                except Exception as e:
                    messagebox.showerror("錯誤", f"無法預覽樣板: {str(e)}", parent=manager_dialog)
        
        # 新增按鈕
        ttk.Button(btn_frame, text="載入選中樣板", command=load_selected_template).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="預覽選中樣板", command=preview_selected_template).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="刪除選中樣板", command=delete_selected_template).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="儲存當前樣板", command=save_current_template).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="關閉", command=manager_dialog.destroy).pack(side="right", padx=5)
    
    def generate_code(self):
        if not self.selected_range or not self.code_template or not self.excel_files:
            return
        
        self.show_loading_screen("正在生成程式碼，請稍候...")
        
        def generate_task():
            try:
                final_code = self.code_generator.generate_code(
                    self.excel_files, 
                    self.dfs, 
                    self.selected_ranges if hasattr(self, 'selected_ranges') else [], 
                    self.code_template, 
                    self.selected_range
                )
                
                # 在主線程中顯示生成的代碼
                self.root.after(0, lambda: self.code_text.delete("1.0", tk.END))
                self.root.after(0, lambda: self.code_text.insert("1.0", final_code))
                
                self.root.after(0, lambda: self.save_button.config(state="normal"))
                self.root.after(0, lambda: self.copy_button.config(state="normal"))
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"生成程式碼時發生錯誤: {str(e)}"))
                # 顯示詳細的錯誤信息
                import traceback
                self.log(traceback.format_exc())
            finally:
                self.root.after(0, self.hide_loading_screen)
        
        thread = threading.Thread(target=generate_task)
        thread.daemon = True
        thread.start()

    def save_code(self):
        save_path = filedialog.asksaveasfilename(
            title="儲存程式碼",
            defaultextension=".c",
            filetypes=[("C 程式碼", "*.c"), ("C++ 程式碼", "*.cpp"), ("所有檔案", "*.*")]
        )
        
        if save_path:
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(self.code_text.get("1.0", tk.END))
            messagebox.showinfo("成功", f"程式碼已儲存至 {save_path}")

    def copy_to_clipboard(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.code_text.get("1.0", tk.END))
        messagebox.showinfo("成功", "程式碼已複製到剪貼簿")
        
    def save_config_to_file(self):
        """儲存目前的設定到檔案"""
        if not self.excel_files:
            messagebox.showwarning("警告", "尚未選擇任何檔案，無法儲存設定")
            return
                
        # 開啟檔案對話框讓使用者選擇儲存位置
        file_path = filedialog.asksaveasfilename(
            title="儲存設定檔",
            defaultextension=".json",
            filetypes=[("JSON 檔案", "*.json"), ("所有檔案", "*.*")]
        )
        
        if not file_path:
            return  # 使用者取消了操作
        
        # 收集目前的設定資訊
        config_data = {
            "excel_files": self.excel_files,
            "selected_sheet": self.selected_sheet,
            "selected_ranges": self.selected_ranges,
        }
        
        # 確定並儲存模板類型和內容
        selected_preset = self.template_combo.get()
        if selected_preset and selected_preset in self.template_combo['values']:
            # 使用預設模板
            config_data["template_type"] = "preset"
            config_data["preset_template"] = selected_preset
            # 同時儲存模板內容，以便向下相容
            config_data["code_template"] = self.code_template
        elif self.code_template:
            # 使用自訂模板
            config_data["template_type"] = "custom"
            config_data["code_template"] = self.code_template
        
        # 儲存設定到檔案
        if save_config(config_data, file_path):
            messagebox.showinfo("成功", f"設定已儲存至 {file_path}")
        else:
            messagebox.showerror("錯誤", "儲存設定檔時發生錯誤")
    
    def load_excel_sheets_without_close(self, excel_files):
        """載入Excel工作表但不自動關閉載入視窗"""
        self.log("開始載入工作表 (不自動關閉載入視窗)...")
        
        # 更新載入訊息
        self.update_loading_message(f"載入 {len(excel_files)} 個 Excel 檔案的工作表...")
        
        def load_task():
            try:
                # 假設所有文件有相同的工作表，使用第一個文件獲取工作表列表
                if not excel_files:
                    self.log("沒有Excel檔案可載入")
                    return
                        
                xl = pd.ExcelFile(excel_files[0])
                sheets = xl.sheet_names
                
                # 在主線程中更新UI
                self.root.after(0, lambda: self.sheet_combobox.config(values=sheets, state="readonly"))
                self.root.after(0, lambda: self.sheet_combobox.current(0))
                
                self.log(f"已載入工作表: {sheets}")
                
            except Exception as e:
                # 在主線程中顯示錯誤
                self.log(f"載入工作表時發生錯誤: {str(e)}")
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"無法讀取 Excel 檔案: {str(e)}"))
        
        # 直接在當前線程執行，以確保完成後代碼可以繼續執行
        load_task()
        
    def excel_sheet_selected_without_close(self, sheet_name):
        """選擇工作表但不自動關閉載入視窗"""
        self.log(f"選擇工作表: {sheet_name} (不自動關閉載入視窗)...")
        
        # 更新載入訊息
        self.update_loading_message(f"載入工作表 '{sheet_name}' 的資料...")
        
        try:
            # 逐个加载每个文件的选定工作表
            for file_path in self.excel_files:
                # 修改讀取方式，先用 object 讀取所有數據，再嘗試轉換為數值類型
                df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=object)

                # 定義轉換函數，避免使用 errors='ignore'
                def convert_to_numeric(column):
                    try:
                        return pd.to_numeric(column)
                    except (ValueError, TypeError):
                        return column
                # 嘗試將可以轉換為數值的列轉為數值類型
                df = df.apply(convert_to_numeric)

                # 重置索引，確保從0開始連續
                df = df.reset_index(drop=True)
                
                # 打印数据框信息
                self.log(f"讀取檔案: {os.path.basename(file_path)}")
                self.log(f"資料框形狀: {df.shape}")
                
                self.dfs[file_path] = df
            
            # 在主線程中更新UI
            self.root.after(0, lambda: self.select_range_btn.config(state="normal"))
            self.root.after(0, lambda: self.range_label.config(text="尚未選擇範圍"))
            self.root.after(0, lambda: self.view_data_btn.config(state="disabled"))
            self.root.after(0, lambda: self.template_btn.config(state="disabled"))
            self.root.after(0, lambda: self.template_combo.config(state="disabled"))
            self.root.after(0, lambda: self.import_template_btn.config(state="disabled"))
            self.root.after(0, lambda: self.manage_templates_btn.config(state="disabled"))
            self.root.after(0, lambda: self.generate_button.config(state="disabled"))
            
            # 清空之前選擇的範圍
            self.selected_range = None
            self.selected_ranges = []
            
            self.log("工作表資料載入完成")
            
        except Exception as e:
            self.log(f"工作表 '{sheet_name}' 載入資料時發生錯誤: {str(e)}")
            import traceback
            self.log(traceback.format_exc())
            raise
            
    def load_config_from_file(self):
        """從檔案載入設定"""
        file_path = filedialog.askopenfilename(
            title="載入設定檔",
            filetypes=[("JSON 檔案", "*.json"), ("所有檔案", "*.*")]
        )
        
        if not file_path:
            return  # 使用者取消了操作
        
        # 載入設定資料
        config_data = load_config(file_path)
        if not config_data:
            messagebox.showerror("錯誤", "載入設定檔時發生錯誤")
            return
        
        # 開始載入設定前，顯示等待視窗
        self.show_loading_screen("正在載入設定，請稍候...")
        
        # 設定全域變數追蹤載入狀態
        self.config_loading_completed = False
        # 記錄開始時間以實現超時機制
        self.loading_start_time = time.time()
        
        def apply_config():
            try:
                # 暫存原始方法，以確保所有流程結束後可以恢復
                original_hide_loading = self.hide_loading_screen
                
                # 建立暫時方法，防止提前關閉載入視窗
                def temp_hide_loading():
                    self.log("暫時攔截關閉載入視窗操作")
                    pass
                
                # 處理模板設定
                if "template_type" in config_data:
                    template_type = config_data["template_type"]
                    self.log(f"載入模板類型: {template_type}")
                    self.update_loading_message(f"載入模板: {template_type}...")
                    
                    if template_type == "custom" and "code_template" in config_data:
                        # 載入自訂模板
                        self.log("載入自訂樣板...")
                        self.code_template = config_data["code_template"]
                        self.root.after(0, lambda: self.template_preview.config(text="已載入自訂樣板"))
                        self.root.after(0, lambda: self.template_combo.set(""))  # 清空預設模板選擇
                        
                    elif template_type == "preset" and "preset_template" in config_data:
                        # 載入預設模板選擇
                        preset_name = config_data["preset_template"]
                        self.log(f"載入預設樣板: {preset_name}")
                        
                        # 檢查預設樣板名稱是否在可用選項中
                        def set_preset_template():
                            if preset_name in self.template_combo['values']:
                                self.template_combo.set(preset_name)
                                # 觸發模板選擇事件
                                self.on_template_selected(None)
                            else:
                                self.log(f"找不到預設樣板: {preset_name}")
                                messagebox.showwarning("警告", f"找不到預設樣板: {preset_name}")
                        
                        self.root.after(0, set_preset_template)
                elif "code_template" in config_data:
                    # 向下相容舊的設定檔格式
                    self.log("使用舊格式載入樣板...")
                    self.code_template = config_data["code_template"]
                    self.root.after(0, lambda: self.template_preview.config(text="已載入自訂樣板"))
                    self.root.after(0, lambda: self.template_combo.set(""))  # 清空預設模板選擇
                
                # 套用設定
                # 1. 載入 Excel 檔案
                if "excel_files" in config_data and config_data["excel_files"]:
                    # 檢查檔案是否存在
                    existing_files = []
                    for file_path in config_data["excel_files"]:
                        if os.path.exists(file_path):
                            existing_files.append(file_path)
                        else:
                            self.log(f"警告: 檔案 {file_path} 不存在，將被跳過")
                    
                    if existing_files:
                        self.update_loading_message(f"載入 {len(existing_files)} 個 Excel 檔案...")
                        
                        # 在 UI 執行緒更新檔案列表
                        def update_file_list():
                            self.excel_files = existing_files
                            self.file_label.config(text=f"已載入 {len(self.excel_files)} 個檔案")
                            self.update_files_listbox()
                        
                        self.root.after(0, update_file_list)
                        
                        # 儲存要載入的工作表名稱以便稍後使用
                        target_sheet = config_data.get("selected_sheet", None)
                        
                        try:
                            # 直接載入工作表 (不使用多線程和不自動關閉載入視窗)
                            self.load_excel_sheets_without_close(existing_files)
                            
                            # 如果有指定的工作表，則選擇它
                            if target_sheet and target_sheet in self.sheet_combobox['values']:
                                self.log(f"選擇工作表: {target_sheet}")
                                self.update_loading_message(f"選擇工作表: {target_sheet}...")
                                
                                # 設置工作表
                                self.sheet_combobox.set(target_sheet)
                                self.selected_sheet = target_sheet
                                
                                # 載入選定工作表的資料
                                try:
                                    self.excel_sheet_selected_without_close(target_sheet)
                                    
                                    # 設定選擇的範圍
                                    if "selected_ranges" in config_data and config_data["selected_ranges"]:
                                        self.log("設定資料範圍...")
                                        self.update_loading_message("設定資料範圍...")
                                        
                                        self.selected_ranges = config_data["selected_ranges"]
                                        if self.selected_ranges:
                                            self.selected_range = self.selected_ranges[0]  # 相容舊邏輯
                                            range_strs = [r['range_str'] for r in self.selected_ranges]
                                            
                                            def update_range_ui():
                                                self.range_label.config(text=f"已選擇 {len(range_strs)} 個範圍: {', '.join(range_strs)}")
                                                self.view_data_btn.config(state="normal")
                                                self.template_btn.config(state="normal")
                                                self.template_combo.config(state="normal")
                                                self.import_template_btn.config(state="normal")
                                                self.manage_templates_btn.config(state="normal")
                                                
                                                # 確保在有範圍和樣板的情況下啟用生成按鈕
                                                if self.code_template:
                                                    self.log("啟用生成程式碼按鈕...")
                                                    self.generate_button.config(state="normal")
                                            
                                            self.root.after(0, update_range_ui)
                                except Exception as e:
                                    self.log(f"選擇工作表 '{target_sheet}' 時發生錯誤: {str(e)}")
                                    import traceback
                                    self.log(traceback.format_exc())
                            else:
                                # 如果沒有找到指定的工作表
                                self.log(f"工作表 '{target_sheet}' 不存在於所選檔案中")
                                if target_sheet:
                                    self.root.after(0, lambda: messagebox.showwarning(
                                        "警告", 
                                        f"找不到設定檔中指定的工作表 '{target_sheet}'，請手動選擇工作表"
                                    ))
                        except Exception as e:
                            self.log(f"載入工作表時發生錯誤: {str(e)}")
                            import traceback
                            self.log(traceback.format_exc())
                    else:
                        # 即使沒有有效檔案，也確保在有樣板的情況下顯示
                        if self.code_template:
                            self.log("無法載入檔案，但已載入程式碼樣板")
                        
                        self.root.after(0, lambda: messagebox.showwarning(
                            "警告", 
                            "設定檔中的檔案路徑均無效，請重新選擇檔案"
                        ))
                else:
                    # 即使沒有檔案路徑，也確保在有樣板的情況下顯示
                    if self.code_template:
                        self.log("設定檔中沒有檔案路徑，但已載入程式碼樣板")
                    
                    self.root.after(0, lambda: messagebox.showwarning(
                        "警告", 
                        "設定檔中沒有指定檔案路徑，請手動選擇檔案"
                    ))
            except Exception as e:
                # 發生例外時的處理
                self.log(f"載入設定時發生未預期的錯誤: {e}")
                import traceback
                self.log(traceback.format_exc())
                
                self.root.after(0, lambda: messagebox.showerror(
                    "錯誤", 
                    f"載入設定時發生錯誤: {str(e)}"
                ))
            finally:
                # 標記載入完成並關閉載入視窗
                self.config_loading_completed = True
                self.root.after(0, self.hide_loading_screen)
        
        # 啟動處理任務
        thread = threading.Thread(target=apply_config)
        thread.daemon = True
        thread.start()
        
        # 定期檢查載入狀態，確保 loading 視窗在適當的時間關閉
        def check_loading_completed():
            # 檢查是否超時 (例如 60 秒)
            if hasattr(self, 'loading_start_time'):
                elapsed_time = time.time() - self.loading_start_time
                if elapsed_time > 60:
                    self.log(f"載入設定已超時 ({elapsed_time:.1f}秒)，強制結束")
                    self.config_loading_completed = True
                    self.root.after(0, self.hide_loading_screen)
                    self.root.after(0, lambda: messagebox.showwarning(
                        "警告", 
                        "載入設定超時，請檢查檔案狀態或重試"
                    ))
                    return
                
                self.log(f"檢查載入狀態 (已等待 {elapsed_time:.1f}秒): config_loading_completed = {getattr(self, 'config_loading_completed', False)}")
            
            if not hasattr(self, 'config_loading_completed'):
                self.log("警告: 找不到 config_loading_completed 變數")
                # 如果找不到變數，直接強制關閉 loading 視窗
                self.config_loading_completed = True
                self.root.after(0, self.hide_loading_screen)
                return
                
            if not self.config_loading_completed:
                # 如果載入尚未完成，繼續檢查
                self.root.after(500, check_loading_completed)
            else:
                self.log("載入完成，確保關閉 loading 視窗")
                self.root.after(0, self.hide_loading_screen)
        
        # 開始檢查載入狀態
        self.root.after(500, check_loading_completed)

    def remember_recent_files(self):
        """記住最近使用的文件"""
        try:
            recent_files = {}
            recent_files_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "recent_files.json")
            
            # 載入既有的記錄（如果存在）
            if os.path.exists(recent_files_path):
                with open(recent_files_path, "r", encoding="utf-8") as f:
                    recent_files = json.load(f)
            
            # 更新最近使用的檔案列表
            if self.excel_files:
                # 只保存檔案路徑，限制數量為最近10個
                recent_files["last_used"] = self.excel_files[:10]
                # 記錄最後一次的工作表選擇
                if self.selected_sheet:
                    recent_files["last_sheet"] = self.selected_sheet
                
                # 儲存更新後的記錄
                with open(recent_files_path, "w", encoding="utf-8") as f:
                    json.dump(recent_files, f, ensure_ascii=False, indent=4)
                
                self.log("已更新最近使用的檔案記錄")
        except Exception as e:
            self.log(f"記錄最近檔案時出錯: {str(e)}")

    def load_recent_files(self):
        """載入最近使用的檔案記錄"""
        try:
            recent_files_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "recent_files.json")
            if os.path.exists(recent_files_path):
                with open(recent_files_path, "r", encoding="utf-8") as f:
                    recent_files = json.load(f)
                
                # 檢查檔案是否存在
                if "last_used" in recent_files and recent_files["last_used"]:
                    # 添加「最近檔案」選單項目
                    self.add_recent_files_menu(recent_files["last_used"])
        except Exception as e:
            self.log(f"載入最近檔案記錄時出錯: {str(e)}")

    def add_recent_files_menu(self, recent_files):
        """更新最近檔案選單"""
        # 清空現有的選單項目
        self.recent_menu.delete(0, tk.END)
        
        # 檢查是否有最近的檔案
        if not recent_files:
            self.recent_menu.add_command(label="（無最近檔案）", state="disabled")
            return
        
        # 添加最近檔案到選單
        for file_path in recent_files:
            if os.path.exists(file_path):
                # 使用 lambda 來傳遞不同的檔案路徑
                self.recent_menu.add_command(
                    label=os.path.basename(file_path),
                    command=lambda fp=file_path: self.open_recent_file(fp)
                )
        
        # 添加清除記錄選項
        if recent_files:
            self.recent_menu.add_separator()
            self.recent_menu.add_command(label="清除最近檔案記錄", command=self.clear_recent_files)

    def open_recent_file(self, file_path):
        """開啟最近使用的檔案"""
        if os.path.exists(file_path):
            self.excel_files = [file_path]
            self.file_label.config(text=os.path.basename(file_path))
            self.update_files_listbox()
            self.excel_handler.load_sheets(self.excel_files)
        else:
            messagebox.showwarning("檔案不存在", f"檔案 {file_path} 已不存在")
            # 從記錄中移除不存在的檔案
            self.remove_nonexistent_file(file_path)

    def remove_nonexistent_file(self, file_path):
        """從最近檔案記錄中移除不存在的檔案"""
        try:
            recent_files_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "recent_files.json")
            if os.path.exists(recent_files_path):
                with open(recent_files_path, "r", encoding="utf-8") as f:
                    recent_files = json.load(f)
                
                if "last_used" in recent_files and file_path in recent_files["last_used"]:
                    recent_files["last_used"].remove(file_path)
                    
                    with open(recent_files_path, "w", encoding="utf-8") as f:
                        json.dump(recent_files, f, ensure_ascii=False, indent=4)
                    
                    # 更新選單
                    self.add_recent_files_menu(recent_files["last_used"])
        except Exception as e:
            self.log(f"移除不存在的檔案記錄時出錯: {str(e)}")

    def clear_recent_files(self):
        """清除最近檔案記錄"""
        if messagebox.askyesno("確認", "確定要清除所有最近檔案記錄嗎？"):
            try:
                recent_files_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "recent_files.json")
                if os.path.exists(recent_files_path):
                    with open(recent_files_path, "r", encoding="utf-8") as f:
                        recent_files = json.load(f)
                    
                    recent_files["last_used"] = []
                    
                    with open(recent_files_path, "w", encoding="utf-8") as f:
                        json.dump(recent_files, f, ensure_ascii=False, indent=4)
                    
                    # 更新選單
                    self.add_recent_files_menu([])
                    
                    self.log("已清除最近檔案記錄")
            except Exception as e:
                self.log(f"清除最近檔案記錄時出錯: {str(e)}")

    def show_about(self):
        """顯示關於對話框"""
        about_text = "ExcelCode Pro\n\n"
        about_text += "版本: 1.0.0\n\n"
        about_text += "此工具用於將 Excel 表格資料轉換為 C/C++ 程式碼。\n"
        about_text += "支援單檔案和多檔案處理，多範圍選擇，\n"
        about_text += "以及自訂程式碼樣板。\n\n"
        about_text += "© 2024 WWKing - Alphabet Studio 版權所有"
        
        messagebox.showinfo("關於", about_text)

    