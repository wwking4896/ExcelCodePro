import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import re
from utils import excel_notation_to_index

class ExcelHandler:
    def __init__(self, gui_instance):
        self.gui = gui_instance  # 保存對主GUI實例的引用
        self.df = None  # 臨時存儲當前處理的數據框
    
    def load_sheets(self, excel_files):
        """加載所有選中文件中的工作表"""
        self.gui.show_loading_screen("正在載入Excel檔案，請稍候...")
        
        def load_task():
            try:
                # 假設所有文件有相同的工作表，使用第一個文件獲取工作表列表
                if not excel_files:
                    self.gui.root.after(0, self.gui.hide_loading_screen)
                    return
                        
                xl = pd.ExcelFile(excel_files[0])
                sheets = xl.sheet_names
                
                # 在主線程中更新UI
                self.gui.root.after(0, lambda: self.gui.sheet_combobox.config(values=sheets, state="readonly"))
                self.gui.root.after(0, lambda: self.gui.sheet_combobox.current(0))
                
            except Exception as e:
                # 在主線程中顯示錯誤
                self.gui.root.after(0, lambda: messagebox.showerror("錯誤", f"無法讀取 Excel 檔案: {str(e)}"))
            finally:
                # 在主線程中關閉等待畫面
                self.gui.root.after(0, self.gui.hide_loading_screen)
            
        # 啟動一個獨立線程來執行耗時任務
        thread = threading.Thread(target=load_task)
        thread.daemon = True
        thread.start()
    
    def on_sheet_selected(self, event):
        """當選擇工作表時的處理"""
        selected_sheet = self.gui.sheet_combobox.get()
        if selected_sheet:
            self.gui.show_loading_screen("正在讀取所選工作表中的數據...")
            
            def load_data_task():
                try:
                    # 逐个加载每个文件的选定工作表
                    for file_path in self.gui.excel_files:
                        # 修改讀取方式，先用 object 讀取所有數據，再嘗試轉換為數值類型
                        df = pd.read_excel(file_path, sheet_name=selected_sheet, dtype=object)
                        # 嘗試將可以轉換為數值的列轉為數值類型
                        df = df.apply(pd.to_numeric, errors='ignore')
                        # 重置索引，確保從0開始連續
                        df = df.reset_index(drop=True)
                        
                        # 打印数据框信息
                        self.gui.log(f"讀取檔案: {os.path.basename(file_path)}")
                        self.gui.log(f"資料框形狀: {df.shape}")
                        self.gui.log(f"資料框欄位: {df.columns.tolist()}")
                        self.gui.log(f"資料範例 (前3行):\n{df.head(3)}")
                        
                        self.gui.dfs[file_path] = df
                    
                    # 在主線程中更新UI
                    self.gui.root.after(0, lambda: self.gui.select_range_btn.config(state="normal"))
                    self.gui.root.after(0, lambda: self.gui.range_label.config(text="尚未選擇範圍"))
                    self.gui.root.after(0, lambda: self.gui.view_data_btn.config(state="disabled"))
                    self.gui.root.after(0, lambda: self.gui.template_btn.config(state="disabled"))
                    self.gui.root.after(0, lambda: self.gui.template_combo.config(state="disabled"))
                    self.gui.root.after(0, lambda: self.gui.generate_button.config(state="disabled"))
                    
                    # 清空之前選擇的範圍
                    self.gui.root.after(0, lambda: setattr(self.gui, 'selected_range', None))
                    self.gui.root.after(0, lambda: setattr(self.gui, 'selected_ranges', []))
                    
                except Exception as e:
                    self.gui.root.after(0, lambda: messagebox.showerror("錯誤", f"無法讀取工作表: {str(e)}"))
                    self.gui.root.after(0, lambda: self.gui.select_range_btn.config(state="disabled"))
                    # 顯示詳細錯誤
                    import traceback
                    self.gui.log(traceback.format_exc())
                finally:
                    # 確保在所有操作完成後關閉loading視窗
                    self.gui.root.after(0, self.gui.hide_loading_screen)
            
            # 使用新的線程執行耗時的數據讀取操作
            thread = threading.Thread(target=load_data_task)
            thread.daemon = True
            thread.start()
    
    def select_multiple_ranges(self):
        """選擇多個數據範圍"""
        # 使用第一个文件的数据框来设置范围
        if not self.gui.excel_files or not self.gui.dfs:
            return
                
        first_file = self.gui.excel_files[0]
        self.df = self.gui.dfs[first_file]  # 使用第一个文件设置范围
        
        # 建立一个范围选择窗口
        range_dialog = tk.Toplevel(self.gui.root)
        range_dialog.title("選擇多個資料範圍")
        range_dialog.geometry("500x400")
        range_dialog.grab_set()  # 模态窗口
        
        # 新增一個列表框顯示已選擇的範圍
        ranges_frame = ttk.LabelFrame(range_dialog, text="已選擇的範圍")
        ranges_frame.pack(fill="x", padx=10, pady=5)
        
        ranges_listbox = tk.Listbox(ranges_frame, height=5)
        ranges_listbox.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        
        # 填充已有的範圍
        for range_info in self.gui.selected_ranges:
            ranges_listbox.insert(tk.END, range_info['range_str'])
        
        # 滾動條
        scrollbar = ttk.Scrollbar(ranges_frame, orient="vertical", command=ranges_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        ranges_listbox.config(yscrollcommand=scrollbar.set)
        
        # 範圍輸入區域
        input_frame = ttk.Frame(range_dialog)
        input_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(input_frame, text="輸入範圍 (例如: A1:G10):").pack(side="left", padx=5)
        
        range_var = tk.StringVar()
        range_entry = ttk.Entry(input_frame, textvariable=range_var, width=20)
        range_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        # 添加和刪除按鈕
        btn_frame = ttk.Frame(range_dialog)
        btn_frame.pack(fill="x", padx=10, pady=5)
        
        def add_range():
            range_str = range_var.get().strip()
            if not range_str:
                messagebox.showerror("錯誤", "請輸入有效的範圍", parent=range_dialog)
                return
            
            try:
                # 解析範圍
                if ":" in range_str:
                    start, end = range_str.split(":")
                    
                    # 解析開始和結束座標
                    match_start = re.match(r'([A-Z]+)([0-9]+)', start.upper())
                    match_end = re.match(r'([A-Z]+)([0-9]+)', end.upper())
                    
                    if match_start and match_end:
                        # 提取列字母和行號
                        start_col_str, start_row_str = match_start.groups()
                        end_col_str, end_row_str = match_end.groups()
                        
                        # 將行號向下調整一行，因為目前的行號偏移導致實際選擇的是向下一行的數據
                        start_row = int(start_row_str)
                        end_row = int(end_row_str)
                        
                        # 如果目前輸入K3實際選擇的是K4的數據，那麼我們需要將行號-1來糾正
                        adjusted_start_row = start_row - 1
                        adjusted_end_row = end_row - 1
                        
                        # 打印調整信息
                        self.gui.log(f"使用者輸入範圍: {range_str}")
                        self.gui.log(f"調整為: {start_col_str}{adjusted_start_row}:{end_col_str}{adjusted_end_row}")
                        
                        # 轉換調整後的座標為索引
                        adjusted_start = f"{start_col_str}{adjusted_start_row}"
                        adjusted_end = f"{end_col_str}{adjusted_end_row}"
                        
                        # 使用調整後的座標獲取列索引和行索引
                        excel_start_row, start_col = excel_notation_to_index(adjusted_start)
                        excel_end_row, end_col = excel_notation_to_index(adjusted_end)
                        
                        # excel_notation_to_index 已經會將行號-1轉為索引
                        start_row = excel_start_row
                        end_row = excel_end_row
                        
                    else:
                        # 如果無法解析，使用原始功能進行解析
                        start_row, start_col = excel_notation_to_index(start)
                        end_row, end_col = excel_notation_to_index(end)
                    
                    # 驗證範圍有效性
                    if start_row > end_row or start_col > end_col:
                        messagebox.showerror("錯誤", "範圍無效: 起始位置必須小於等於結束位置", parent=range_dialog)
                        return
                    
                    # 檢查是否為負數
                    if start_row < 0 or start_col < 0:
                        messagebox.showerror("錯誤", "範圍無效: 索引不能為負數", parent=range_dialog)
                        return
                    
                    # 打印最終使用的索引
                    self.gui.log(f"最終使用的DataFrame索引: 從 ({start_row}, {start_col}) 到 ({end_row}, {end_col})")
                    
                    # 創建範圍信息
                    range_info = {
                        'start_row': start_row,
                        'start_col': start_col,
                        'end_row': end_row,
                        'end_col': end_col,
                        'range_str': range_str  # 保留原始輸入
                    }
                    
                    # 添加到範圍列表
                    self.gui.selected_ranges.append(range_info)
                    ranges_listbox.insert(tk.END, range_str)
                    range_var.set("")  # 清空輸入框
                else:
                    messagebox.showerror("錯誤", "範圍格式不正確，請使用如 A1:G10 的格式", parent=range_dialog)
            except Exception as e:
                messagebox.showerror("錯誤", f"無法解析範圍: {str(e)}", parent=range_dialog)
                import traceback
                self.gui.log(traceback.format_exc())

        def remove_range():
            selected_idx = ranges_listbox.curselection()
            if selected_idx:
                idx = selected_idx[0]
                # 從列表和界面中刪除
                del self.gui.selected_ranges[idx]
                ranges_listbox.delete(idx)
            else:
                messagebox.showinfo("提示", "請先選擇要刪除的範圍", parent=range_dialog)
        
        ttk.Button(btn_frame, text="添加範圍", command=add_range).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="刪除所選範圍", command=remove_range).pack(side="left", padx=5)
        
        # 顯示工作表數據預覽
        preview_frame = ttk.LabelFrame(range_dialog, text="資料預覽（前5行）")
        preview_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        preview_text = ScrolledText(preview_frame, height=8, font=("Courier New", 9))
        preview_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 顯示表頭和前5行數據
        try:
            # 添加列標題行
            headers = [f"{self.get_column_letter(i)}" for i in range(min(10, self.df.shape[1]))]
            preview_text.insert("1.0", "行號  " + "  ".join(headers) + "\n")
            preview_text.insert("end", "───" + "─"*50 + "\n")
            
            # 顯示前5行實際資料
            for i in range(min(5, self.df.shape[0])):
                row_data = self.df.iloc[i, :min(10, self.df.shape[1])]
                row_str = f"{i+1:3d} │ "  # 顯示Excel行號，從1開始
                row_str += " │ ".join([str(val)[:8] for val in row_data])
                preview_text.insert("end", row_str + "\n")
            
            # 如果資料太多，顯示省略號
            if self.df.shape[0] > 5:
                preview_text.insert("end", "...(更多行省略)...")
            
            # 添加總行數和列數信息
            preview_text.insert("end", f"\n\n資料總行數: {self.df.shape[0]}, 總列數: {self.df.shape[1]}")
        except Exception as e:
            preview_text.insert("1.0", f"載入預覽資料時出錯: {str(e)}\n")
            import traceback
            preview_text.insert("end", traceback.format_exc())
        preview_text.config(state="disabled")
        
        # 確認按鈕
        def confirm_ranges():
            if not self.gui.selected_ranges:
                messagebox.showwarning("警告", "尚未選擇任何範圍", parent=range_dialog)
                return
            
            # 更新界面顯示
            range_strs = [r['range_str'] for r in self.gui.selected_ranges]
            self.gui.range_label.config(text=f"已選擇 {len(range_strs)} 個範圍: {', '.join(range_strs)}")
            
            # 更新為兼容原始邏輯的 selected_range
            self.gui.selected_range = self.gui.selected_ranges[0]  # 保留第一個範圍作為默認
            
            # 打印所選範圍的詳細信息
            self.gui.log(f"已確認 {len(self.gui.selected_ranges)} 個資料範圍:")
            for i, r in enumerate(self.gui.selected_ranges):
                self.gui.log(f"範圍 {i+1}: {r['range_str']}, 索引從 ({r['start_row']}, {r['start_col']}) 到 ({r['end_row']}, {r['end_col']})")
            
            # 啟用其他按鈕
            self.gui.view_data_btn.config(state="normal")
            self.gui.template_btn.config(state="normal")
            self.gui.template_combo.config(state="normal")
            
            range_dialog.destroy()
        
        ttk.Button(range_dialog, text="確認選擇範圍", command=confirm_ranges, width=20).pack(pady=10)
        
        # 綁定回車鍵到添加範圍
        range_entry.bind("<Return>", lambda event: add_range())
        
        # 設置初始焦點
        range_entry.focus_set()
        
    def get_column_letter(self, col_idx):
        """將數字列索引轉換為 Excel 列標記 (例如: 0->A, 1->B, 26->AA)"""
        result = ""
        while True:
            col_idx, remainder = divmod(col_idx, 26)
            result = chr(65 + remainder) + result
            if col_idx == 0:
                break
            col_idx -= 1
        return result
    
    def preview_data(self):
        """預覽所選範圍的數據"""
        if not hasattr(self.gui, 'selected_ranges') or not self.gui.selected_ranges or not self.gui.excel_files:
            return
        
        # 建立預覽窗口
        preview_window = tk.Toplevel(self.gui.root)
        preview_window.title("資料預覽")
        preview_window.geometry("700x500")
        
        # 頂部選擇區域
        top_frame = ttk.Frame(preview_window)
        top_frame.pack(fill="x", padx=10, pady=5)
        
        # 範圍選擇
        range_frame = ttk.Frame(top_frame)
        range_frame.pack(side="left", fill="x", expand=True)
        
        ttk.Label(range_frame, text="選擇範圍:").pack(side="left", padx=5)
        
        range_var = tk.StringVar()
        range_combo = ttk.Combobox(range_frame, textvariable=range_var, state="readonly", width=15)
        range_combo.pack(side="left", fill="x", expand=True, padx=5)
        
        range_strs = [r['range_str'] for r in self.gui.selected_ranges]
        range_combo['values'] = range_strs
        range_combo.current(0)
        
        # 文件選擇器 (如果有多個文件)
        file_frame = ttk.Frame(top_frame)
        file_frame.pack(side="right", fill="x", expand=True)
        
        ttk.Label(file_frame, text="選擇檔案:").pack(side="left", padx=5)
        
        file_var = tk.StringVar()
        file_combo = ttk.Combobox(file_frame, textvariable=file_var, state="readonly", width=40)
        file_combo.pack(side="left", fill="x", expand=True, padx=5)
        
        file_names = [os.path.basename(f) for f in self.gui.excel_files]
        file_combo['values'] = file_names
        file_combo.current(0)
        
        # 數據顯示區域
        preview_text = ScrolledText(preview_window, wrap=tk.WORD, font=("Courier New", 10))
        preview_text.pack(fill="both", expand=True, padx=10, pady=10)
        
        def update_preview():
            range_idx = range_combo.current()
            file_idx = file_combo.current()
            
            if range_idx >= 0 and range_idx < len(self.gui.selected_ranges) and file_idx >= 0 and file_idx < len(self.gui.excel_files):
                selected_range = self.gui.selected_ranges[range_idx]
                selected_file = self.gui.excel_files[file_idx]
                
                try:
                    df = self.gui.dfs[selected_file]
                    
                    # 提取所選範圍的數據
                    start_row = selected_range['start_row']
                    start_col = selected_range['start_col']
                    end_row = selected_range['end_row']
                    end_col = selected_range['end_col']
                    
                    # 打印調試信息
                    preview_text.delete("1.0", tk.END)  # 清空先前內容
                    preview_text.insert("1.0", f"嘗試讀取範圍: 從 ({start_row}, {start_col}) 到 ({end_row}, {end_col})\n")
                    preview_text.insert(tk.END, f"數據框形狀: {df.shape}\n\n")
                    
                    # 確保索引在有效範圍內
                    if start_row < 0 or start_col < 0 or start_row >= df.shape[0] or start_col >= df.shape[1]:
                        preview_text.insert(tk.END, f"錯誤: 起始位置 ({start_row}, {start_col}) 超出有效範圍 (0-{df.shape[0]-1}, 0-{df.shape[1]-1})\n")
                        return
                    
                    if end_row >= df.shape[0] or end_col >= df.shape[1]:
                        preview_text.insert(tk.END, f"警告: 結束位置 ({end_row}, {end_col}) 超出範圍，已調整到有效範圍\n")
                        end_row = min(end_row, df.shape[0]-1)
                        end_col = min(end_col, df.shape[1]-1)
                    
                    # 檢查是否為有效範圍
                    if start_row > end_row or start_col > end_col:
                        preview_text.insert(tk.END, "錯誤: 無效的範圍 (起始位置大於結束位置)\n")
                        return
                    
                    try:
                        selected_data = df.iloc[start_row:end_row+1, start_col:end_col+1]
                        
                        if selected_data.empty:
                            preview_text.insert(tk.END, "警告: 所選範圍內沒有數據\n")
                        else:
                            preview_text.insert(tk.END, selected_data.to_string())
                    except Exception as e:
                        preview_text.insert(tk.END, f"處理數據時出錯: {str(e)}\n")
                        import traceback
                        preview_text.insert(tk.END, traceback.format_exc())
                except Exception as e:
                    preview_text.delete("1.0", tk.END)
                    preview_text.insert("1.0", f"讀取檔案時出錯: {str(e)}\n")
        
        # 綁定選擇事件
        range_combo.bind("<<ComboboxSelected>>", lambda event: update_preview())
        file_combo.bind("<<ComboboxSelected>>", lambda event: update_preview())
        
        # 初始顯示
        update_preview()
        
        # 關閉按鈕
        ttk.Button(preview_window, text="關閉", command=preview_window.destroy).pack(pady=5)