# code_generator.py - 更新版本
import re
import pandas as pd
import numpy as np
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
from utils import format_cell_value

import threading

class CodeGenerator:
    def __init__(self, gui_instance):
        self.gui = gui_instance  # 保存對主GUI實例的引用
    
    def validate_template(self, template, selected_ranges=None):
        """
        Validate template syntax and check if all placeholders are supported
        驗證模板語法並檢查所有佔位符是否被支援
        """
        unsupported_tags = []
        # Check for unsupported template tags
        # 檢查不支援的模板標記
        pattern = r'{{([^}]+)}}'
        matches = re.findall(pattern, template)
        
        supported_tags = [
            'LOOP_START', 'LOOP_END', 'VALUE', 'ROW_INDEX', 'COL_INDEX',
            'ALL_COLUMNS', 'ALL_ROWS', 'DIRECTION:ROW', 'DIRECTION:COLUMN',
            'FILE_NAME', 'FILE_INDEX', 'FILE_COUNT', 'ROW_COUNT', 'COL_COUNT',
            'FILES_LOOP_START', 'FILES_LOOP_END', 'RANGES_LOOP_START', 'RANGES_LOOP_END',
            'RANGE_LOOP_START', 'RANGE_LOOP_END', 'RANGE_DATA_LOOP_START', 'RANGE_DATA_LOOP_END',
            'MAX_ROW_COUNT', 'MAX_COL_COUNT', 'RANGE_COUNT'
        ]
        
        for match in matches:
            is_supported = False
            
            # Check basic supported tags
            # 檢查基本支援的標記
            if any(tag in match for tag in supported_tags):
                is_supported = True
            # Check named range tags like RANGE[name]_LOOP_START
            # 檢查命名範圍標記
            elif re.match(r'RANGE\[.+\]_', match):
                is_supported = True
            # Check indexed tags like ROW:0, COL:1
            # 檢查索引標記
            elif re.match(r'(ROW|COL):\d+', match):
                is_supported = True
            # Check argument block tags
            # 檢查參數區塊標記
            elif re.match(r'ARGUMENT_(START|END):\w+', match):
                is_supported = True
            # Check numbered range tags like RANGE:1_LOOP_START
            # 檢查編號範圍標記
            elif re.match(r'RANGE:\d+_', match):
                is_supported = True
            
            if not is_supported:
                unsupported_tags.append(match)
        
        # Log validation results
        # 記錄驗證結果
        if unsupported_tags:
            self.gui.log(f"發現不支援的模板標記: {', '.join(unsupported_tags)}")
            return False, unsupported_tags
        else:
            self.gui.log("模板語法驗證通過")
            return True, []
    
    def load_template_from_file(self, file_path):
        """從檔案載入樣板內容"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                template_content = f.read()
            return template_content
        except Exception as e:
            messagebox.showerror("錯誤", f"無法讀取樣板檔案: {str(e)}")
            return None

    def set_template(self):
        """設置自訂程式碼樣板，並提供方向選擇"""
        # 建立樣板設定窗口
        template_dialog = tk.Toplevel(self.gui.root)
        template_dialog.title("設定程式碼樣板")
        template_dialog.geometry("700x700")  # 增加高度以容納更多說明
        template_dialog.grab_set()  # 模态窗口
        template_dialog.minsize(700, 700)  # 設定最小尺寸
        
        # 样板说明
        instruction_frame = ttk.Frame(template_dialog)
        instruction_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(instruction_frame, text="程式碼樣板說明：", font=("Arial", 10, "bold")).pack(anchor="w")
        instruction_text = """
在樣板中，您可以使用以下標記：

基本標記:
- {{LOOP_START}} - 資料迴圈開始 (使用第一個選擇的範圍)
- {{LOOP_END}} - 資料迴圈結束
- {{VALUE}} - 當前儲存格值
- {{ROW_INDEX}} - 當前資料列索引
- {{COL_INDEX}} - 當前資料行索引
- {{ALL_COLUMNS}} - 當前行的所有欄位值
- {{ALL_ROWS}} - 當前列的所有行值 (直向讀取模式)

方向控制標記:
- {{DIRECTION:ROW}} - 指定為橫向讀取模式 (預設)
- {{DIRECTION:COLUMN}} - 指定為直向讀取模式

多範圍精確標記:
- {{RANGE[範圍名稱]_LOOP_START}} - 指定範圍的資料迴圈開始
- {{RANGE[範圍名稱]_LOOP_END}} - 指定範圍的資料迴圈結束
- {{RANGE[範圍名稱]_ROW_COUNT}} - 指定範圍的資料列數
- {{RANGE[範圍名稱]_COL_COUNT}} - 指定範圍的資料行數
- {{RANGE[範圍名稱]_VALUE[行,列]}} - 指定範圍的特定儲存格值

檔案相關標記:
- {{FILE_NAME}} - 當前處理的檔案名稱
- {{FILE_INDEX}} - 當前檔案索引
- {{FILE_COUNT}} - 總檔案數量

範例: 
{{RANGE[左上]_VALUE[0,0]}} - 讀取名為"左上"的範圍中第一列第一行的值
{{RANGE[權重表]_LOOP_START}} - 開始迴圈處理名為"權重表"的範圍
"""
            
        instruction_box = ScrolledText(instruction_frame, height=15, font=("Courier New", 9))
        instruction_box.pack(fill="x", pady=5)
        instruction_box.insert("1.0", instruction_text)
        instruction_box.config(state="disabled")
        
        # 讀取方向選擇 (新增)
        direction_frame = ttk.LabelFrame(template_dialog, text="資料讀取方向")
        direction_frame.pack(fill="x", padx=10, pady=5)
        
        self.direction_var = tk.StringVar(value="row")  # 預設為橫向讀取
        
        ttk.Radiobutton(
            direction_frame, 
            text="橫向讀取 (Row by Row)",
            variable=self.direction_var,
            value="row"
        ).pack(side="left", padx=20, pady=5)
        
        ttk.Radiobutton(
            direction_frame, 
            text="直向讀取 (Column by Column)",
            variable=self.direction_var,
            value="column"
        ).pack(side="left", padx=20, pady=5)
        
        # 範圍定義區域
        range_frame = ttk.LabelFrame(template_dialog, text="範圍定義")
        range_frame.pack(fill="x", padx=10, pady=5)
        
        range_list_frame = ttk.Frame(range_frame)
        range_list_frame.pack(fill="x", padx=5, pady=5)
        
        # 顯示已定義的範圍
        ttk.Label(range_list_frame, text="已定義的範圍:").pack(side="left", padx=5)
        
        range_names = tk.StringVar()
        range_entry = ttk.Entry(range_list_frame, textvariable=range_names, width=40, state="readonly")
        range_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        # 初始化顯示已有的命名範圍
        if hasattr(self.gui, 'named_ranges'):
            range_names.set(", ".join(self.gui.named_ranges.keys()))
        
        # 範圍定義按鈕
        def define_range():
            # 建立範圍定義對話框
            range_dialog = tk.Toplevel(template_dialog)
            range_dialog.title("定義範圍")
            range_dialog.geometry("400x200")
            range_dialog.grab_set()
            
            ttk.Label(range_dialog, text="範圍名稱:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
            name_var = tk.StringVar()
            name_entry = ttk.Entry(range_dialog, textvariable=name_var, width=20)
            name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="we")
            
            ttk.Label(range_dialog, text="Excel 範圍 (例如: A1:G10):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
            range_var = tk.StringVar()
            range_entry = ttk.Entry(range_dialog, textvariable=range_var, width=20)
            range_entry.grid(row=1, column=1, padx=5, pady=5, sticky="we")
            
            # 定義範圍字典 (如果尚未存在)
            if not hasattr(self.gui, 'named_ranges'):
                self.gui.named_ranges = {}
            
            def save_range():
                name = name_var.get().strip()
                range_str = range_var.get().strip()
                
                if not name:
                    messagebox.showerror("錯誤", "請輸入範圍名稱", parent=range_dialog)
                    return
                    
                if not range_str:
                    messagebox.showerror("錯誤", "請輸入有效的Excel範圍", parent=range_dialog)
                    return
                    
                try:
                    # 驗證範圍格式
                    if ":" in range_str:
                        start, end = range_str.split(":")
                        # 這裡可以加入更多驗證邏輯
                    else:
                        messagebox.showerror("錯誤", "範圍格式不正確，請使用如 A1:G10 的格式", parent=range_dialog)
                        return
                    
                    # 存儲範圍定義
                    self.gui.named_ranges[name] = range_str
                    
                    # 更新範圍列表顯示
                    range_names.set(", ".join(self.gui.named_ranges.keys()))
                    
                    # 關閉對話框
                    range_dialog.destroy()
                    
                except Exception as e:
                    messagebox.showerror("錯誤", f"處理範圍時出錯: {str(e)}", parent=range_dialog)
            
            # 確認和取消按鈕
            btn_frame = ttk.Frame(range_dialog)
            btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
            
            ttk.Button(btn_frame, text="儲存範圍", command=save_range).pack(side="left", padx=5)
            ttk.Button(btn_frame, text="取消", command=range_dialog.destroy).pack(side="left", padx=5)
            
            # 設置初始焦點
            name_entry.focus_set()
        
        # 刪除範圍功能
        def delete_range():
            if not hasattr(self.gui, 'named_ranges') or not self.gui.named_ranges:
                messagebox.showinfo("提示", "沒有定義的範圍可刪除", parent=template_dialog)
                return
                
            # 建立範圍選擇對話框
            select_dialog = tk.Toplevel(template_dialog)
            select_dialog.title("選擇要刪除的範圍")
            select_dialog.geometry("300x200")
            select_dialog.grab_set()
            
            ttk.Label(select_dialog, text="選擇要刪除的範圍:").pack(padx=5, pady=5, anchor="w")
            
            # 建立範圍列表框
            range_listbox = tk.Listbox(select_dialog)
            range_listbox.pack(fill="both", expand=True, padx=5, pady=5)
            
            # 填充範圍列表
            for name in self.gui.named_ranges.keys():
                range_listbox.insert(tk.END, name)
            
            def confirm_delete():
                selected = range_listbox.curselection()
                if not selected:
                    messagebox.showinfo("提示", "請選擇要刪除的範圍", parent=select_dialog)
                    return
                    
                # 獲取選擇的範圍名稱
                selected_idx = selected[0]
                selected_name = range_listbox.get(selected_idx)
                
                # 確認刪除
                if messagebox.askyesno("確認", f"確定要刪除範圍 '{selected_name}' 嗎?", parent=select_dialog):
                    # 刪除範圍
                    del self.gui.named_ranges[selected_name]
                    
                    # 更新範圍列表顯示
                    range_names.set(", ".join(self.gui.named_ranges.keys()))
                    
                    # 關閉對話框
                    select_dialog.destroy()
            
            # 確認和取消按鈕
            btn_frame = ttk.Frame(select_dialog)
            btn_frame.pack(pady=10)
            
            ttk.Button(btn_frame, text="刪除", command=confirm_delete).pack(side="left", padx=5)
            ttk.Button(btn_frame, text="取消", command=select_dialog.destroy).pack(side="left", padx=5)
        
        # 範圍操作按鈕
        btn_frame = ttk.Frame(range_frame)
        btn_frame.pack(fill="x", pady=5)
        
        ttk.Button(btn_frame, text="定義新範圍", command=define_range).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="刪除範圍", command=delete_range).pack(side="left", padx=5)
        
        # 新增「從檔案載入」按鈕
        load_frame = ttk.Frame(template_dialog)
        load_frame.pack(fill="x", padx=10, pady=5)
        
        def load_template_file():
            file_path = filedialog.askopenfilename(
                title="選擇樣板檔案",
                filetypes=[("文字檔案", "*.txt"), ("C/C++檔案", "*.c;*.cpp;*.h"), ("所有檔案", "*.*")],
                parent=template_dialog
            )
            if file_path:
                template_content = self.load_template_from_file(file_path)
                if template_content:
                    template_text.delete("1.0", tk.END)
                    template_text.insert("1.0", template_content)
                    
        ttk.Button(load_frame, text="從檔案載入樣板", command=load_template_file).pack(side="left", padx=5)
        
        # 样板编辑区域标题
        ttk.Separator(template_dialog, orient="horizontal").pack(fill="x", padx=10, pady=5)
        ttk.Label(template_dialog, text="請在下方編輯程式碼樣板：", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=5)
        
        # 样板编辑区域
        template_frame = ttk.Frame(template_dialog)
        template_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        template_text = ScrolledText(template_frame, font=("Courier New", 10))
        template_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 如果已经有样板，就显示
        if self.gui.code_template:
            template_text.insert("1.0", self.gui.code_template)
        else:
            # 預設範例樣板 - 添加命名範圍支持和方向控制
            default_template = """// 資料處理範例 - 支援橫向與直向讀取
{{DIRECTION:ROW}}  // 設定為橫向讀取模式 (預設，可改為 {{DIRECTION:COLUMN}} 進行直向讀取)

#define LEFT_TOP_ROWS {{RANGE[左上]_ROW_COUNT}}
#define LEFT_TOP_COLS {{RANGE[左上]_COL_COUNT}}

// 左上角區域數據 - 橫向讀取 (Row by Row)
unsigned int left_top_area[LEFT_TOP_ROWS][LEFT_TOP_COLS] = {
{{RANGE[左上]_LOOP_START}}
    { {{ALL_COLUMNS}} },  // Row {{ROW_INDEX}}
{{RANGE[左上]_LOOP_END}}
};

// 特定儲存格值參考
int first_value = {{RANGE[左上]_VALUE[0,0]}};
int second_value = {{RANGE[左上]_VALUE[1,1]}};

// 使用 ALL_ROWS 示範直向讀取
void process_column_data() {
{{DIRECTION:COLUMN}}  // 切換為直向讀取
{{RANGE[左上]_LOOP_START}}
    int column_{{COL_INDEX}}_data[] = { {{ALL_ROWS}} };  // Column {{COL_INDEX}}
{{RANGE[左上]_LOOP_END}}
}
"""
            template_text.insert("1.0", default_template)
        
        # 按鈕區域
        btn_frame = ttk.Frame(template_dialog)
        btn_frame.pack(fill="x", padx=10, pady=(10, 20))  # 增加底部間距
        
        # 儲存到檔案按鈕
        def save_template_to_file():
            template_content = template_text.get("1.0", tk.END).strip()
            if not template_content:
                messagebox.showerror("錯誤", "樣板不能為空", parent=template_dialog)
                return
                
            file_path = filedialog.asksaveasfilename(
                title="儲存樣板至檔案",
                defaultextension=".txt",
                filetypes=[("文字檔案", "*.txt"), ("C檔案", "*.c"), ("C++檔案", "*.cpp"), ("所有檔案", "*.*")],
                parent=template_dialog
            )
            
            if file_path:
                try:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(template_content)
                    messagebox.showinfo("成功", f"樣板已儲存至 {file_path}", parent=template_dialog)
                except Exception as e:
                    messagebox.showerror("錯誤", f"儲存樣板時發生錯誤: {str(e)}", parent=template_dialog)
        
        # 确认按钮
        def confirm_template():
            template_content = template_text.get("1.0", tk.END).strip()
            if not template_content:
                messagebox.showerror("錯誤", "樣板不能為空", parent=template_dialog)
                return
            
            # Add template validation
            # 加入模板驗證
            is_valid, unsupported_tags = self.validate_template(template_content)
            if not is_valid:
                result = messagebox.askyesno(
                    "模板驗證警告", 
                    f"發現不支援的標記: {', '.join(unsupported_tags)}\n\n是否仍要繼續使用此模板？",
                    parent=template_dialog
                )
                if not result:
                    return
            
            # 設定模板内容
            self.gui.code_template = template_content
            
            # 記錄讀取方向設定
            self.gui.template_direction = self.direction_var.get()
            self.gui.log(f"已設定模板讀取方向: {self.gui.template_direction}")
            
            # 檢查樣板中的方向標記
            if "{{DIRECTION:ROW}}" in template_content:
                self.gui.template_direction = "row"
                self.gui.log("模板中指定橫向讀取模式")
            elif "{{DIRECTION:COLUMN}}" in template_content:
                self.gui.template_direction = "column"
                self.gui.log("模板中指定直向讀取模式")
            
            self.gui.template_preview.config(text="已設定自訂樣板")
            self.gui.generate_button.config(state="normal")
            template_dialog.destroy()
        
        # 儲存至檔案按鈕
        save_file_btn = ttk.Button(btn_frame, text="儲存至檔案", command=save_template_to_file, width=15)
        save_file_btn.pack(side="left", padx=5)
        
        # 放置明显的确认按钮
        confirm_btn = ttk.Button(btn_frame, text="確認樣板", command=confirm_template, width=15)
        confirm_btn.pack(side="right", padx=5)
        
        # 取消按钮
        cancel_btn = ttk.Button(btn_frame, text="取消", command=template_dialog.destroy, width=10)
        cancel_btn.pack(side="right", padx=5)
        
        # 在所有元件都加入後，使視窗自適應內容大小
        template_dialog.update_idletasks()
        # 重置視窗大小以適應所有內容
        template_dialog.geometry("")
    
    def get_default_template(self, template_name):
        """獲取預設樣板內容"""
        if template_name == "陣列初始化":
            return """// 一維陣列初始化
#define MAX_SIZE 100

unsigned int weights[MAX_SIZE] = {
{{LOOP_START}}
    {{VALUE}},  // 索引 {{ROW_INDEX}}
{{LOOP_END}}
};"""
        elif template_name == "簡化權重表設定":  # 新增簡化版本
            return """// 簡化權重表初始化
void initWeights() {
{{LOOP_START}}
    // 設定權重值 [行{{ROW_INDEX}}]: {{VALUE}}
    weights[{{ROW_INDEX}}] = {{VALUE}};
{{LOOP_END}}
}"""
        elif template_name == "二維陣列":
            return """// 二維陣列初始化
#define ROW_COUNT {{ROW_COUNT}}
#define COL_COUNT {{COL_COUNT}}

unsigned int table[ROW_COUNT][COL_COUNT] = {
{{LOOP_START}}
    { {{ALL_COLUMNS}} },
{{LOOP_END}}
};"""
        elif template_name == "二維陣列-直向讀取":
            # 加入直向讀取的新範本
            return """// 二維陣列初始化 - 直向讀取模式
{{DIRECTION:COLUMN}}  // 指定為直向讀取
#define COL_COUNT 20
#define ROW_COUNT 4

unsigned int table[ROW_COUNT][COL_COUNT] = {
{{LOOP_START}}
    { {{ALL_ROWS}} },  // Column {{COL_INDEX}}
{{LOOP_END}}
};"""
        elif template_name == "三維陣列":
            return """// 三維陣列初始化
#define FILE_COUNT {{FILE_COUNT}}
#define ROW_COUNT {{ROW_COUNT}}
#define COL_COUNT {{COL_COUNT}}

unsigned int data3d[FILE_COUNT][ROW_COUNT][COL_COUNT] = {
{{FILES_LOOP_START}}
    // 來自檔案: {{FILE_NAME}}
    {
{{LOOP_START}}
        { {{ALL_COLUMNS}} },
{{LOOP_END}}
    },
{{FILES_LOOP_END}}
};"""
        elif template_name == "三維陣列-直向讀取":
            # 加入直向讀取的三維陣列範本
            return """// 三維陣列初始化 - 直向讀取模式
{{DIRECTION:COLUMN}}  // 指定為直向讀取
#define FILE_COUNT {{FILE_COUNT}}
#define COL_COUNT {{ROW_COUNT}}  // 注意：ROW_COUNT 實際表示列數
#define ROW_COUNT {{COL_COUNT}}  // 注意：COL_COUNT 實際表示行數

unsigned int data3d[FILE_COUNT][ROW_COUNT][COL_COUNT] = {
{{FILES_LOOP_START}}
    // 來自檔案: {{FILE_NAME}}
    {
{{LOOP_START}}
        { {{ALL_ROWS}} },  // Column {{COL_INDEX}}
{{LOOP_END}}
    },
{{FILES_LOOP_END}}
};"""
        elif template_name == "四維陣列 (範圍優先)":
            return """// 四維陣列初始化 - 範圍優先 [範圍][檔案][行][列]
// 注意：此模板需要多個範圍和多個檔案才能正常運作
#define RANGE_COUNT {{RANGE_COUNT}}  // 範圍數量
#define FILE_COUNT {{FILE_COUNT}}    // 檔案數量
#define ROW_COUNT {{ROW_COUNT}}      // 每個範圍的最大行數
#define COL_COUNT {{COL_COUNT}}      // 每個範圍的最大列數

// 定義各個範圍的實際大小
unsigned int range_dimensions[RANGE_COUNT][2] = {
{{RANGES_LOOP_START}}
    { {{RANGE_ROW_COUNT}}, {{RANGE_COL_COUNT}} },  // 範圍 {{RANGE_INDEX}}: {{RANGE_STR}}
{{RANGES_LOOP_END}}
};

// 四維陣列: [範圍][檔案][行][列]
unsigned int data4d[RANGE_COUNT][FILE_COUNT][ROW_COUNT][COL_COUNT] = {
{{RANGES_LOOP_START}}
    // 範圍 {{RANGE_INDEX}}: {{RANGE_STR}}
    {
{{FILES_LOOP_START}}
        // 來自檔案: {{FILE_NAME}}
        {
{{RANGE_LOOP_START}}
            { {{ALL_COLUMNS}} },
{{RANGE_LOOP_END}}
        },
{{FILES_LOOP_END}}
    },
{{RANGES_LOOP_END}}
};

// 範例：取得特定範圍、特定檔案的資料
unsigned int get_value(unsigned int range_idx, unsigned int file_idx, unsigned int row, unsigned int col) {
    // 邊界檢查
    if (range_idx >= RANGE_COUNT || file_idx >= FILE_COUNT || 
        row >= range_dimensions[range_idx][0] || col >= range_dimensions[range_idx][1]) {
        return 0;  // 超出範圍，返回預設值
    }
    
    return data4d[range_idx][file_idx][row][col];
}"""
        elif template_name == "四維陣列-直向讀取":
            # 加入直向讀取的四維陣列範本
            return """// 四維陣列初始化 - 直向讀取模式
{{DIRECTION:COLUMN}}  // 指定為直向讀取
#define RANGE_COUNT {{RANGE_COUNT}}  // 範圍數量
#define FILE_COUNT {{FILE_COUNT}}    // 檔案數量
#define COL_COUNT {{ROW_COUNT}}      // 每個範圍的最大列數
#define ROW_COUNT {{COL_COUNT}}      // 每個範圍的最大行數

// 定義各個範圍的實際大小
unsigned int range_dimensions[RANGE_COUNT][2] = {
{{RANGES_LOOP_START}}
    { {{RANGE_COL_COUNT}}, {{RANGE_ROW_COUNT}} },  // 範圍 {{RANGE_INDEX}}: {{RANGE_STR}}
{{RANGES_LOOP_END}}
};

// 四維陣列: [範圍][檔案][行][列]
unsigned int data4d[RANGE_COUNT][FILE_COUNT][ROW_COUNT][COL_COUNT] = {
{{RANGES_LOOP_START}}
    // 範圍 {{RANGE_INDEX}}: {{RANGE_STR}}
    {
{{FILES_LOOP_START}}
        // 來自檔案: {{FILE_NAME}}
        {
{{RANGE_LOOP_START}}
            { {{ALL_ROWS}} },  // Column {{COL_INDEX}}
{{RANGE_LOOP_END}}
        },
{{FILES_LOOP_END}}
    },
{{RANGES_LOOP_END}}
};"""
        elif template_name == "四維陣列 (檔案優先)":
            return """// 四維陣列初始化 - 檔案優先 [檔案][範圍][行][列]
#define FILE_COUNT {{FILE_COUNT}}     // 檔案數量
#define RANGE_COUNT {{RANGE_COUNT}}   // 範圍數量
#define ROW_COUNT {{ROW_COUNT}}       // 每個範圍的最大行數
#define COL_COUNT {{COL_COUNT}}       // 每個範圍的最大列數

// 定義各個範圍的實際大小
unsigned int range_dimensions[RANGE_COUNT][2] = {
{{RANGES_LOOP_START}}
    { {{RANGE_ROW_COUNT}}, {{RANGE_COL_COUNT}} },  // 範圍 {{RANGE_INDEX}}: {{RANGE_STR}}
{{RANGES_LOOP_END}}
};

// 四維陣列: [檔案][範圍][行][列]
unsigned int data4d[FILE_COUNT][RANGE_COUNT][ROW_COUNT][COL_COUNT] = {
{{FILES_LOOP_START}}
    // 來自檔案: {{FILE_NAME}}
    {
{{RANGES_LOOP_START}}
        // 範圍 {{RANGE_INDEX}}: {{RANGE_STR}}
        {
{{RANGE_LOOP_START}}
            { {{ALL_COLUMNS}} },
{{RANGE_LOOP_END}}
        },
{{RANGES_LOOP_END}}
    },
{{FILES_LOOP_END}}
};"""
        elif template_name == "三維多範圍陣列":
            return """// 三維多範圍陣列初始化
#define FILE_COUNT {{FILE_COUNT}}           // 檔案數量
#define RANGE_COUNT {{RANGE_COUNT}}         // 範圍數量
#define MAX_ROW_COUNT {{MAX_ROW_COUNT}}     // 最大行數
#define MAX_COL_COUNT {{MAX_COL_COUNT}}     // 最大列數

// 定義各個範圍的實際大小
unsigned int range_dimensions[RANGE_COUNT][2] = {
{{RANGES_LOOP_START}}
    { {{RANGE_ROW_COUNT}}, {{RANGE_COL_COUNT}} },  // 範圍 {{RANGE_INDEX}}: {{RANGE_STR}}
{{RANGES_LOOP_END}}
};

// 三維多範圍陣列: [檔案][範圍][行][列] (使用平坦化儲存)
unsigned int* data3d_multi[FILE_COUNT][RANGE_COUNT];

// 初始化函數
void init_3d_multi_array() {
{{FILES_LOOP_START}}
    // 檔案 {{FILE_INDEX}}: {{FILE_NAME}}
{{RANGES_LOOP_START}}
    // 範圍 {{RANGE_INDEX}}: {{RANGE_STR}}
    data3d_multi[{{FILE_INDEX}}][{{RANGE_INDEX}}] = malloc({{RANGE_ROW_COUNT}} * {{RANGE_COL_COUNT}} * sizeof(unsigned int));
    
    // 初始化資料
    unsigned int temp_data[] = {
{{RANGE_DATA_LOOP_START}}
        {{ALL_COLUMNS}},
{{RANGE_DATA_LOOP_END}}
    };
    
    memcpy(data3d_multi[{{FILE_INDEX}}][{{RANGE_INDEX}}], temp_data, sizeof(temp_data));
{{RANGES_LOOP_END}}
{{FILES_LOOP_END}}
}"""
        elif template_name == "權重表設定":
            return """// 遊戲權重表初始化
void initVariableWeights() {
{{LOOP_START}}
    normal_table_weight[{{COL:0}}][{{COL:1}}][{{COL:2}}][{{COL:3}}] = {{VALUE}};
{{LOOP_END}}
}"""
        elif template_name == "權重表設定-直向讀取":
            # 加入直向讀取的權重表範本
            return """// 遊戲權重表初始化 - 直向讀取模式
{{DIRECTION:COLUMN}}  // 指定為直向讀取
void initVariableWeightsByColumn() {
{{LOOP_START}}
    // 設定第 {{COL_INDEX}} 列的所有權重
    int column_{{COL_INDEX}}_weights[] = { {{ALL_ROWS}} };
    for (int i = 0; i < sizeof(column_{{COL_INDEX}}_weights)/sizeof(int); i++) {
        column_weights[{{COL_INDEX}}][i] = column_{{COL_INDEX}}_weights[i];
    }
{{LOOP_END}}
}"""
        elif template_name == "多範圍處理":
            return """// 多範圍數據處理
#define RANGE_1_ROW_COUNT {{RANGE_1_ROW_COUNT}}
#define RANGE_1_COL_COUNT {{RANGE_1_COL_COUNT}}
#define RANGE_2_ROW_COUNT {{RANGE_2_ROW_COUNT}}
#define RANGE_2_COL_COUNT {{RANGE_2_COL_COUNT}}

// 第一個範圍的數據
unsigned int first_area[RANGE_1_ROW_COUNT][RANGE_1_COL_COUNT] = {
{{RANGE:1_LOOP_START}}
    { {{ALL_COLUMNS}} },  // Row {{ROW_INDEX}}
{{RANGE:1_LOOP_END}}
};

// 第二個範圍的數據
unsigned int second_area[RANGE_2_ROW_COUNT][RANGE_2_COL_COUNT] = {
{{RANGE:2_LOOP_START}}
    { {{ALL_COLUMNS}} },  // Row {{ROW_INDEX}}
{{RANGE:2_LOOP_END}}
};"""
        elif template_name == "命名範圍處理":
            return """// 命名範圍資料處理範例
#define LEFT_TOP_ROWS {{RANGE[左上]_ROW_COUNT}}
#define LEFT_TOP_COLS {{RANGE[左上]_COL_COUNT}}
#define RIGHT_TOP_ROWS {{RANGE[右上]_ROW_COUNT}}
#define RIGHT_TOP_COLS {{RANGE[右上]_COL_COUNT}}

// 左上角區域數據
unsigned int left_top_area[LEFT_TOP_ROWS][LEFT_TOP_COLS] = {
{{RANGE[左上]_LOOP_START}}
    { {{ALL_COLUMNS}} },  // Row {{ROW_INDEX}}
{{RANGE[左上]_LOOP_END}}
};

// 右上角區域數據
unsigned int right_top_area[RIGHT_TOP_ROWS][RIGHT_TOP_COLS] = {
{{RANGE[右上]_LOOP_START}}
    { {{ALL_COLUMNS}} },  // Row {{ROW_INDEX}}
{{RANGE[右上]_LOOP_END}}
};

// 特定欄位值示例
int left_top_first_value = {{RANGE[左上]_VALUE[0,0]}};
int right_top_first_value = {{RANGE[右上]_VALUE[0,0]}};
"""
        else:
            return ""

    def convert_range_notation_to_indices(self, range_name):
        """
        將範圍名稱轉換為索引資訊
        
        Args:
            range_name (str): 範圍名稱
        
        Returns:
            tuple: (start_row, start_col, end_row, end_col) 或 None 如果範圍不存在
        """
        # 檢查範圍是否存在
        if not hasattr(self.gui, 'named_ranges') or range_name not in self.gui.named_ranges:
            self.gui.log(f"警告: 未找到命名範圍 '{range_name}'")
            return None
        
        # 獲取範圍
        range_str = self.gui.named_ranges[range_name]
        
        try:
            # 解析範圍
            start, end = range_str.split(":")
            from utils import excel_notation_to_index
            start_row, start_col = excel_notation_to_index(start, self.gui)
            end_row, end_col = excel_notation_to_index(end, self.gui)

            self.gui.log(f"範圍 {range_name}: 原始輸入 {range_str}")
            self.gui.log(f"範圍 {range_name}: 轉換後索引 start_row={start_row}, start_col={start_col}, end_row={end_row}, end_col={end_col}")
            
            return (start_row, start_col, end_row, end_col)
        except Exception as e:
            self.gui.log(f"解析範圍 '{range_name}' 時出錯: {str(e)}")
            return None

    def process_named_range_value(self, template, dfs, excel_files):
        """處理模板中的命名範圍特定值引用"""
        # 正則表達式匹配所有命名範圍的值引用，例如 {{RANGE[範圍名]_VALUE[0,0]}}
        value_pattern = r'{{RANGE\[([^\]]+)\]_VALUE\[(\d+),(\d+)\]}}'
        matches = re.findall(value_pattern, template)
        
        result = template
        
        for range_name, row_idx, col_idx in matches:
            placeholder = f"{{{{RANGE[{range_name}]_VALUE[{row_idx},{col_idx}]}}}}"
            
            # 檢查命名範圍是否存在
            if not hasattr(self.gui, 'named_ranges') or range_name not in self.gui.named_ranges:
                # 範圍未定義，保留原始標記
                self.gui.log(f"警告: 未找到命名範圍 '{range_name}'")
                continue
                
            # 獲取範圍
            range_indices = self.convert_range_notation_to_indices(range_name)
            if not range_indices:
                continue
            
            start_row, start_col, end_row, end_col = range_indices
            
            try:
                # 取得目標儲存格的相對位置
                target_row = start_row + int(row_idx)
                target_col = start_col + int(col_idx)
                
                # 檢查位置是否在範圍內
                if target_row > end_row or target_col > end_col:
                    self.gui.log(f"警告: 位置 [{row_idx},{col_idx}] 超出範圍 '{range_name}' 的界限")
                    continue
                
                # 使用第一個文件的資料
                first_file = excel_files[0]
                df = dfs[first_file]
                
                # 獲取目標值
                if target_row < df.shape[0] and target_col < df.shape[1]:
                    value = df.iloc[target_row, target_col]
                    formatted_value = format_cell_value(value)
                    result = result.replace(placeholder, formatted_value)
                else:
                    self.gui.log(f"警告: 位置 [{target_row},{target_col}] 超出資料範圍")
            except Exception as e:
                self.gui.log(f"處理命名範圍值時出錯: {str(e)}")
        
        return result

    def process_named_range_metadata(self, template):
        """
        處理模板中的命名範圍元數據，如行數、列數和範圍全名
        
        Args:
            template (str): 包含命名範圍標記的模板字符串
            
        Returns:
            str: 處理後的模板字符串，替換了相關標記
        """
        result = template
        
        # 處理行數
        row_count_pattern = r'{{RANGE\[([^\]]+)\]_ROW_COUNT}}'
        matches = re.findall(row_count_pattern, result)
        
        for range_name in matches:
            placeholder = f"{{{{RANGE[{range_name}]_ROW_COUNT}}}}"
            
            # 檢查命名範圍是否存在
            range_indices = self.convert_range_notation_to_indices(range_name)
            if not range_indices:
                continue
                
            start_row, start_col, end_row, end_col = range_indices
            
            try:
                # 計算行數
                row_count = end_row - start_row + 1
                
                # 替換標記
                result = result.replace(placeholder, str(row_count))
            except Exception as e:
                self.gui.log(f"處理命名範圍 {range_name} 行數時出錯: {str(e)}")
        
        # 處理列數
        col_count_pattern = r'{{RANGE\[([^\]]+)\]_COL_COUNT}}'
        matches = re.findall(col_count_pattern, result)
        
        for range_name in matches:
            placeholder = f"{{{{RANGE[{range_name}]_COL_COUNT}}}}"
            
            # 檢查命名範圍是否存在
            range_indices = self.convert_range_notation_to_indices(range_name)
            if not range_indices:
                continue
                
            start_row, start_col, end_row, end_col = range_indices
            
            try:
                # 計算列數
                col_count = end_col - start_col + 1
                
                # 替換標記
                result = result.replace(placeholder, str(col_count))
            except Exception as e:
                self.gui.log(f"處理命名範圍 {range_name} 列數時出錯: {str(e)}")
        
        # 處理範圍全名 (新增功能)
        full_name_pattern = r'{{RANGE\[([^\]]+)\]_FULL_NAME}}'
        matches = re.findall(full_name_pattern, result)
        
        for range_name in matches:
            placeholder = f"{{{{RANGE[{range_name}]_FULL_NAME}}}}"
            
            # 檢查命名範圍是否存在
            if not hasattr(self.gui, 'named_ranges') or range_name not in self.gui.named_ranges:
                self.gui.log(f"警告: 未找到命名範圍 '{range_name}'")
                result = result.replace(placeholder, range_name)
                continue
            
            try:
                # 獲取實際範圍字串
                range_str = self.gui.named_ranges[range_name]
                full_name = f"{range_name} ({range_str})"
                
                # 替換標記
                result = result.replace(placeholder, full_name)
            except Exception as e:
                self.gui.log(f"處理命名範圍 {range_name} 全名時出錯: {str(e)}")
                result = result.replace(placeholder, range_name)
        
        return result

    def process_named_range_loops(self, template, dfs, excel_files):
        """處理模板中的命名範圍循環"""
        # 找出所有命名範圍循環
        loop_pattern = r'{{RANGE\[([^\]]+)\]_LOOP_START}}(.*?){{RANGE\[(\1)\]_LOOP_END}}'
        matches = re.findall(loop_pattern, template, re.DOTALL)
        
        result = template
        
        for range_name, loop_content, _ in matches:
            start_tag = f"{{{{RANGE[{range_name}]_LOOP_START}}}}"
            end_tag = f"{{{{RANGE[{range_name}]_LOOP_END}}}}"
            full_pattern = f"{start_tag}{loop_content}{end_tag}"
            
            # 檢查命名範圍是否存在
            range_indices = self.convert_range_notation_to_indices(range_name)
            if not range_indices:
                continue
                
            start_row, start_col, end_row, end_col = range_indices
            
            try:
                # 使用第一個文件的資料
                first_file = excel_files[0]
                df = dfs[first_file]
                
                # 提取所選範圍的數據
                selected_data = df.iloc[start_row:end_row+1, start_col:end_col+1]
                
                # 檢查是否為直向讀取模式
                is_column_mode = self.check_direction_mode(loop_content)
                
                # 生成循環內容
                loop_result = []
                
                if is_column_mode:
                    # 直向讀取模式 - 按列處理
                    for col_idx in range(selected_data.shape[1]):
                        column_data = selected_data.iloc[:, col_idx]
                        line = self.process_column_data(column_data, loop_content, start_col, col_idx, selected_data.shape[1])
                        loop_result.append(line)
                else:
                    # 橫向讀取模式 - 按行處理
                    for row_idx in range(selected_data.shape[0]):
                        row_data = selected_data.iloc[row_idx, :]
                        line = self.process_row_data(row_data, loop_content, start_row, row_idx, selected_data.shape[0])
                        loop_result.append(line)
                
                # 替換整個循環區塊
                result = result.replace(full_pattern, "".join(loop_result))
                
            except Exception as e:
                self.gui.log(f"處理命名範圍循環時出錯: {str(e)}")
                import traceback
                self.gui.log(traceback.format_exc())
        
        return result

    def check_direction_mode(self, template_section):
        """
        檢查模板片段中是否指定了直向讀取模式
        
        Args:
            template_section (str): 模板代碼片段
            
        Returns:
            bool: 如果是直向讀取模式返回 True，否則返回 False
        """
        # 檢查模板片段中是否含有直向讀取標記
        if "{{DIRECTION:COLUMN}}" in template_section:
            return True
            
        # 如果沒有明確指定，檢查全域設定
        if hasattr(self.gui, 'template_direction') and self.gui.template_direction == "column":
            return True
            
        # 預設為橫向讀取
        return False

    def process_range_data_loop(self, content, df, range_indices, is_column_mode=False):
        """
        Process RANGE_DATA_LOOP tags in template content
        處理模板中的 RANGE_DATA_LOOP 標記
        """
        if "{{RANGE_DATA_LOOP_START}}" not in content:
            return content
            
        start_row, start_col, end_row, end_col = range_indices
        selected_data = df.iloc[start_row:end_row+1, start_col:end_col+1]
        
        # Split template to get loop content
        # 分割模板以獲取循環內容
        parts = content.split("{{RANGE_DATA_LOOP_START}}")
        before_loop = parts[0]
        
        loop_and_after = parts[1].split("{{RANGE_DATA_LOOP_END}}")
        loop_content = loop_and_after[0]
        after_loop = loop_and_after[1] if len(loop_and_after) > 1 else ""
        
        loop_result = []
        
        if is_column_mode:
            # Column-wise reading
            # 直向讀取
            for col_idx in range(selected_data.shape[1]):
                column_data = selected_data.iloc[:, col_idx]
                line = self.process_column_data(column_data, loop_content, start_col, col_idx, selected_data.shape[1])
                loop_result.append(line)
        else:
            # Row-wise reading
            # 橫向讀取
            for row_idx in range(selected_data.shape[0]):
                row_data = selected_data.iloc[row_idx, :]
                line = self.process_row_data(row_data, loop_content, start_row, row_idx, selected_data.shape[0])
                loop_result.append(line)
        
        return before_loop + "".join(loop_result) + after_loop

    def generate_code(self, excel_files, dfs, selected_ranges, code_template, selected_range):
        """生成程式碼"""
        # 首先檢查是否有檔案和範圍
        if not excel_files:
            messagebox.showerror("錯誤", "請先選擇文件和資料範圍")
            return code_template
        
        # 移除所有方向控制標記，但記住最後的設定
        is_column_mode = "{{DIRECTION:COLUMN}}" in code_template
        template = code_template.replace("{{DIRECTION:ROW}}", "")
        template = template.replace("{{DIRECTION:COLUMN}}", "")
        
        # 如果全域有方向設定，使用它
        if hasattr(self.gui, 'template_direction'):
            is_column_mode = self.gui.template_direction == "column"
        
        self.gui.log(f"讀取方向: {'直向(Column)' if is_column_mode else '橫向(Row)'}")
        
        # 處理檔案數量
        template = template.replace("{{FILE_COUNT}}", str(len(excel_files)))
        
        # 計算最大行數和列數（通用方法）
        if selected_ranges:
            max_row_count = max(
                range_info['end_row'] - range_info['start_row'] + 1 
                for range_info in selected_ranges
            )
            
            max_col_count = max(
                range_info['end_col'] - range_info['start_col'] + 1 
                for range_info in selected_ranges
            )
            
            template = template.replace("{{MAX_ROW_COUNT}}", str(max_row_count))
            template = template.replace("{{MAX_COL_COUNT}}", str(max_col_count))
            template = template.replace("{{RANGE_COUNT}}", str(len(selected_ranges)))
            
            # 處理第一個範圍的行列數（向下相容）
            first_range = selected_ranges[0]
            template = template.replace("{{ROW_COUNT}}", str(first_range['end_row'] - first_range['start_row'] + 1))
            template = template.replace("{{COL_COUNT}}", str(first_range['end_col'] - first_range['start_col'] + 1))
        
        # 處理命名範圍的行列數
        template = self.process_named_range_metadata(template)
        
        # 處理命名範圍特定值引用（在任何循環處理之前）
        template = self.process_named_range_value(template, dfs, excel_files)
        
        # 使用正則表達式找出所有參數區塊
        argument_pattern = r'{{ARGUMENT_START:(\w+)}}(.*?){{ARGUMENT_END:\1}}'
        arguments = re.findall(argument_pattern, template, re.DOTALL)

        # 處理參數區塊
        for argument_name, argument_content in arguments:
            # 從備註中提取範圍名稱
            range_match = re.search(r'範圍名稱=([^\n]+)', argument_content)
            range_names = []

            if range_match:
                range_names = [name.strip() for name in range_match.group(1).split(',')]
                
            # 處理這個參數的內容
            processed_argument = self.process_argument(
                argument_content, 
                excel_files, 
                dfs, 
                range_names, 
                is_column_mode
            )
            
            # 替換原始內容
            template = template.replace(
                f'{{{{ARGUMENT_START:{argument_name}}}}}' + argument_content + f'{{{{ARGUMENT_END:{argument_name}}}}}', 
                processed_argument
            )

        # 檢查不配對的參數區塊標籤，修復可能的錯誤
        mismatch_pattern = r'{{ARGUMENT_START:(\w+)}}(.*?){{ARGUMENT_END:(\w+)}}'
        mismatch_arguments = re.findall(mismatch_pattern, template, re.DOTALL)
        
        for start_name, content, end_name in mismatch_arguments:
            if start_name != end_name:  # 檢測到標籤不匹配
                self.gui.log(f"警告: 參數區塊標籤不匹配 - 開始: {start_name}, 結束: {end_name}")
                # 嘗試處理該區塊
                range_match = re.search(r'範圍名稱=([^\n]+)', content)
                if range_match:
                    range_names = [name.strip() for name in range_match.group(1).split(',')]
                    
                    processed_argument = self.process_argument(
                        content, 
                        excel_files, 
                        dfs, 
                        range_names, 
                        is_column_mode
                    )
                    
                    # 替換原始內容，使用檢測到的不匹配標籤
                    template = template.replace(
                        f'{{{{ARGUMENT_START:{start_name}}}}}' + content + f'{{{{ARGUMENT_END:{end_name}}}}}', 
                        processed_argument
                    )

        # 處理參數區塊外的傳統標記
        template = self.process_traditional_template(
            template, 
            excel_files, 
            dfs, 
            selected_ranges,
            selected_range,
            is_column_mode
        )

        return template

    def process_traditional_template(self, template, excel_files, dfs, selected_ranges, selected_range, is_column_mode):
        """處理參數區塊外的傳統標記"""
        final_code = template
        
        # 如果沒有任何需要處理的標記，直接返回
        traditional_markers = [
            "{{LOOP_START}}", "{{RANGE[", "{{RANGE:", "{{FILES_LOOP_START}}", "{{RANGES_LOOP_START}}"
        ]
        if not any(marker in final_code for marker in traditional_markers):
            return final_code
        
        self.gui.log("處理參數區塊外的傳統標記...")
        
        # 處理命名範圍循環
        final_code = self.process_named_range_loops(final_code, dfs, excel_files)
        
        # 判斷模板類型並相應處理
        if "{{RANGES_LOOP_START}}" in final_code and "{{FILES_LOOP_START}}" in final_code:
            # 四維陣列模板 - 範圍優先
            self.gui.log("檢測到四維陣列模板（範圍優先）")
            if selected_ranges:
                final_code = self.process_4d_range_first_template(
                    final_code,
                    excel_files,
                    dfs,
                    selected_ranges,
                    len(excel_files),
                    max(r['end_row'] - r['start_row'] + 1 for r in selected_ranges),
                    max(r['end_col'] - r['start_col'] + 1 for r in selected_ranges),
                    is_column_mode
                )
        elif "{{FILES_LOOP_START}}" in final_code and "{{RANGES_LOOP_START}}" in final_code:
            # 四維陣列模板 - 檔案優先
            self.gui.log("檢測到四維陣列模板（檔案優先）")
            if selected_ranges:
                final_code = self.process_4d_file_first_template(
                    final_code,
                    excel_files,
                    dfs,
                    selected_ranges,
                    len(excel_files),
                    max(r['end_row'] - r['start_row'] + 1 for r in selected_ranges),
                    max(r['end_col'] - r['start_col'] + 1 for r in selected_ranges),
                    is_column_mode
                )
        elif "{{FILES_LOOP_START}}" in final_code and "{{RANGE_DATA_LOOP_START}}" in final_code:
            # 三維多範圍陣列模板
            self.gui.log("檢測到三維多範圍陣列模板")
            final_code = self.process_3d_multi_range_template(
                final_code,
                excel_files,
                dfs,
                selected_ranges,
                len(excel_files),
                is_column_mode
            )
        elif "{{FILES_LOOP_START}}" in final_code and "{{LOOP_START}}" in final_code:
            # 三維陣列模板
            self.gui.log("檢測到三維陣列模板")
            if selected_ranges:
                first_range = selected_ranges[0]
                final_code = self.process_3d_template(
                    final_code,
                    excel_files,
                    dfs,
                    first_range['start_row'],
                    first_range['start_col'],
                    first_range['end_row'],
                    first_range['end_col'],
                    len(excel_files),
                    first_range['end_row'] - first_range['start_row'] + 1,
                    is_column_mode
                )
        elif "{{RANGE:1_LOOP_START}}" in final_code or "{{RANGE:2_LOOP_START}}" in final_code:
            # 多範圍處理模板
            self.gui.log("檢測到多範圍處理模板")
            final_code = self.process_multi_range_template(
                final_code, 
                excel_files, 
                dfs, 
                selected_ranges, 
                is_column_mode
            )
        elif "{{LOOP_START}}" in final_code and "{{LOOP_END}}" in final_code:
            # 標準單範圍模板
            self.gui.log("檢測到標準單範圍模板")
            if selected_ranges:
                first_range = selected_ranges[0]
                final_code = self.process_standard_template(
                    final_code,
                    excel_files,
                    dfs,
                    first_range['start_row'],
                    first_range['start_col'],
                    first_range['end_row'],
                    first_range['end_col'],
                    first_range['end_row'] - first_range['start_row'] + 1,
                    is_column_mode
                )
            elif selected_range:
                # 向下相容：使用舊的 selected_range
                final_code = self.process_standard_template(
                    final_code,
                    excel_files,
                    dfs,
                    selected_range['start_row'],
                    selected_range['start_col'],
                    selected_range['end_row'],
                    selected_range['end_col'],
                    selected_range['end_row'] - selected_range['start_row'] + 1,
                    is_column_mode
                )
        
        return final_code

    def process_argument(self, template, excel_files, dfs, range_names, is_column_mode):
        """Process a specific argument block with its ranges"""
        final_code = template
        
        # Process file loops
        # If no FILES_LOOP_START, use the first file
        if "{{FILES_LOOP_START}}" not in final_code and "{{FILES_LOOP_END}}" not in final_code:
            # Use data from the first file
            if excel_files:
                file_path = excel_files[0]
                df = dfs[file_path]
                
                # Process named range loops
                for range_name in range_names:
                    range_loop_start = f"{{{{RANGE[{range_name}]_LOOP_START}}}}"
                    range_loop_end = f"{{{{RANGE[{range_name}]_LOOP_END}}}}"
                    
                    if range_loop_start in final_code:
                        # Get range information using convert_range_notation_to_indices
                        range_indices = self.convert_range_notation_to_indices(range_name)
                        
                        if range_indices:
                            start_row, start_col, end_row, end_col = range_indices
                            self.gui.log(f"處理引數範圍 {range_name}: {start_row}:{end_row}, {start_col}:{end_col}")
                            
                            # Extract data for the selected range
                            selected_data = df.iloc[start_row:end_row+1, start_col:end_col+1]
                            
                            # Split template to get loop content
                            parts = final_code.split(range_loop_start)
                            before_loop = parts[0]
                            
                            loop_and_after = parts[1].split(range_loop_end)
                            loop_content = loop_and_after[0]
                            after_loop = loop_and_after[1] if len(loop_and_after) > 1 else ""
                            
                            loop_result = []
                            
                            # Check if loop content specifies read direction
                            local_is_column_mode = self.check_direction_mode(loop_content) or is_column_mode
                            
                            if local_is_column_mode:
                                # Column-wise reading
                                for col_idx in range(selected_data.shape[1]):
                                    column_data = selected_data.iloc[:, col_idx]
                                    line = self.process_column_data(column_data, loop_content, start_col, col_idx, selected_data.shape[1])
                                    loop_result.append(line)
                            else:
                                # Row-wise reading
                                for row_idx in range(selected_data.shape[0]):
                                    row_data = selected_data.iloc[row_idx, :]
                                    line = self.process_row_data(row_data, loop_content, start_row, row_idx, selected_data.shape[0])
                                    loop_result.append(line)
                            
                            # Replace loop content
                            final_code = before_loop + "".join(loop_result) + after_loop
                
                # Process RANGE_DATA_LOOP if present
                # 處理 RANGE_DATA_LOOP（如果存在）
                for range_name in range_names:
                    if "{{RANGE_DATA_LOOP_START}}" in final_code:
                        range_indices = self.convert_range_notation_to_indices(range_name)
                        if range_indices:
                            final_code = self.process_range_data_loop(
                                final_code, df, range_indices, is_column_mode
                            )
                            break  # 只處理第一個找到的範圍
                
                # If there's still a standard loop, process it
                if "{{LOOP_START}}" in final_code:
                    final_code = self.process_standard_template(
                        final_code, 
                        [file_path], 
                        {file_path: df}, 
                        0, 0, 
                        df.shape[0] - 1, 
                        df.shape[1] - 1, 
                        df.shape[0], 
                        is_column_mode
                    )
        else:
            # Process file loops when FILES_LOOP_START and FILES_LOOP_END are present
            parts = final_code.split("{{FILES_LOOP_START}}")
            before_files_loop = parts[0]
            
            files_loop_and_after = parts[1].split("{{FILES_LOOP_END}}")
            files_loop_content = files_loop_and_after[0]
            after_files_loop = files_loop_and_after[1] if len(files_loop_and_after) > 1 else ""
            
            files_result = []
            
            # Process each file
            for file_idx, file_path in enumerate(excel_files):
                df = dfs[file_path]
                file_content = files_loop_content.replace("{{FILE_INDEX}}", str(file_idx))
                file_content = file_content.replace("{{FILE_NAME}}", os.path.basename(file_path))
                
                # Process each named range within this file
                for range_name in range_names:
                    range_loop_start = f"{{{{RANGE[{range_name}]_LOOP_START}}}}"
                    range_loop_end = f"{{{{RANGE[{range_name}]_LOOP_END}}}}"
                    
                    if range_loop_start in file_content:
                        range_indices = self.convert_range_notation_to_indices(range_name)
                        if range_indices:
                            start_row, start_col, end_row, end_col = range_indices
                            
                            selected_data = df.iloc[start_row:end_row+1, start_col:end_col+1]
                            
                            # Split template to get range loop content
                            range_parts = file_content.split(range_loop_start)
                            before_range_loop = range_parts[0]
                            
                            range_loop_and_after = range_parts[1].split(range_loop_end)
                            range_loop_content = range_loop_and_after[0]
                            after_range_loop = range_loop_and_after[1] if len(range_loop_and_after) > 1 else ""
                            
                            loop_result = []
                            
                            local_is_column_mode = self.check_direction_mode(range_loop_content) or is_column_mode
                            
                            if local_is_column_mode:
                                # Column-wise reading
                                for col_idx in range(selected_data.shape[1]):
                                    column_data = selected_data.iloc[:, col_idx]
                                    line = self.process_column_data(column_data, range_loop_content, start_col, col_idx, selected_data.shape[1])
                                    loop_result.append(line)
                            else:
                                # Row-wise reading
                                for row_idx in range(selected_data.shape[0]):
                                    row_data = selected_data.iloc[row_idx, :]
                                    line = self.process_row_data(row_data, range_loop_content, start_row, row_idx, selected_data.shape[0])
                                    loop_result.append(line)
                            
                            # Replace range loop content
                            file_content = file_content.replace(
                                f"{range_loop_start}{range_loop_content}{range_loop_end}", 
                                "".join(loop_result)
                            )
                
                # Process RANGE_DATA_LOOP in file content if present
                # 在檔案內容中處理 RANGE_DATA_LOOP（如果存在）
                for range_name in range_names:
                    if "{{RANGE_DATA_LOOP_START}}" in file_content:
                        range_indices = self.convert_range_notation_to_indices(range_name)
                        if range_indices:
                            file_content = self.process_range_data_loop(
                                file_content, df, range_indices, is_column_mode
                            )
                            break  # 只處理第一個找到的範圍
                
                # Process last file comma handling
                is_last_file = (file_idx == len(excel_files) - 1)
                if is_last_file and file_content.rstrip().endswith(","):
                    file_content = file_content.rstrip().rstrip(",") + file_content[len(file_content.rstrip()):]
                
                files_result.append(file_content)
            
            # Combine all files
            final_code = before_files_loop + "".join(files_result) + after_files_loop
        
        return final_code
    
    def process_3d_multi_range_template(self, template, excel_files, dfs, selected_ranges, file_count, is_column_mode=False):
        """處理三維多範圍陣列樣板"""
        if not selected_ranges or len(selected_ranges) < 1:
            messagebox.showerror("錯誤", "需要選擇至少一個數據範圍來處理三維多範圍陣列模板")
            return template
        
        final_code = template
        
        # 處理範圍維度信息
        if "{{RANGES_LOOP_START}}" in final_code and "{{RANGES_LOOP_END}}" in final_code:
            # 處理範圍維度定義部分
            range_dim_parts = final_code.split("unsigned int range_dimensions[RANGE_COUNT][2] = {")
            before_range_dim = range_dim_parts[0]
            
            if len(range_dim_parts) > 1:
                range_dim_and_after = range_dim_parts[1].split("};")
                range_dim_content = "unsigned int range_dimensions[RANGE_COUNT][2] = {" + range_dim_and_after[0] + "};"
                after_range_dim = range_dim_and_after[1] if len(range_dim_and_after) > 1 else ""
                
                range_dim_result = self.process_range_dimensions(range_dim_content, selected_ranges, is_column_mode)
                
                # 將處理後的範圍維度部分重新組合回代碼中
                final_code = before_range_dim + range_dim_result + after_range_dim
        
        # 處理檔案循環
        if "{{FILES_LOOP_START}}" in final_code and "{{FILES_LOOP_END}}" in final_code:
            parts = final_code.split("{{FILES_LOOP_START}}")
            before_files_loop = parts[0]
            
            files_loop_and_after = parts[1].split("{{FILES_LOOP_END}}")
            files_loop_content = files_loop_and_after[0]
            after_files_loop = files_loop_and_after[1] if len(files_loop_and_after) > 1 else ""
            
            files_result = []
            
            # 處理每個檔案
            for file_idx, file_path in enumerate(excel_files):
                df = dfs[file_path]
                file_content = files_loop_content.replace("{{FILE_INDEX}}", str(file_idx))
                file_content = file_content.replace("{{FILE_NAME}}", os.path.basename(file_path))
                
                # 處理檔案內的範圍循環
                if "{{RANGES_LOOP_START}}" in file_content and "{{RANGES_LOOP_END}}" in file_content:
                    ranges_parts = file_content.split("{{RANGES_LOOP_START}}")
                    before_ranges_loop = ranges_parts[0]
                    
                    ranges_loop_and_after = ranges_parts[1].split("{{RANGES_LOOP_END}}")
                    ranges_loop_content = ranges_loop_and_after[0]
                    after_ranges_loop = ranges_loop_and_after[1] if len(ranges_loop_and_after) > 1 else ""
                    
                    ranges_result = []
                    
                    # 處理每個範圍
                    for range_idx, range_info in enumerate(selected_ranges):
                        range_content = ranges_loop_content
                        
                        # 替換範圍相關信息
                        range_content = range_content.replace("{{RANGE_INDEX}}", str(range_idx))
                        range_content = range_content.replace("{{RANGE_STR}}", range_info['range_str'])
                        
                        # 處理範圍內的數據循環 - 加入 RANGE_DATA_LOOP 支援
                        if "{{RANGE_DATA_LOOP_START}}" in range_content and "{{RANGE_DATA_LOOP_END}}" in range_content:
                            # 提取此範圍的數據
                            start_row = range_info['start_row']
                            start_col = range_info['start_col']
                            end_row = range_info['end_row']
                            end_col = range_info['end_col']
                            
                            range_indices = (start_row, start_col, end_row, end_col)
                            
                            # 使用新的 process_range_data_loop 方法
                            range_content = self.process_range_data_loop(
                                range_content, df, range_indices, is_column_mode
                            )
                        
                        # 處理最後一個範圍的逗號
                        is_last_range = (range_idx == len(selected_ranges) - 1)
                        if is_last_range and range_content.rstrip().endswith(","):
                            range_content = range_content.rstrip().rstrip(",") + range_content[len(range_content.rstrip()):]
                        
                        ranges_result.append(range_content)
                    
                    # 組合該檔案的所有範圍
                    file_content = before_ranges_loop + "".join(ranges_result) + after_ranges_loop
                
                # 處理最後一個檔案的逗號
                is_last_file = (file_idx == file_count - 1)
                if is_last_file and file_content.rstrip().endswith(","):
                    file_content = file_content.rstrip().rstrip(",") + file_content[len(file_content.rstrip()):]
                
                files_result.append(file_content)
            
            # 組合所有檔案
            final_code = before_files_loop + "".join(files_result) + after_files_loop
        
        # 處理完所有標記後，檢查模板中是否有剩餘文本需要保留
        try:
            # 找出最後一個標記
            last_tags = ["{{FILES_LOOP_END}}", "{{RANGES_LOOP_END}}", "{{RANGE_DATA_LOOP_END}}"]
            original_template = template
            processed_template = final_code
            
            # 找出模板中最後一個標記的位置
            last_tag_pos = -1
            last_tag = None
            for tag in last_tags:
                pos = original_template.rfind(tag)
                if pos > last_tag_pos:
                    last_tag_pos = pos
                    last_tag = tag
            
            if last_tag_pos != -1:
                # 計算標記結束位置
                last_tag_end = last_tag_pos + len(last_tag)
                # 提取標記之後的文本
                remaining_text = original_template[last_tag_end:]
                
                if remaining_text.strip():
                    self.gui.log(f"找到模板中的剩餘文本: {remaining_text[:20]}...")
                    
                    # 在處理後的代碼中找到對應標記的位置
                    final_tag_pos = processed_template.rfind(last_tag)
                    if final_tag_pos != -1:
                        final_tag_end = final_tag_pos + len(last_tag)
                        # 添加剩餘文本到最終代碼
                        final_code = processed_template[:final_tag_end] + remaining_text
                        self.gui.log("已添加模板中的剩餘文本到生成的代碼")
        except Exception as e:
            self.gui.log(f"處理模板剩餘文本時出錯: {str(e)}")
            import traceback
            self.gui.log(traceback.format_exc())
        
        return final_code

    def process_range_dimensions(self, range_dim_content, selected_ranges, is_column_mode=False):
        """處理範圍維度定義部分"""
        if "{{RANGES_LOOP_START}}" in range_dim_content and "{{RANGES_LOOP_END}}" in range_dim_content:
            parts = range_dim_content.split("{{RANGES_LOOP_START}}")
            before_ranges_loop = parts[0]
            
            ranges_loop_and_after = parts[1].split("{{RANGES_LOOP_END}}")
            ranges_loop_content = ranges_loop_and_after[0]
            after_ranges_loop = ranges_loop_and_after[1] if len(ranges_loop_and_after) > 1 else ""
            
            ranges_result = []
            
            # 為每個範圍生成維度信息
            for range_idx, range_info in enumerate(selected_ranges):
                range_content = ranges_loop_content
                
                # 計算此範圍的大小
                range_row_count = range_info['end_row'] - range_info['start_row'] + 1
                range_col_count = range_info['end_col'] - range_info['start_col'] + 1
                
                # 根據讀取方向調整維度信息
                if is_column_mode:
                    # 直向讀取時交換行列數
                    range_dimension = f"{range_col_count}, {range_row_count}"
                else:
                    # 橫向讀取保持原樣
                    range_dimension = f"{range_row_count}, {range_col_count}"
                
                # 替換範圍相關的標記
                range_content = range_content.replace("{{RANGE_INDEX}}", str(range_idx))
                range_content = range_content.replace("{{RANGE_STR}}", range_info['range_str'])
                range_content = range_content.replace("{{RANGE_ROW_COUNT}}", str(range_row_count))
                range_content = range_content.replace("{{RANGE_COL_COUNT}}", str(range_col_count))
                
                # 處理最後一個範圍的逗號
                is_last_range = (range_idx == len(selected_ranges) - 1)
                if is_last_range and range_content.rstrip().endswith(","):
                    range_content = range_content.rstrip().rstrip(",") + range_content[len(range_content.rstrip()):]
                
                ranges_result.append(range_content)
            
            # 組合所有範圍的維度信息
            return before_ranges_loop + "".join(ranges_result) + after_ranges_loop
        
        return range_dim_content

    # 處理標準模板函數 - 其餘函數保持不變...
    def process_4d_range_first_template(self, template, excel_files, dfs, selected_ranges, file_count, row_count, col_count, is_column_mode=False):
        """處理四維陣列樣板 - 範圍優先 [範圍][檔案][行][列]"""
        if not selected_ranges or len(selected_ranges) < 1:
            messagebox.showerror("錯誤", "需要選擇至少一個數據範圍來處理四維陣列模板")
            return template
        
        final_code = template
        
        # 處理範圍循環
        if "{{RANGES_LOOP_START}}" in final_code and "{{RANGES_LOOP_END}}" in final_code:
            parts = final_code.split("{{RANGES_LOOP_START}}")
            before_ranges_loop = parts[0]
            
            ranges_loop_and_after = parts[1].split("{{RANGES_LOOP_END}}")
            ranges_loop_content = ranges_loop_and_after[0]
            after_ranges_loop = ranges_loop_and_after[1] if len(ranges_loop_and_after) > 1 else ""
            
            ranges_result = []
            
            # 處理每個範圍
            for range_idx, range_info in enumerate(selected_ranges):
                range_content = ranges_loop_content
                
                # 計算此範圍的大小
                range_row_count = range_info['end_row'] - range_info['start_row'] + 1
                range_col_count = range_info['end_col'] - range_info['start_col'] + 1
                
                # 替換範圍相關的標記
                range_content = range_content.replace("{{RANGE_INDEX}}", str(range_idx))
                range_content = range_content.replace("{{RANGE_STR}}", range_info['range_str'])
                range_content = range_content.replace("{{RANGE_ROW_COUNT}}", str(range_row_count))
                range_content = range_content.replace("{{RANGE_COL_COUNT}}", str(range_col_count))
                
                # 處理檔案循環
                if "{{FILES_LOOP_START}}" in range_content and "{{FILES_LOOP_END}}" in range_content:
                    files_parts = range_content.split("{{FILES_LOOP_START}}")
                    before_files_loop = files_parts[0]
                    
                    files_loop_and_after = files_parts[1].split("{{FILES_LOOP_END}}")
                    files_loop_content = files_loop_and_after[0]
                    after_files_loop = files_loop_and_after[1] if len(files_loop_and_after) > 1 else ""
                    
                    files_result = []
                    
                    # 為每個文件處理數據
                    for file_idx, file_path in enumerate(excel_files):
                        file_content = files_loop_content
                        
                        # 替換文件相關標記
                        file_content = file_content.replace("{{FILE_INDEX}}", str(file_idx))
                        file_content = file_content.replace("{{FILE_NAME}}", os.path.basename(file_path))
                        
                        # 處理範圍內的數據循環
                        if "{{RANGE_LOOP_START}}" in file_content and "{{RANGE_LOOP_END}}" in file_content:
                            df = dfs[file_path]
                            
                            # 提取此範圍的數據
                            start_row = range_info['start_row']
                            start_col = range_info['start_col']
                            end_row = range_info['end_row']
                            end_col = range_info['end_col']
                            
                            selected_data = df.iloc[start_row:end_row+1, start_col:end_col+1]
                            
                            # 分割模板以獲取循環內容
                            loop_parts = file_content.split("{{RANGE_LOOP_START}}")
                            before_loop = loop_parts[0]
                            
                            loop_and_after = loop_parts[1].split("{{RANGE_LOOP_END}}")
                            loop_content = loop_and_after[0]
                            after_loop = loop_and_after[1] if len(loop_and_after) > 1 else ""
                            
                            loop_result = []
                            
                            # 檢查循環內容是否指定了讀取方向
                            local_is_column_mode = self.check_direction_mode(loop_content) or is_column_mode
                            
                            if local_is_column_mode:
                                # 直向讀取 - 按列處理
                                for col_idx in range(selected_data.shape[1]):
                                    column_data = selected_data.iloc[:, col_idx]
                                    line = self.process_column_data(column_data, loop_content, start_col, col_idx, selected_data.shape[1])
                                    loop_result.append(line)
                            else:
                                # 橫向讀取 - 按行處理
                                for row_idx in range(selected_data.shape[0]):
                                    row_data = selected_data.iloc[row_idx, :]
                                    line = self.process_row_data(row_data, loop_content, start_row, row_idx, range_row_count)
                                    loop_result.append(line)
                            
                            # 組合該文件的所有行
                            file_content = before_loop + "".join(loop_result) + after_loop
                        
                        # 處理最後一個文件的逗號
                        is_last_file = (file_idx == file_count - 1)
                        if is_last_file and file_content.rstrip().endswith(","):
                            file_content = file_content.rstrip().rstrip(",") + file_content[len(file_content.rstrip()):]
                        
                        files_result.append(file_content)
                    
                    # 組合該範圍的所有文件
                    range_content = before_files_loop + "".join(files_result) + after_files_loop
                
                # 處理最後一個範圍的逗號
                is_last_range = (range_idx == len(selected_ranges) - 1)
                if is_last_range and range_content.rstrip().endswith(","):
                    range_content = range_content.rstrip().rstrip(",") + range_content[len(range_content.rstrip()):]
                
                ranges_result.append(range_content)
            
            # 組合所有範圍
            final_code = before_ranges_loop + "".join(ranges_result) + after_ranges_loop
        
        return final_code

    def process_4d_file_first_template(self, template, excel_files, dfs, selected_ranges, file_count, row_count, col_count, is_column_mode=False):
        """處理四維陣列樣板 - 檔案優先 [檔案][範圍][行][列]"""
        if not selected_ranges or len(selected_ranges) < 1:
            messagebox.showerror("錯誤", "需要選擇至少一個數據範圍來處理四維陣列模板")
            return template
        
        final_code = template
        
        # 處理範圍維度定義部分
        if "{{RANGES_LOOP_START}}" in final_code and "{{RANGES_LOOP_END}}" in final_code:
            try:
                range_dim_parts = final_code.split("unsigned int range_dimensions[RANGE_COUNT][2] = {")
                # if len(range_dim_parts) > 1:
                #     range_dim_and_after = range_dim_parts[1].split("};")
                #     range_dim_content = "unsigned int range_dimensions[RANGE_COUNT][2] = {" + range_dim_and_after[0] + "};"
                #     after_range_dim = range_dim_and_after[1] if len(range_dim_and_after) > 1 else ""
                    
                #     range_dim_result = self.process_range_dimensions(range_dim_content, selected_ranges, is_column_mode)
                    
                #     # 將處理後的範圍維度部分重新組合回代碼中
                #     final_code = range_dim_parts[0] + range_dim_result + after_range_dim
                # else:
                #     self.gui.log("警告: 未找到標準範圍維度定義格式，嘗試替代格式")
                #     # 嘗試其他可能的格式，例如 static 或不同的維度
                #     alt_patterns = [
                #         "static unsigned int range_dimensions[NORMAL_TABLE_COUNT][3] = {",
                #         "unsigned int range_dimensions[NORMAL_TABLE_COUNT][2] = {"
                #     ]
                #     for pattern in alt_patterns:
                #         if pattern in final_code:
                #             self.gui.log(f"使用替代格式: {pattern}")
                #             range_dim_parts = final_code.split(pattern)
                #             if len(range_dim_parts) > 1:
                #                 range_dim_and_after = range_dim_parts[1].split("};")
                #                 range_dim_content = pattern + range_dim_and_after[0] + "};"
                #                 after_range_dim = range_dim_and_after[1] if len(range_dim_and_after) > 1 else ""
                                
                #                 range_dim_result = self.process_range_dimensions(range_dim_content, selected_ranges, is_column_mode)
                                
                #                 # 將處理後的範圍維度部分重新組合回代碼中
                #                 final_code = range_dim_parts[0] + range_dim_result + after_range_dim
                #                 break
            except Exception as e:
                self.gui.log(f"處理範圍維度時出錯: {str(e)}")
                import traceback
                self.gui.log(traceback.format_exc())
        
        # 處理檔案循環
        if "{{FILES_LOOP_START}}" in final_code and "{{FILES_LOOP_END}}" in final_code:
            parts = final_code.split("{{FILES_LOOP_START}}")
            before_files_loop = parts[0]
            
            files_loop_and_after = parts[1].split("{{FILES_LOOP_END}}")
            files_loop_content = files_loop_and_after[0]
            after_files_loop = files_loop_and_after[1] if len(files_loop_and_after) > 1 else ""
            
            files_result = []
            
            # 處理每個檔案
            for file_idx, file_path in enumerate(excel_files):
                df = dfs[file_path]
                file_content = files_loop_content.replace("{{FILE_INDEX}}", str(file_idx))
                file_content = file_content.replace("{{FILE_NAME}}", os.path.basename(file_path))
                
                # 處理檔案內的範圍循環
                if "{{RANGES_LOOP_START}}" in file_content and "{{RANGES_LOOP_END}}" in file_content:
                    ranges_parts = file_content.split("{{RANGES_LOOP_START}}")
                    before_ranges_loop = ranges_parts[0]
                    
                    ranges_loop_and_after = ranges_parts[1].split("{{RANGES_LOOP_END}}")
                    ranges_loop_content = ranges_loop_and_after[0]
                    after_ranges_loop = ranges_loop_and_after[1] if len(ranges_loop_and_after) > 1 else ""
                    
                    ranges_result = []
                    
                    # 處理每個範圍
                    for range_idx, range_info in enumerate(selected_ranges):
                        range_content = ranges_loop_content
                        
                        # 替換範圍相關信息
                        range_content = range_content.replace("{{RANGE_INDEX}}", str(range_idx))
                        range_content = range_content.replace("{{RANGE_STR}}", range_info['range_str'])
                        
                        # 處理範圍內的數據循環
                        if "{{RANGE_LOOP_START}}" in range_content and "{{RANGE_LOOP_END}}" in range_content:
                            # 提取此範圍的數據
                            start_row = range_info['start_row']
                            start_col = range_info['start_col']
                            end_row = range_info['end_row']
                            end_col = range_info['end_col']
                            
                            selected_data = df.iloc[start_row:end_row+1, start_col:end_col+1]
                            
                            # 分割模板以獲取循環內容
                            loop_parts = range_content.split("{{RANGE_LOOP_START}}")
                            before_loop = loop_parts[0]
                            
                            loop_and_after = loop_parts[1].split("{{RANGE_LOOP_END}}")
                            loop_content = loop_and_after[0]
                            after_loop = loop_and_after[1] if len(loop_and_after) > 1 else ""
                            
                            loop_result = []
                            
                            # 檢查循環內容是否指定了讀取方向
                            local_is_column_mode = self.check_direction_mode(loop_content) or is_column_mode
                            
                            if local_is_column_mode:
                                # 直向讀取 - 按列處理
                                for col_idx in range(selected_data.shape[1]):
                                    column_data = selected_data.iloc[:, col_idx]
                                    line = self.process_column_data(column_data, loop_content, start_col, col_idx, selected_data.shape[1])
                                    loop_result.append(line)
                            else:
                                # 橫向讀取 - 按行處理
                                for row_idx in range(selected_data.shape[0]):
                                    row_data = selected_data.iloc[row_idx, :]
                                    line = self.process_row_data(row_data, loop_content, start_row, row_idx, end_row - start_row + 1)
                                    loop_result.append(line)
                            
                            # 組合該範圍的所有行
                            range_content = before_loop + "".join(loop_result) + after_loop
                        
                        # 處理最後一個範圍的逗號
                        is_last_range = (range_idx == len(selected_ranges) - 1)
                        if is_last_range and range_content.rstrip().endswith(","):
                            range_content = range_content.rstrip().rstrip(",") + range_content[len(range_content.rstrip()):]
                        
                        ranges_result.append(range_content)
                    
                    # 組合該檔案的所有範圍
                    file_content = before_ranges_loop + "".join(ranges_result) + after_ranges_loop
                
                # 處理最後一個檔案的逗號
                is_last_file = (file_idx == file_count - 1)
                if is_last_file and file_content.rstrip().endswith(","):
                    file_content = file_content.rstrip().rstrip(",") + file_content[len(file_content.rstrip()):]
                
                files_result.append(file_content)
            
            # 組合所有檔案
            final_code = before_files_loop + "".join(files_result) + after_files_loop
        
        return final_code

    def process_multi_range_template(self, template, excel_files, dfs, selected_ranges, is_column_mode=False):
        """處理包含多个数据范围的模板"""
        if not selected_ranges or len(selected_ranges) < 1:
            messagebox.showerror("錯誤", "需要選擇至少一個數據範圍來處理多範圍模板")
            return template
        
        final_code = template
        
        # 添加范围常量定义
        for idx, range_info in enumerate(selected_ranges):
            range_num = idx + 1
            row_count = range_info['end_row'] - range_info['start_row'] + 1
            col_count = range_info['end_col'] - range_info['start_col'] + 1
            
            # 替换范围大小常量
            final_code = final_code.replace(f"{{{{RANGE_{range_num}_ROW_COUNT}}}}", str(row_count))
            final_code = final_code.replace(f"{{{{RANGE_{range_num}_COL_COUNT}}}}", str(col_count))
        
        # 处理每个范围的循环
        for range_idx, range_info in enumerate(selected_ranges):
            range_num = range_idx + 1
            
            # 检查是否有此范围的循环标记
            loop_start = f"{{{{RANGE:{range_num}_LOOP_START}}}}"
            loop_end = f"{{{{RANGE:{range_num}_LOOP_END}}}}"
            
            if loop_start in final_code and loop_end in final_code:
                # 分割模板以获取循环内容
                parts = final_code.split(loop_start)
                before_loop = parts[0]
                
                loop_and_after = parts[1].split(loop_end)
                loop_content = loop_and_after[0]
                after_loop = loop_and_after[1] if len(loop_and_after) > 1 else ""
                
                # 获取此范围的数据
                first_file = excel_files[0]  # 使用第一个文件
                df = dfs[first_file]
                
                start_row = range_info['start_row']
                start_col = range_info['start_col']
                end_row = range_info['end_row']
                end_col = range_info['end_col']
                
                selected_data = df.iloc[start_row:end_row+1, start_col:end_col+1]
                
                # 檢查循環內容是否指定了讀取方向
                local_is_column_mode = self.check_direction_mode(loop_content) or is_column_mode
                
                # 生成循环代码
                loop_result = []
                
                if local_is_column_mode:
                    # 直向讀取 - 按列處理
                    for col_idx in range(selected_data.shape[1]):
                        column_data = selected_data.iloc[:, col_idx]
                        line = self.process_column_data(column_data, loop_content, start_col, col_idx, selected_data.shape[1])
                        loop_result.append(line)
                else:
                    # 橫向讀取 - 按行處理
                    for row_idx in range(selected_data.shape[0]):
                        row_data = selected_data.iloc[row_idx, :]
                        line = self.process_row_data(row_data, loop_content, start_row, row_idx, end_row - start_row + 1)
                        loop_result.append(line)
                
                # 更新代码
                final_code = before_loop + "".join(loop_result) + after_loop
        
        # 处理常规循环（如果存在）
        if "{{LOOP_START}}" in final_code and "{{LOOP_END}}" in final_code:
            # 使用第一个范围处理常规循环
            if selected_ranges:
                selected_range = selected_ranges[0]
                
                first_file = excel_files[0]
                df = dfs[first_file]
                
                start_row = selected_range['start_row']
                start_col = selected_range['start_col']
                end_row = selected_range['end_row']
                end_col = selected_range['end_col']
                
                selected_data = df.iloc[start_row:end_row+1, start_col:end_col+1]
                
                # 分割模板
                parts = final_code.split("{{LOOP_START}}")
                before_loop = parts[0]
                
                loop_and_after = parts[1].split("{{LOOP_END}}")
                loop_content = loop_and_after[0]
                after_loop = loop_and_after[1] if len(loop_and_after) > 1 else ""
                
                # 檢查循環內容是否指定了讀取方向
                local_is_column_mode = self.check_direction_mode(loop_content) or is_column_mode
                
                # 处理循环
                loop_result = []
                
                if local_is_column_mode:
                    # 直向讀取 - 按列處理
                    for col_idx in range(selected_data.shape[1]):
                        column_data = selected_data.iloc[:, col_idx]
                        line = self.process_column_data(column_data, loop_content, start_col, col_idx, selected_data.shape[1])
                        loop_result.append(line)
                else:
                    # 橫向讀取 - 按行處理
                    for row_idx in range(selected_data.shape[0]):
                        row_data = selected_data.iloc[row_idx, :]
                        line = self.process_row_data(row_data, loop_content, start_row, row_idx, end_row - start_row + 1)
                        loop_result.append(line)
                
                # 更新代码
                final_code = before_loop + "".join(loop_result) + after_loop
            
        return final_code
    
    def process_3d_template(self, template, excel_files, dfs, start_row, start_col, end_row, end_col, file_count, row_count, is_column_mode=False):
        """處理三維陣列樣板"""
        parts = template.split("{{FILES_LOOP_START}}")
        before_files_loop = parts[0]
        
        files_loop_and_after = parts[1].split("{{FILES_LOOP_END}}")
        files_loop_content = files_loop_and_after[0]
        after_files_loop = files_loop_and_after[1] if len(files_loop_and_after) > 1 else ""
        
        files_result = []
        
        # 为每个文件处理数据
        for file_idx, file_path in enumerate(excel_files):
            df = dfs[file_path]
            selected_data = df.iloc[start_row:end_row+1, start_col:end_col+1]
            
            # 替换文件相关标记
            file_content = files_loop_content.replace("{{FILE_INDEX}}", str(file_idx))
            file_content = file_content.replace("{{FILE_NAME}}", os.path.basename(file_path))
            
            # 处理每个文件内的数据循环
            if "{{LOOP_START}}" in file_content and "{{LOOP_END}}" in file_content:
                loop_parts = file_content.split("{{LOOP_START}}")
                before_loop = loop_parts[0]
                
                loop_and_after = loop_parts[1].split("{{LOOP_END}}")
                loop_content = loop_and_after[0]
                after_loop = loop_and_after[1] if len(loop_and_after) > 1 else ""
                
                # 檢查循環內容是否指定了讀取方向
                local_is_column_mode = self.check_direction_mode(loop_content) or is_column_mode
                
                loop_result = []
                
                if local_is_column_mode:
                    # 直向讀取 - 按列處理
                    for col_idx in range(selected_data.shape[1]):
                        column_data = selected_data.iloc[:, col_idx]
                        line = self.process_column_data(column_data, loop_content, start_col, col_idx, selected_data.shape[1])
                        loop_result.append(line)
                else:
                    # 橫向讀取 - 按行處理
                    for row_idx in range(selected_data.shape[0]):
                        row_data = selected_data.iloc[row_idx, :]
                        line = self.process_row_data(row_data, loop_content, start_row, row_idx, row_count)
                        loop_result.append(line)
                
                # 组合该文件的所有行
                file_content = before_loop + "".join(loop_result) + after_loop
            
            # 处理最后一个文件的逗号
            is_last_file = (file_idx == file_count - 1)
            if is_last_file and file_content.rstrip().endswith(","):
                file_content = file_content.rstrip().rstrip(",") + file_content[len(file_content.rstrip()):]
            
            files_result.append(file_content)
        
        # 组合最终代码
        return before_files_loop + "".join(files_result) + after_files_loop
    
    def process_standard_template(self, template, excel_files, dfs, start_row, start_col, end_row, end_col, row_count, is_column_mode=False):
        """處理標準樣板（單文件）"""
        # 使用第一个文件的数据
        first_file = excel_files[0]
        df = dfs[first_file]
        
        # 確保資料框索引是重置過的
        df = df.reset_index(drop=True)
        
        # 詳細輸出調試信息
        self.gui.log(f"資料範圍詳情:")
        self.gui.log(f"開始行: {start_row} (Excel行號: {start_row+1})")
        self.gui.log(f"開始列: {start_col} (Excel列標: {self.get_column_letter(start_col)})")
        self.gui.log(f"結束行: {end_row} (Excel行號: {end_row+1})")
        self.gui.log(f"結束列: {end_col} (Excel列標: {self.get_column_letter(end_col)})")
        self.gui.log(f"資料框形狀: {df.shape}")
        self.gui.log(f"讀取方向: {'直向(Column)' if is_column_mode else '橫向(Row)'}")
        
        # 處理範圍超出實際資料大小的情況
        if start_row < 0:
            self.gui.log(f"警告: 起始行 {start_row} 不能為負數，已調整為 0")
            start_row = 0
        
        if start_col < 0:
            self.gui.log(f"警告: 起始列 {start_col} 不能為負數，已調整為 0")
            start_col = 0
        
        # 調整結束索引以確保在有效範圍內
        end_row = min(end_row, df.shape[0]-1)
        end_col = min(end_col, df.shape[1]-1)
        
        # 檢查範圍是否有效
        if start_row > end_row or start_col > end_col:
            self.gui.log(f"錯誤: 無效的範圍，起始位置 ({start_row}, {start_col}) 大於結束位置 ({end_row}, {end_col})")
            return template
        
        try:
            # 提取所選範圍的資料
            self.gui.log(f"嘗試提取範圍: 從 [{start_row}, {start_col}] 到 [{end_row}, {end_col}]")
            
            # 使用深拷貝避免修改原始資料
            selected_data = df.iloc[start_row:end_row+1, start_col:end_col+1].copy()
            
            self.gui.log(f"提取的資料形狀: {selected_data.shape}")
            if not selected_data.empty:
                self.gui.log(f"提取的資料前3行:")
                for i in range(min(3, selected_data.shape[0])):
                    self.gui.log(f"Row {i}: {selected_data.iloc[i].tolist()}")
            
            # 檢查循環內容是否指定了讀取方向
            local_is_column_mode = is_column_mode
            
            # 處理模板中的循環部分
            if "{{LOOP_START}}" in template and "{{LOOP_END}}" in template:
                parts = template.split("{{LOOP_START}}")
                before_loop = parts[0]
                
                loop_and_after = parts[1].split("{{LOOP_END}}")
                loop_content = loop_and_after[0]
                after_loop = loop_and_after[1] if len(loop_and_after) > 1 else ""
                
                # 再次檢查循環內容中的方向標記
                local_is_column_mode = self.check_direction_mode(loop_content) or local_is_column_mode
                
                loop_result = []
                
                if local_is_column_mode:
                    # 直向讀取模式 - 按列處理
                    self.gui.log("使用直向讀取模式處理資料")
                    for col_idx in range(selected_data.shape[1]):
                        column_data = selected_data.iloc[:, col_idx]
                        line = self.process_column_data(column_data, loop_content, start_col, col_idx, selected_data.shape[1])
                        loop_result.append(line)
                else:
                    # 橫向讀取模式 - 按行處理
                    self.gui.log("使用橫向讀取模式處理資料")
                    for row_idx in range(selected_data.shape[0]):
                        row_data = selected_data.iloc[row_idx, :]
                        line = self.process_row_data(row_data, loop_content, start_row, row_idx, selected_data.shape[0])
                        loop_result.append(line)
                
                return before_loop + "".join(loop_result) + after_loop
            else:
                self.gui.log("警告: 模板中未找到循環標記 {{LOOP_START}} 和 {{LOOP_END}}")
                return template
        
        except Exception as e:
            self.gui.log(f"處理資料時出錯: {str(e)}")
            import traceback
            self.gui.log(traceback.format_exc())
            return template

    def get_column_letter(self, col_idx):
        """將數字列索引轉換為 Excel 列標記 (例如: 0->A, 1->B, 26->AA)"""
        result = ""
        temp = col_idx + 1  # 轉為 1-based 索引
        while temp > 0:
            temp, remainder = divmod(temp - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def process_row_data(self, row, loop_content, start_row, row_idx, row_count):
        """
        處理單行資料的模板替換 (橫向讀取模式)
        
        Args:
            row: 當前處理的資料行
            loop_content: 循環內容的模板
            start_row: 開始行索引
            row_idx: 當前行索引
            row_count: 總行數
                
        Returns:
            str: 替換後的程式碼行
        """
        # 處理當前行
        line = loop_content
        
        # 替換ROW_INDEX
        line = line.replace("{{ROW_INDEX}}", str(row_idx))
        
        # 替換COL_INDEX (在橫向讀取中設為-1，表示不適用)
        line = line.replace("{{COL_INDEX}}", "-1")
        
        # 處理 {{ALL_COLUMNS}} 標記
        if "{{ALL_COLUMNS}}" in line:
            all_values = []
            
            # 檢查資料是否為空
            if len(row) == 0:
                self.gui.log(f"警告: 行 {row_idx} 資料為空")
                line = line.replace("{{ALL_COLUMNS}}", "0")
            else:
                # 列出所有資料
                for col_idx in range(len(row)):
                    try:
                        # 獲取當前值
                        if hasattr(row, 'iloc'):
                            value = row.iloc[col_idx]
                        else:
                            value = row[col_idx]
                        
                        # 格式化單元格值
                        str_value = format_cell_value(value)
                        
                        # 特別檢查並處理數字字串
                        # 如果是純數字但被包在引號中（如"1000"），移除引號
                        if str_value.startswith('"') and str_value.endswith('"'):
                            inner_value = str_value[1:-1]
                            # 檢查內部值是否為整數或浮點數
                            if inner_value.isdigit() or \
                            (inner_value.startswith('-') and inner_value[1:].isdigit()) or \
                            (inner_value.replace('.', '', 1).isdigit() and inner_value.count('.') == 1):
                                str_value = inner_value
                        
                        all_values.append(str_value)
                        
                    except Exception as e:
                        self.gui.log(f"處理行 {row_idx} 列 {col_idx} 時出錯: {str(e)}")
                        all_values.append("0")
                
                # 連接所有值，以逗號分隔
                all_columns_str = ", ".join(all_values)
                line = line.replace("{{ALL_COLUMNS}}", all_columns_str)
        
        # 處理 {{ALL_ROWS}} 標記 (在橫向讀取中不適用，保留為空字串)
        if "{{ALL_ROWS}}" in line:
            line = line.replace("{{ALL_ROWS}}", "/* ROW MODE: ALL_ROWS not applicable */")
        
        # 處理 {{ROW:n}} 標記
        row_references = re.findall(r'{{ROW:(\d+)}}', line)
        for ref in row_references:
            try:
                ref_idx = int(ref)
                if hasattr(row, 'iloc') and 0 <= ref_idx < len(row):
                    ref_value = row.iloc[ref_idx]
                    ref_str = format_cell_value(ref_value)
                    # 檢查並移除不必要的引號
                    if ref_str.startswith('"') and ref_str.endswith('"'):
                        inner_value = ref_str[1:-1]
                        if inner_value.isdigit() or \
                        (inner_value.startswith('-') and inner_value[1:].isdigit()) or \
                        (inner_value.replace('.', '', 1).isdigit() and inner_value.count('.') == 1):
                            ref_str = inner_value
                    line = line.replace(f"{{{{ROW:{ref}}}}}", ref_str)
                    self.gui.log(f"ROW:{ref} = {ref_str}")
                elif isinstance(row, (list, pd.Series)) and 0 <= ref_idx < len(row):
                    ref_value = row[ref_idx]
                    ref_str = format_cell_value(ref_value)
                    # 檢查並移除不必要的引號
                    if ref_str.startswith('"') and ref_str.endswith('"'):
                        inner_value = ref_str[1:-1]
                        if inner_value.isdigit() or \
                        (inner_value.startswith('-') and inner_value[1:].isdigit()) or \
                        (inner_value.replace('.', '', 1).isdigit() and inner_value.count('.') == 1):
                            ref_str = inner_value
                    line = line.replace(f"{{{{ROW:{ref}}}}}", ref_str)
                    self.gui.log(f"ROW:{ref} = {ref_str}")
                else:
                    self.gui.log(f"ROW:{ref} 超出範圍 (0-{len(row)-1})")
                    line = line.replace(f"{{{{ROW:{ref}}}}}", "0")
            except Exception as e:
                self.gui.log(f"處理 ROW:{ref} 時出錯: {str(e)}")
                line = line.replace(f"{{{{ROW:{ref}}}}}", "0")
        
        # 處理 {{COL:n}} 標記
        col_references = re.findall(r'{{COL:(\d+)}}', line)
        for ref in col_references:
            try:
                ref_idx = int(ref)
                if hasattr(row, 'index') and 0 <= ref_idx < len(row.index):
                    col_name = row.index[ref_idx]
                    col_value = row[col_name]
                    col_str = format_cell_value(col_value)
                    # 檢查並移除不必要的引號
                    if col_str.startswith('"') and col_str.endswith('"'):
                        inner_value = col_str[1:-1]
                        if inner_value.isdigit() or \
                        (inner_value.startswith('-') and inner_value[1:].isdigit()) or \
                        (inner_value.replace('.', '', 1).isdigit() and inner_value.count('.') == 1):
                            col_str = inner_value
                    line = line.replace(f"{{{{COL:{ref}}}}}", col_str)
                    self.gui.log(f"COL:{ref} = {col_str}")
                elif isinstance(row, (list, pd.Series)) and 0 <= ref_idx < len(row):
                    col_value = row[ref_idx]
                    col_str = format_cell_value(col_value)
                    # 檢查並移除不必要的引號
                    if col_str.startswith('"') and col_str.endswith('"'):
                        inner_value = col_str[1:-1]
                        if inner_value.isdigit() or \
                        (inner_value.startswith('-') and inner_value[1:].isdigit()) or \
                        (inner_value.replace('.', '', 1).isdigit() and inner_value.count('.') == 1):
                            col_str = inner_value
                    line = line.replace(f"{{{{COL:{ref}}}}}", col_str)
                    self.gui.log(f"COL:{ref} = {col_str}")
                else:
                    self.gui.log(f"COL:{ref} 超出索引範圍 (0-{len(row.index)-1 if hasattr(row, 'index') else len(row)-1})")
                    line = line.replace(f"{{{{COL:{ref}}}}}", "0")
            except Exception as e:
                self.gui.log(f"處理 COL:{ref} 時出錯: {str(e)}")
                line = line.replace(f"{{{{COL:{ref}}}}}", "0")
        
        # 替換VALUE為第一個列的值
        if len(row) > 0:
            try:
                if hasattr(row, 'iloc'):
                    value = row.iloc[0]
                else:
                    value = row[0]
                str_value = format_cell_value(value)
                # 檢查並移除不必要的引號
                if str_value.startswith('"') and str_value.endswith('"'):
                    inner_value = str_value[1:-1]
                    if inner_value.isdigit() or \
                    (inner_value.startswith('-') and inner_value[1:].isdigit()) or \
                    (inner_value.replace('.', '', 1).isdigit() and inner_value.count('.') == 1):
                        str_value = inner_value
                line = line.replace("{{VALUE}}", str_value)
            except Exception as e:
                self.gui.log(f"處理 VALUE 時出錯: {str(e)}")
                line = line.replace("{{VALUE}}", "0")
        else:
            self.gui.log("沒有資料可用於 VALUE")
            line = line.replace("{{VALUE}}", "0")
        
        # 修改這一段逗號處理邏輯
        is_last_row = (row_idx == row_count - 1)
        
        # 檢查是否是陣列初始化項目
        contains_array_item = ("{" in line and "}" in line)
        is_array_element = "{" in line or "}" in line
        
        # 移除現有的行尾逗號（如果有）
        if line.rstrip().endswith(","):
            line = line.rstrip()[:-1]
        
        # 根據陣列結構和位置決定是否添加逗號
        if is_last_row and "}" in line and not "{" in line:
            # 如果是陣列最後一行且包含閉合括號，不添加逗號
            pass
        else:
            # 所有其他行都需要添加逗號
            line = line.rstrip() + ","
        
        return line

    def process_column_data(self, column, loop_content, start_col, col_idx, col_count):
        """
        處理單列資料的模板替換 (直向讀取模式)
        
        Args:
            column: 當前處理的資料列
            loop_content: 循環內容的模板
            start_col: 開始列索引
            col_idx: 當前列索引
            col_count: 總列數
            
        Returns:
            str: 替換後的程式碼行
        """
        # 處理當前列
        line = loop_content
        
        # 替換COL_INDEX
        line = line.replace("{{COL_INDEX}}", str(col_idx))
        
        # 替換ROW_INDEX (在直向讀取中設為-1，表示不適用)
        line = line.replace("{{ROW_INDEX}}", "-1")
        
        # 處理 {{ALL_ROWS}} 標記 - 直向讀取模式的主要特點
        if "{{ALL_ROWS}}" in line:
            all_values = []
            
            # 檢查資料是否為空
            if len(column) == 0:
                self.gui.log(f"警告: 列 {col_idx} 資料為空")
                line = line.replace("{{ALL_ROWS}}", "0")
            else:
                # 列出所有資料
                for row_idx in range(len(column)):
                    try:
                        # 獲取當前值
                        if hasattr(column, 'iloc'):
                            value = column.iloc[row_idx]
                        else:
                            value = column[row_idx]
                        
                        # 格式化單元格值
                        str_value = format_cell_value(value)
                        
                        # 特別檢查並處理數字字串
                        if str_value.startswith('"') and str_value.endswith('"'):
                            inner_value = str_value[1:-1]
                            if inner_value.isdigit() or \
                            (inner_value.startswith('-') and inner_value[1:].isdigit()) or \
                            (inner_value.replace('.', '', 1).isdigit() and inner_value.count('.') == 1):
                                str_value = inner_value
                        
                        all_values.append(str_value)
                        
                    except Exception as e:
                        self.gui.log(f"處理列 {col_idx} 行 {row_idx} 時出錯: {str(e)}")
                        all_values.append("0")
                
                # 連接所有值，以逗號分隔
                all_rows_str = ", ".join(all_values)
                line = line.replace("{{ALL_ROWS}}", all_rows_str)
        
        # 處理 {{ALL_COLUMNS}} 標記 (在直向讀取中不適用，保留為空字串)
        if "{{ALL_COLUMNS}}" in line:
            line = line.replace("{{ALL_COLUMNS}}", "/* COLUMN MODE: ALL_COLUMNS not applicable */")
        
        # 處理 {{ROW:n}} 標記 - 直向讀取時，這表示第n個行的值
        row_references = re.findall(r'{{ROW:(\d+)}}', line)
        for ref in row_references:
            try:
                ref_idx = int(ref)
                if hasattr(column, 'iloc') and 0 <= ref_idx < len(column):
                    ref_value = column.iloc[ref_idx]
                    ref_str = format_cell_value(ref_value)
                    # 檢查並移除不必要的引號
                    if ref_str.startswith('"') and ref_str.endswith('"'):
                        inner_value = ref_str[1:-1]
                        if inner_value.isdigit() or \
                        (inner_value.startswith('-') and inner_value[1:].isdigit()) or \
                        (inner_value.replace('.', '', 1).isdigit() and inner_value.count('.') == 1):
                            ref_str = inner_value
                    line = line.replace(f"{{{{ROW:{ref}}}}}", ref_str)
                elif isinstance(column, (list, pd.Series)) and 0 <= ref_idx < len(column):
                    ref_value = column[ref_idx]
                    ref_str = format_cell_value(ref_value)
                    # 檢查並移除不必要的引號
                    if ref_str.startswith('"') and ref_str.endswith('"'):
                        inner_value = ref_str[1:-1]
                        if inner_value.isdigit() or \
                        (inner_value.startswith('-') and inner_value[1:].isdigit()) or \
                        (inner_value.replace('.', '', 1).isdigit() and inner_value.count('.') == 1):
                            ref_str = inner_value
                    line = line.replace(f"{{{{ROW:{ref}}}}}", ref_str)
                else:
                    self.gui.log(f"ROW:{ref} 超出範圍 (0-{len(column)-1})")
                    line = line.replace(f"{{{{ROW:{ref}}}}}", "0")
            except Exception as e:
                self.gui.log(f"處理 ROW:{ref} 時出錯: {str(e)}")
                line = line.replace(f"{{{{ROW:{ref}}}}}", "0")
        
        # 處理 {{COL:n}} 標記 - 直向讀取時不常用，但依然處理
        col_references = re.findall(r'{{COL:(\d+)}}', line)
        for ref in col_references:
            # 在直向讀取中，COL: 參考主要指的是當前列的偏移
            ref_idx = int(ref)
            line = line.replace(f"{{{{COL:{ref}}}}}", str(col_idx + ref_idx))
        
        # 替換VALUE為第一個行的值
        if len(column) > 0:
            try:
                if hasattr(column, 'iloc'):
                    value = column.iloc[0]
                else:
                    value = column[0]
                str_value = format_cell_value(value)
                # 檢查並移除不必要的引號
                if str_value.startswith('"') and str_value.endswith('"'):
                    inner_value = str_value[1:-1]
                    if inner_value.isdigit() or \
                    (inner_value.startswith('-') and inner_value[1:].isdigit()) or \
                    (inner_value.replace('.', '', 1).isdigit() and inner_value.count('.') == 1):
                        str_value = inner_value
                line = line.replace("{{VALUE}}", str_value)
            except Exception as e:
                self.gui.log(f"處理 VALUE 時出錯: {str(e)}")
                line = line.replace("{{VALUE}}", "0")
        else:
            self.gui.log("沒有資料可用於 VALUE")
            line = line.replace("{{VALUE}}", "0")
        
        # 處理最後一列的逗號
        is_last_col = (col_idx == col_count - 1)
        if is_last_col and line.rstrip().endswith(","):
            line = line.rstrip().rstrip(",") + line[len(line.rstrip()):]
        
        return line