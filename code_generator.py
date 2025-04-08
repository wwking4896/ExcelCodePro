import re
import pandas as pd
import numpy as np
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
import threading

class CodeGenerator:
    def __init__(self, gui_instance):
        self.gui = gui_instance  # 保存對主GUI實例的引用
    
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
        """設置自訂程式碼樣板"""
        # 建立樣板設定窗口
        template_dialog = tk.Toplevel(self.gui.root)
        template_dialog.title("設定程式碼樣板")
        template_dialog.geometry("700x650")  # 增加視窗高度，從550改為650
        template_dialog.grab_set()  # 模态窗口
        template_dialog.minsize(700, 650)  # 設定最小尺寸，確保按鈕不會被隱藏
        
        # 样板说明
        instruction_frame = ttk.Frame(template_dialog)
        instruction_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(instruction_frame, text="程式碼樣板說明：", font=("Arial", 10, "bold")).pack(anchor="w")
        instruction_text = """
    在樣板中，您可以使用以下標記：
    - {{LOOP_START}} - 資料迴圈開始 (使用第一個選擇的範圍)
    - {{LOOP_END}} - 資料迴圈結束 (使用第一個選擇的範圍)
    - {{FILES_LOOP_START}} - 檔案迴圈開始 (用於三維陣列)
    - {{FILES_LOOP_END}} - 檔案迴圈結束 (用於三維陣列)
    - {{RANGES_LOOP_START}} - 範圍迴圈開始 (用於四維陣列)
    - {{RANGES_LOOP_END}} - 範圍迴圈結束 (用於四維陣列)
    - {{RANGE_DATA_LOOP_START}} - 範圍資料迴圈開始 (用於三維多範圍陣列)
    - {{RANGE_DATA_LOOP_END}} - 範圍資料迴圈結束 (用於三維多範圍陣列)
    - {{FILE_NAME}} - 當前處理的檔案名稱
    - {{FILE_INDEX}} - 當前檔案索引
    - {{FILE_COUNT}} - 總檔案數量
    - {{RANGE_INDEX}} - 當前範圍索引
    - {{RANGE_STR}} - 當前範圍字串表示
    - {{RANGE_COUNT}} - 總範圍數量
    - {{RANGE_ROW_COUNT}} - 當前範圍的行數
    - {{RANGE_COL_COUNT}} - 當前範圍的列數
    - {{MAX_ROW_COUNT}} - 所有範圍中的最大行數
    - {{MAX_COL_COUNT}} - 所有範圍中的最大列數
    - {{ROW_INDEX}} - 當前資料列索引
    - {{COL_INDEX}} - 當前資料行索引
    - {{VALUE}} - 當前儲存格值
    - {{ALL_COLUMNS}} - 當前行的所有欄位值
    - {{ROW:數字}} - 獲取資料列的數字欄位值（例如：{{ROW:0}}取得該列第1個欄位的值）
    - {{COL:數字}} - 獲取資料行的數字欄位值（例如：{{COL:0}}取得該行第1個欄位的值）

    多範圍處理標記：
    - {{RANGE:n_LOOP_START}} - 第n個範圍的資料迴圈開始 (n從1開始)
    - {{RANGE:n_LOOP_END}} - 第n個範圍的資料迴圈結束
    - {{RANGE_n_ROW_COUNT}} - 第n個範圍的資料列數
    - {{RANGE_n_COL_COUNT}} - 第n個範圍的資料行數
        """
            
        instruction_box = ScrolledText(instruction_frame, height=12, font=("Courier New", 9))
        instruction_box.pack(fill="x", pady=5)
        instruction_box.insert("1.0", instruction_text)
        instruction_box.config(state="disabled")
        
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
            # 预设样板 - 添加多範圍支持
            multi_range_template = """// 多範圍數據處理範例
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
    };

    // 範例：將兩個區域的數據組合處理的函數
    void process_combined_data() {
        for (int i = 0; i < RANGE_1_ROW_COUNT && i < RANGE_2_ROW_COUNT; i++) {
            // 範例處理邏輯
            first_area[i][0] += second_area[i][0];
        }
    }
    """
            template_text.insert("1.0", multi_range_template)
        
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
            self.gui.code_template = template_text.get("1.0", tk.END).strip()
            if self.gui.code_template:
                self.gui.template_preview.config(text="已設定自訂樣板")
                self.gui.generate_button.config(state="normal")
                template_dialog.destroy()
            else:
                messagebox.showerror("錯誤", "樣板不能為空", parent=template_dialog)
        
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
        elif template_name == "二維陣列":
            return """// 二維陣列初始化
    #define ROW_COUNT 20
    #define COL_COUNT 4

    unsigned int table[ROW_COUNT][COL_COUNT] = {
    {{LOOP_START}}
        { {{ALL_COLUMNS}} },
    {{LOOP_END}}
    };"""
        elif template_name == "三維陣列":
            # 添加三维数组模板
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
        elif template_name == "四維陣列 (範圍優先)":
            # 原版四維陣列樣板 - 以範圍為優先
            return """// 四維陣列初始化 - 範圍優先 [範圍][檔案][行][列]
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
    }

    // 範例：列印特定範圍的資料摘要
    void print_range_summary(unsigned int range_idx) {
        if (range_idx >= RANGE_COUNT) {
            return;
        }
        
        printf("範圍 %u 資料摘要 (行數: %u, 列數: %u):\\n", 
            range_idx, range_dimensions[range_idx][0], range_dimensions[range_idx][1]);
            
        // 以第一個檔案為例
        for (int row = 0; row < range_dimensions[range_idx][0]; row++) {
            for (int col = 0; col < range_dimensions[range_idx][1]; col++) {
                printf("%u ", data4d[range_idx][0][row][col]);
            }
            printf("\\n");
        }
    }"""
        elif template_name == "四維陣列 (檔案優先)":
            # 新版四維陣列樣板 - 以檔案為優先
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
    };

    // 範例：取得特定檔案、特定範圍的資料
    unsigned int get_value(unsigned int file_idx, unsigned int range_idx, unsigned int row, unsigned int col) {
        // 邊界檢查
        if (file_idx >= FILE_COUNT || range_idx >= RANGE_COUNT || 
            row >= range_dimensions[range_idx][0] || col >= range_dimensions[range_idx][1]) {
            return 0;  // 超出範圍，返回預設值
        }
        
        return data4d[file_idx][range_idx][row][col];
    }

    // 範例：列印特定檔案、特定範圍的資料摘要
    void print_range_summary(unsigned int file_idx, unsigned int range_idx) {
        if (file_idx >= FILE_COUNT || range_idx >= RANGE_COUNT) {
            return;
        }
        
        printf("檔案 %u, 範圍 %u 資料摘要 (行數: %u, 列數: %u):\\n", 
            file_idx, range_idx, range_dimensions[range_idx][0], range_dimensions[range_idx][1]);
            
        for (int row = 0; row < range_dimensions[range_idx][0]; row++) {
            for (int col = 0; col < range_dimensions[range_idx][1]; col++) {
                printf("%u ", data4d[file_idx][range_idx][row][col]);
            }
            printf("\\n");
        }
    }"""
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

    // 三維多範圍陣列: [檔案][範圍][行][列]
    unsigned int data3d_multi[FILE_COUNT][RANGE_COUNT][MAX_ROW_COUNT][MAX_COL_COUNT] = {
    {{FILES_LOOP_START}}
        // 來自檔案: {{FILE_NAME}}
        {
    {{RANGES_LOOP_START}}
            // 範圍 {{RANGE_INDEX}}: {{RANGE_STR}}
            {
    {{RANGE_DATA_LOOP_START}}
                { {{ALL_COLUMNS}} },
    {{RANGE_DATA_LOOP_END}}
            },
    {{RANGES_LOOP_END}}
        },
    {{FILES_LOOP_END}}
    };

    // 範例：取得特定檔案、特定範圍的資料
    unsigned int get_value(unsigned int file_idx, unsigned int range_idx, unsigned int row, unsigned int col) {
        // 邊界檢查
        if (file_idx >= FILE_COUNT || range_idx >= RANGE_COUNT || 
            row >= range_dimensions[range_idx][0] || col >= range_dimensions[range_idx][1]) {
            return 0;  // 超出範圍，返回預設值
        }
        
        return data3d_multi[file_idx][range_idx][row][col];
    }

    // 範例：列印特定檔案、特定範圍的資料
    void print_range_data(unsigned int file_idx, unsigned int range_idx) {
        if (file_idx >= FILE_COUNT || range_idx >= RANGE_COUNT) {
            return;
        }
        
        printf("檔案 %u, 範圍 %u 資料摘要 (行數: %u, 列數: %u):\\n", 
            file_idx, range_idx, range_dimensions[range_idx][0], range_dimensions[range_idx][1]);
            
        for (int row = 0; row < range_dimensions[range_idx][0]; row++) {
            for (int col = 0; col < range_dimensions[range_idx][1]; col++) {
                printf("%u ", data3d_multi[file_idx][range_idx][row][col]);
            }
            printf("\\n");
        }
    }"""
        elif template_name == "權重表設定":
            return """// 遊戲權重表初始化
    void initVariableWeights() {
    {{LOOP_START}}
        normal_table_weight[{{COL:0}}][{{COL:1}}][{{COL:2}}][{{COL:3}}] = {{VALUE}};
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
    };

    // 處理兩個區域數據的函數
    void process_dual_data() {
        // 範例處理邏輯
        for (int i = 0; i < RANGE_1_ROW_COUNT && i < RANGE_2_ROW_COUNT; i++) {
            // 示例操作
            printf("Row %d: First area value = %d, Second area value = %d\\n", 
                i, first_area[i][0], second_area[i][0]);
        }
    }"""
        else:
            return ""
    
    def generate_code(self, excel_files, dfs, selected_ranges, code_template, selected_range):
        """生成程式碼"""
        # 檢查樣板類型
        is_3d_template = "{{FILES_LOOP_START}}" in code_template and "{{FILES_LOOP_END}}" in code_template
        is_4d_template = "{{RANGES_LOOP_START}}" in code_template and "{{RANGES_LOOP_END}}" in code_template
        is_multi_range = "{{RANGE:" in code_template
        is_3d_multi_range = is_3d_template and is_4d_template and "{{RANGE_DATA_LOOP_START}}" in code_template
        is_4d_range_first = "unsigned int data4d[RANGE_COUNT][FILE_COUNT]" in code_template
        is_4d_file_first = "unsigned int data4d[FILE_COUNT][RANGE_COUNT]" in code_template
        
        # 使用第一個選擇的範圍作為主範圍
        if selected_ranges:
            selected_range = selected_ranges[0]
        
        # 提取範圍
        start_row = selected_range['start_row']
        start_col = selected_range['start_col']
        end_row = selected_range['end_row']
        end_col = selected_range['end_col']
        
        # 計算實際的行數和列數
        row_count = end_row - start_row + 1
        col_count = end_col - start_col + 1
        file_count = len(excel_files)
        range_count = len(selected_ranges)
        
        # 計算所有範圍中的最大行數和列數
        max_row_count = 0
        max_col_count = 0
        for range_info in selected_ranges:
            range_row_count = range_info['end_row'] - range_info['start_row'] + 1
            range_col_count = range_info['end_col'] - range_info['start_col'] + 1
            max_row_count = max(max_row_count, range_row_count)
            max_col_count = max(max_col_count, range_col_count)
        
        # 準備最終代碼 - 直接替換所有變數，不使用正則表達式
        template = code_template
        template = template.replace("{{ROW_COUNT}}", str(row_count))
        template = template.replace("{{COL_COUNT}}", str(col_count))
        template = template.replace("{{FILE_COUNT}}", str(file_count))
        template = template.replace("{{RANGE_COUNT}}", str(range_count))
        template = template.replace("{{MAX_ROW_COUNT}}", str(max_row_count))
        template = template.replace("{{MAX_COL_COUNT}}", str(max_col_count))
        
        # 處理不同類型的模板
        if is_3d_multi_range:
            return self.process_3d_multi_range_template(template, excel_files, dfs, selected_ranges, file_count)
        elif is_4d_file_first:
            # 處理以檔案為優先的四維陣列
            return self.process_4d_file_first_template(template, excel_files, dfs, selected_ranges, file_count, row_count, col_count)
        elif is_4d_range_first:
            # 處理以範圍為優先的四維陣列
            return self.process_4d_range_first_template(template, excel_files, dfs, selected_ranges, file_count, row_count, col_count)
        elif is_4d_template:
            # 如果無法確定確切類型，使用通用的四維處理
            return self.process_4d_template(template, excel_files, dfs, selected_ranges, file_count, row_count, col_count)
        elif is_multi_range:
            return self.process_multi_range_template(template, excel_files, dfs, selected_ranges)
        elif is_3d_template:
            return self.process_3d_template(template, excel_files, dfs, start_row, start_col, end_row, end_col, file_count, row_count)
        else:
            return self.process_standard_template(template, excel_files, dfs, start_row, start_col, end_row, end_col, row_count)
    
    def process_3d_multi_range_template(self, template, excel_files, dfs, selected_ranges, file_count):
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
            
            range_dim_and_after = range_dim_parts[1].split("};")
            range_dim_content = "unsigned int range_dimensions[RANGE_COUNT][2] = {" + range_dim_and_after[0] + "};"
            after_range_dim = range_dim_and_after[1] if len(range_dim_and_after) > 1 else ""
            
            range_dim_result = self.process_range_dimensions(range_dim_content, selected_ranges)
            
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
                        
                        # 處理範圍內的數據循環
                        if "{{RANGE_DATA_LOOP_START}}" in range_content and "{{RANGE_DATA_LOOP_END}}" in range_content:
                            # 提取此範圍的數據
                            start_row = range_info['start_row']
                            start_col = range_info['start_col']
                            end_row = range_info['end_row']
                            end_col = range_info['end_col']
                            
                            selected_data = df.iloc[start_row:end_row+1, start_col:end_col+1]
                            
                            # 分割模板以獲取循環內容
                            loop_parts = range_content.split("{{RANGE_DATA_LOOP_START}}")
                            before_loop = loop_parts[0]
                            
                            loop_and_after = loop_parts[1].split("{{RANGE_DATA_LOOP_END}}")
                            loop_content = loop_and_after[0]
                            after_loop = loop_and_after[1] if len(loop_and_after) > 1 else ""
                            
                            loop_result = []
                            
                            # 處理每一行數據
                            for i, row in selected_data.iterrows():
                                row_idx = i - start_row
                                line = self.process_row_data(row, loop_content, start_row, row_idx, end_row - start_row + 1)
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

    def process_range_dimensions(self, range_dim_content, selected_ranges):
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
    
    def process_4d_range_first_template(self, template, excel_files, dfs, selected_ranges, file_count, row_count, col_count):
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
                            
                            # 處理每一行數據
                            for i, row in selected_data.iterrows():
                                row_idx = i - start_row
                                line = self.process_row_data(row, loop_content, start_row, row_idx, range_row_count)
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

    def process_4d_file_first_template(self, template, excel_files, dfs, selected_ranges, file_count, row_count, col_count):
        """處理四維陣列樣板 - 檔案優先 [檔案][範圍][行][列]"""
        if not selected_ranges or len(selected_ranges) < 1:
            messagebox.showerror("錯誤", "需要選擇至少一個數據範圍來處理四維陣列模板")
            return template
        
        final_code = template
        
        # 處理範圍維度定義部分
        if "{{RANGES_LOOP_START}}" in final_code and "{{RANGES_LOOP_END}}" in final_code:
            range_dim_parts = final_code.split("unsigned int range_dimensions[RANGE_COUNT][2] = {")
            before_range_dim = range_dim_parts[0]
            
            range_dim_and_after = range_dim_parts[1].split("};")
            range_dim_content = "unsigned int range_dimensions[RANGE_COUNT][2] = {" + range_dim_and_after[0] + "};"
            after_range_dim = range_dim_and_after[1] if len(range_dim_and_after) > 1 else ""
            
            range_dim_result = self.process_range_dimensions(range_dim_content, selected_ranges)
            
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
                            
                            # 處理每一行數據
                            for i, row in selected_data.iterrows():
                                row_idx = i - start_row
                                line = self.process_row_data(row, loop_content, start_row, row_idx, end_row - start_row + 1)
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

    def process_4d_template(self, template, excel_files, dfs, selected_ranges, file_count, row_count, col_count):
        """處理通用四維陣列樣板 - 自動判斷結構類型"""
        # 檢查四維陣列結構
        if "unsigned int data4d[RANGE_COUNT][FILE_COUNT]" in template:
            # 如果是範圍優先結構 [範圍][檔案][行][列]
            return self.process_4d_range_first_template(template, excel_files, dfs, selected_ranges, file_count, row_count, col_count)
        elif "unsigned int data4d[FILE_COUNT][RANGE_COUNT]" in template:
            # 如果是檔案優先結構 [檔案][範圍][行][列]
            return self.process_4d_file_first_template(template, excel_files, dfs, selected_ranges, file_count, row_count, col_count)
        else:
            # 無法確定結構，使用檔案優先作為預設
            self.gui.log("警告: 無法確定四維陣列結構類型，使用檔案優先結構作為預設")
            return self.process_4d_file_first_template(template, excel_files, dfs, selected_ranges, file_count, row_count, col_count)
    
    def process_multi_range_template(self, template, excel_files, dfs, selected_ranges):
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
                
                # 生成循环代码
                loop_result = []
                for i, row in selected_data.iterrows():
                    row_idx = i - start_row
                    line = self.process_row_data(row, loop_content, start_row, row_idx, end_row - start_row + 1)
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
                
                # 处理循环
                loop_result = []
                for i, row in selected_data.iterrows():
                    row_idx = i - start_row
                    line = self.process_row_data(row, loop_content, start_row, row_idx, end_row - start_row + 1)
                    loop_result.append(line)
                
                # 更新代码
                final_code = before_loop + "".join(loop_result) + after_loop
        
        return final_code
    
    def process_3d_template(self, template, excel_files, dfs, start_row, start_col, end_row, end_col, file_count, row_count):
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
                
                loop_result = []
                
                # 处理每一行数据
                for i, row in selected_data.iterrows():
                    row_idx = i - start_row
                    line = self.process_row_data(row, loop_content, start_row, row_idx, row_count)
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
    
    def process_standard_template(self, template, excel_files, dfs, start_row, start_col, end_row, end_col, row_count):
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
        
        # 試試讀取起始行的第一個儲存格值
        try:
            self.gui.log(f"檢查起始儲存格 ({start_row}, {start_col}) 的值:")
            if start_row < df.shape[0] and start_col < df.shape[1]:
                start_value = df.iloc[start_row, start_col]
                self.gui.log(f"起始儲存格值: {start_value}")
            else:
                self.gui.log(f"起始儲存格 ({start_row}, {start_col}) 超出資料框範圍 {df.shape}")
        except Exception as e:
            self.gui.log(f"讀取起始儲存格時出錯: {str(e)}")
        
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
            
            # 處理模板中的循環部分
            if "{{LOOP_START}}" in template and "{{LOOP_END}}" in template:
                parts = template.split("{{LOOP_START}}")
                before_loop = parts[0]
                
                loop_and_after = parts[1].split("{{LOOP_END}}")
                loop_content = loop_and_after[0]
                after_loop = loop_and_after[1] if len(loop_and_after) > 1 else ""
                
                loop_result = []
                
                # 逐行處理所選資料
                for row_idx in range(selected_data.shape[0]):
                    row = selected_data.iloc[row_idx]
                    line = self.process_row_data(row, loop_content, start_row, row_idx, selected_data.shape[0])
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
        """處理單行資料的模板替換"""
        # 處理當前行
        line = loop_content
        
        # 替換ROW_INDEX
        line = line.replace("{{ROW_INDEX}}", str(row_idx))
        
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
                        str_value = self.format_cell_value(value)
                        
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
        
        # 處理 {{ROW:n}} 標記
        row_references = re.findall(r'{{ROW:(\d+)}}', line)
        for ref in row_references:
            try:
                ref_idx = int(ref)
                if hasattr(row, 'iloc') and 0 <= ref_idx < len(row):
                    ref_value = row.iloc[ref_idx]
                    ref_str = self.format_cell_value(ref_value)
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
                    ref_str = self.format_cell_value(ref_value)
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
                    col_str = self.format_cell_value(col_value)
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
                    col_str = self.format_cell_value(col_value)
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
                str_value = self.format_cell_value(value)
                # 檢查並移除不必要的引號
                if str_value.startswith('"') and str_value.endswith('"'):
                    inner_value = str_value[1:-1]
                    if inner_value.isdigit() or \
                    (inner_value.startswith('-') and inner_value[1:].isdigit()) or \
                    (inner_value.replace('.', '', 1).isdigit() and inner_value.count('.') == 1):
                        str_value = inner_value
                line = line.replace("{{VALUE}}", str_value)
                self.gui.log(f"VALUE = {str_value}")
            except Exception as e:
                self.gui.log(f"處理 VALUE 時出錯: {str(e)}")
                line = line.replace("{{VALUE}}", "0")
        else:
            self.gui.log("沒有資料可用於 VALUE")
            line = line.replace("{{VALUE}}", "0")
        
        # 處理最後一行的逗號
        is_last_row = (row_idx == row_count - 1)
        if is_last_row and line.rstrip().endswith(","):
            line = line.rstrip().rstrip(",") + line[len(line.rstrip()):]
        
        return line

    def format_cell_value(self, value):
        """
        格式化單元格值為適當的字串格式
        - 確保數值不添加引號
        - 只有文字類型才添加引號
        """
        # 處理空值
        if pd.isna(value):
            return "0"
        
        # 檢查是否為數值類型
        if isinstance(value, (int, float, np.int64, np.float64)):
            # 對於浮點數，如果是整數值則轉為整數顯示
            if isinstance(value, (float, np.float64)) and value.is_integer():
                return str(int(value))
            return str(value)
        
        # 處理可能為數字的字串
        if isinstance(value, str):
            # 清理字串
            cleaned = value.strip()
            
            # 嘗試識別數字字串（整數或浮點數）
            try:
                # 如果能轉換為數值，就返回數值形式
                num_val = float(cleaned)
                if num_val.is_integer():
                    return str(int(num_val))
                return str(num_val)
            except (ValueError, TypeError):
                # 不是數值，保持為字串並添加引號
                # 確保特殊字元得到轉義
                escaped_value = cleaned.replace('"', '\\"')
                return f'"{escaped_value}"'
        
        # 其他任何類型，轉為字串並添加引號
        return f'"{str(value)}"'

