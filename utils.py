import re
import pandas as pd
import json
import os

def excel_notation_to_index(notation, gui=None):
    """
    將 Excel 標記轉換為 (行,列) 索引
    例如: A1 -> (0,0), B3 -> (2,1), AA1 -> (0,26)
    """
    match = re.match(r'([A-Z]+)([0-9]+)', notation.upper())
    if not match:
        raise ValueError(f"無效的Excel標記: {notation}")
    
    col_str, row_str = match.groups()
    
    # 修正列標識轉換邏輯 (A->0, B->1, ..., Z->25, AA->26, ...)
    col_idx = 0
    for c in col_str:
        col_idx = col_idx * 26 + (ord(c) - ord('A') + 1)
    col_idx -= 1  # 調整為0基礎索引
    
    # 將行號轉換為索引 (從0開始)
    row_idx = int(row_str) - 1
    
    # 使用 gui.log 如果提供，否則使用 print
    # if gui:
    #     gui.log(f"Excel標記 {notation} 轉換為索引: ({row_idx}, {col_idx})")
    # else:
    #     print(f"Excel標記 {notation} 轉換為索引: ({row_idx}, {col_idx})")
    
    return row_idx, col_idx

def format_cell_value(value, add_quotes=False):
    """
    Format cell value to appropriate string format
    - 將 NaN 轉換為 "0"
    - 將整數浮點數轉換為整數字串
    - 文字值可選擇是否添加引號
    
    Args:
        value: 要格式化的值
        add_quotes: 是否為文字值添加引號，預設為 False
    """
    if pd.isna(value):
        return "0"
    elif isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        else:
            return str(value)
    else:
        # 根據參數決定是否添加引號
        if add_quotes:
            return f'"{value}"'
        else:
            return str(value)

def replace_template_placeholders(template, replacements):
    """
    替換模板中的佔位符
    
    Args:
        template (str): 包含佔位符的模板字符串
        replacements (dict): 佔位符和替換值的字典
        
    Returns:
        str: 替換後的字符串
    """
    result = template
    for placeholder, value in replacements.items():
        result = result.replace(f"{{{{{placeholder}}}}}", str(value))
    return result

def save_config(config_data, file_path, gui=None):
    """
    將設定資料儲存為 JSON 檔案
    
    Args:
        config_data (dict): 要儲存的設定資料
        file_path (str): 儲存的檔案路徑
        gui (object): GUI 實例，用於記錄
    """
    try:
        # 添加命名範圍資訊到設定
        if hasattr(gui, 'named_ranges'):
            config_data["named_ranges"] = gui.named_ranges
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, ensure_ascii=False, indent=4)
        
        if gui:
            gui.log(f"設定已儲存至: {file_path}")
        return True
    except Exception as e:
        if gui:
            gui.log(f"儲存設定檔時出錯: {str(e)}")
        else:
            print(f"儲存設定檔時出錯: {str(e)}")
        return False

def load_config(file_path, gui=None):
    """
    從 JSON 檔案載入設定資料
    
    Args:
        file_path (str): 設定檔案路徑
        gui (object): GUI 實例，用於記錄
        
    Returns:
        dict: 設定資料字典，如果載入失敗則返回 None
    """
    try:
        if not os.path.exists(file_path):
            if gui:
                gui.log(f"設定檔不存在: {file_path}")
            return None
            
        with open(file_path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
        
        # 載入命名範圍資訊
        if "named_ranges" in config_data and gui:
            gui.named_ranges = config_data["named_ranges"]
            gui.log(f"已載入 {len(gui.named_ranges)} 個命名範圍")
        
        if gui:
            gui.log(f"已從 {file_path} 載入設定")
        return config_data
    except Exception as e:
        if gui:
            gui.log(f"載入設定檔時出錯: {str(e)}")
        else:
            print(f"載入設定檔時出錯: {str(e)}")
        return None

def get_templates_directory():
    """
    取得或創建用於存放樣板的目錄
    
    Returns:
        str: 樣板目錄的路徑
    """
    # 使用與程式相同目錄下的 templates 子目錄
    base_dir = os.path.dirname(os.path.abspath(__file__))
    template_dir = os.path.join(base_dir, "templates")
    
    # 如果目錄不存在，則創建它
    if not os.path.exists(template_dir):
        os.makedirs(template_dir)
    
    return template_dir

def get_resource_path(relative_path):
    """獲取資源文件的絕對路徑，兼容開發環境和打包後的環境"""
    try:
        # PyInstaller 創建臨時文件夾 _MEIPASS
        import sys
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, relative_path)