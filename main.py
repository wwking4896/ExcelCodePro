import logging
import sys
import os  # 加入這一行來引入 os 模組

def setup_logging():
    """設定日誌記錄"""
    log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.log")
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_path, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )

def excepthook(exc_type, exc_value, exc_traceback):
    """處理未捕獲的異常"""
    logging.error("未捕獲的異常", exc_info=(exc_type, exc_value, exc_traceback))
    
if __name__ == "__main__":
    setup_logging()
    sys.excepthook = excepthook
    
    # 這裡應該有引入 tk 和 ExcelToCodeApp
    import tkinter as tk
    from gui import ExcelToCodeApp
    
    root = tk.Tk()
    app = ExcelToCodeApp(root)
    root.mainloop()