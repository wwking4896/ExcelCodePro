# build.py - 打包腳本
import PyInstaller.__main__
import os
import shutil

# 清理先前的打包文件
if os.path.exists('dist'):
    shutil.rmtree('dist')
if os.path.exists('build'):
    shutil.rmtree('build')

# 定義打包參數
pyinstaller_args = [
    'main.py',                    # 主程式檔案
    '--name=ExcelCode Pro',     # 可執行檔案名稱
    '--windowed',                 # 使用GUI模式，不顯示控制台視窗
    '--onefile',                  # 生成單一可執行檔案
    '--icon=icon.ico',            # 應用程式圖標 (可自行添加)
    '--add-data=templates;templates',  # 包含樣板目錄
    '--clean',                    # 清理臨時檔案
    '--noconfirm',                # 不詢問確認
]

# 執行PyInstaller
PyInstaller.__main__.run(pyinstaller_args)

# 建立templates目錄(如果不存在)
if not os.path.exists('dist/Excel轉程式碼工具/templates'):
    os.makedirs('dist/Excel轉程式碼工具/templates')

print("打包完成！")