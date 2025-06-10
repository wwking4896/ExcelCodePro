# build.py - Enhanced build script with detailed feedback
import PyInstaller.__main__
import os
import shutil
import sys
from pathlib import Path
import time

class BuildLogger:
    """Build process logger for better feedback"""
    def __init__(self):
        self.build_start_time = time.time()
        self.errors = []
        self.warnings = []
        self.success_operations = []
    
    def log_info(self, message):
        """Log information message"""
        print(f"[INFO] {message}")
    
    def log_success(self, message):
        """Log success message"""
        print(f"[SUCCESS] {message}")
        self.success_operations.append(message)
    
    def log_warning(self, message):
        """Log warning message"""
        print(f"[WARNING] {message}")
        self.warnings.append(message)
    
    def log_error(self, message):
        """Log error message"""
        print(f"[ERROR] {message}")
        self.errors.append(message)
    
    def get_build_summary(self):
        """Generate build summary"""
        duration = time.time() - self.build_start_time
        return {
            'duration': duration,
            'errors': self.errors,
            'warnings': self.warnings,
            'success_operations': self.success_operations,
            'is_success': len(self.errors) == 0
        }

def clean_previous_build(logger):
    """Clean previous build artifacts"""
    logger.log_info("清理先前的建置檔案...")
    
    try:
        if os.path.exists('dist'):
            shutil.rmtree('dist')
            logger.log_success("已清理 dist 目錄")
        
        if os.path.exists('build'):
            shutil.rmtree('build')
            logger.log_success("已清理 build 目錄")
    except Exception as e:
        logger.log_error(f"清理建置目錄時發生錯誤: {str(e)}")
        return False
    
    return True

def build_gui_version(logger):
    """Build GUI version of the application"""
    logger.log_info("開始建置 GUI 版本...")
    
    # Define build parameters for GUI version
    pyinstaller_args = [
        'main.py',                    # Main program file
        '--name=ExcelCode Pro',       # Executable name
        '--windowed',                 # GUI mode, hide console window
        '--onefile',                  # Generate single executable
        '--icon=icon.ico',            # Application icon (if exists)
        '--add-data=templates;templates',  # Include templates directory
        '--clean',                    # Clean temporary files
        '--noconfirm',                # Don't ask for confirmation
    ]
    
    try:
        # Execute PyInstaller
        PyInstaller.__main__.run(pyinstaller_args)
        
        # Check if GUI executable was created
        gui_exe_path = os.path.join('dist', 'ExcelCode Pro.exe')
        if os.path.exists(gui_exe_path):
            logger.log_success(f"GUI 版本建置成功: {gui_exe_path}")
            return True
        else:
            logger.log_error("GUI 版本建置失敗: 找不到執行檔")
            return False
            
    except Exception as e:
        logger.log_error(f"GUI 版本建置時發生錯誤: {str(e)}")
        return False

def build_console_version(logger):
    """Build console version of the application"""
    logger.log_info("開始建置 Console 版本...")
    
    try:
        # Execute PyInstaller for console version
        PyInstaller.__main__.run([
            '--clean',
            '--onefile', 
            '--name=ExcelCode-Console',
            'console.py'
        ])
        
        # Check if console executable was created
        console_exe_path = os.path.join('dist', 'ExcelCode-Console.exe')
        if os.path.exists(console_exe_path):
            logger.log_success(f"Console 版本建置成功: {console_exe_path}")
            return True
        else:
            logger.log_error("Console 版本建置失敗: 找不到執行檔")
            return False
            
    except Exception as e:
        logger.log_error(f"Console 版本建置時發生錯誤: {str(e)}")
        return False

def copy_resources(logger):
    """Copy necessary resources to distribution folder"""
    logger.log_info("複製必要資源檔案...")
    
    copy_operations = []
    
    try:
        # Create main distribution directory
        dist_main_dir = os.path.join('dist', 'ExcelCode')
        if not os.path.exists(dist_main_dir):
            os.makedirs(dist_main_dir)
            logger.log_success(f"已建立目錄: {dist_main_dir}")
        
        # Copy README.md
        if os.path.exists('README.md'):
            shutil.copy('README.md', 'dist')
            copy_operations.append("README.md -> dist/")
            logger.log_success("已複製 README.md")
        else:
            logger.log_warning("找不到 README.md 檔案")
        
        # Copy and ensure templates folder
        if os.path.exists('templates'):
            dest_templates = os.path.join('dist', 'templates')
            if os.path.exists(dest_templates):
                shutil.rmtree(dest_templates)
            shutil.copytree('templates', dest_templates)
            copy_operations.append("templates/ -> dist/templates/")
            logger.log_success("已複製 templates 資料夾")
        else:
            logger.log_warning("找不到 templates 資料夾，將建立空的 templates 目錄")
            os.makedirs(os.path.join('dist', 'templates'), exist_ok=True)
        
        # Copy and ensure configs folder
        if os.path.exists('configs'):
            dest_configs = os.path.join('dist', 'configs')
            if os.path.exists(dest_configs):
                shutil.rmtree(dest_configs)
            shutil.copytree('configs', dest_configs)
            copy_operations.append("configs/ -> dist/configs/")
            logger.log_success("已複製 configs 資料夾")
        else:
            logger.log_warning("找不到 configs 資料夾，將建立空的 configs 目錄")
            os.makedirs(os.path.join('dist', 'configs'), exist_ok=True)
        
        # Copy console executable to main distribution folder
        console_exe_src = os.path.join('dist', 'ExcelCode-Console.exe')
        console_exe_dest = os.path.join(dist_main_dir, 'ExcelCode-Console.exe')
        
        if os.path.exists(console_exe_src):
            shutil.copy(console_exe_src, console_exe_dest)
            copy_operations.append("ExcelCode-Console.exe -> dist/ExcelCode/")
            logger.log_success("已複製 Console 版本執行檔")
        else:
            logger.log_warning("找不到 Console 版本執行檔")
        
        # Copy console.bat if exists
        if os.path.exists('console.bat'):
            shutil.copy('console.bat', dist_main_dir)
            copy_operations.append("console.bat -> dist/ExcelCode/")
            logger.log_success("已複製 console.bat")
        
        return copy_operations
        
    except Exception as e:
        logger.log_error(f"複製資源檔案時發生錯誤: {str(e)}")
        return []

def create_example_files(logger):
    """Create example configuration files"""
    logger.log_info("建立範例檔案...")
    
    try:
        examples_dir = os.path.join('dist', 'ExcelCode', 'examples')
        os.makedirs(examples_dir, exist_ok=True)
        
        # Create example config.json
        example_config = {
            "excel_files": ["path/to/your/excel/file.xlsx"],
            "selected_sheet": "Sheet1",
            "selected_ranges": [
                {
                    "start_row": 0,
                    "start_col": 0,
                    "end_row": 10,
                    "end_col": 3,
                    "range_str": "A1:D11"
                }
            ],
            "template_type": "preset",
            "preset_template": "二維陣列"
        }
        
        import json
        config_path = os.path.join(examples_dir, 'config.json')
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(example_config, f, ensure_ascii=False, indent=2)
        
        logger.log_success(f"已建立範例設定檔: examples/config.json")
        
        # Create usage instructions
        usage_content = """使用說明：

1. 命令列使用方式：
   ExcelCode-Console.exe --config path/to/config.json --output path/to/output.c

2. 批次檔使用方式：
   編輯 console.bat 並設定所需參數
   然後執行 console.bat

3. 範例設定檔可在 'examples' 資料夾中找到

參數說明：
--config, -c    指定設定檔路徑
--output, -o    指定輸出檔案路徑  
--verbose, -v   啟用詳細記錄
"""
        
        usage_path = os.path.join('dist', 'ExcelCode', 'console_usage.txt')
        with open(usage_path, 'w', encoding='utf-8') as f:
            f.write(usage_content)
        
        logger.log_success("已建立使用說明檔案: console_usage.txt")
        return True
        
    except Exception as e:
        logger.log_error(f"建立範例檔案時發生錯誤: {str(e)}")
        return False

def generate_requirements(logger):
    """Generate requirements.txt"""
    logger.log_info("產生 requirements.txt...")
    
    try:
        import subprocess
        result = subprocess.run([sys.executable, '-m', 'pip', 'freeze'], 
                              capture_output=True, text=True)
        
        if result.returncode == 0:
            with open('requirements.txt', 'w', encoding='utf-8') as f:
                f.write(result.stdout)
            logger.log_success("已產生 requirements.txt")
            return True
        else:
            logger.log_error("產生 requirements.txt 失敗")
            return False
            
    except Exception as e:
        logger.log_error(f"產生 requirements.txt 時發生錯誤: {str(e)}")
        return False

def print_build_summary(logger, copy_operations):
    """Print detailed build summary"""
    summary = logger.get_build_summary()
    
    print("\n" + "="*40)
    if summary['is_success']:
        print("🎉 BUILD SUCCESS! 🎉")
    else:
        print("❌ BUILD FAILED! ❌")
    
    print("="*40)
    print(f"建置時間: {summary['duration']:.2f} 秒")
    print(f"成功操作: {len(summary['success_operations'])}")
    print(f"警告數量: {len(summary['warnings'])}")
    print(f"錯誤數量: {len(summary['errors'])}")
    
    if copy_operations:
        print("\n📁 檔案複製操作:")
        for operation in copy_operations:
            print(f"   ✓ {operation}")
    
    if summary['success_operations']:
        print("\n✅ 成功操作:")
        for operation in summary['success_operations']:
            print(f"   ✓ {operation}")
    
    if summary['warnings']:
        print("\n⚠️ 警告訊息:")
        for warning in summary['warnings']:
            print(f"   ⚠️ {warning}")
    
    if summary['errors']:
        print("\n❌ 錯誤訊息:")
        for error in summary['errors']:
            print(f"   ❌ {error}")
    
    print("\n📦 建置輸出結構:")
    print("   dist/")
    print("   ├── ExcelCode Pro.exe          (GUI 版本)")
    print("   ├── ExcelCode-Console.exe      (Console 版本)")
    print("   ├── README.md")
    print("   ├── templates/")
    print("   ├── configs/")
    print("   └── ExcelCode/")
    print("       ├── ExcelCode-Console.exe")
    print("       ├── console.bat")
    print("       ├── console_usage.txt")
    print("       └── examples/")
    print("           └── config.json")
    
    print("="*40)
    
    return summary['is_success']

def main():
    """Main build function"""
    logger = BuildLogger()
    
    print("🚀 開始 ExcelCode Pro 建置程序...")
    print("="*40)
    
    # Step 1: Clean previous build
    if not clean_previous_build(logger):
        print_build_summary(logger, [])
        return 1
    
    # Step 2: Build GUI version  
    if not build_gui_version(logger):
        print_build_summary(logger, [])
        return 1
    
    # Step 3: Build console version
    if not build_console_version(logger):
        print_build_summary(logger, [])
        return 1
    
    # Step 4: Copy resources
    copy_operations = copy_resources(logger)
    
    # Step 5: Create example files
    create_example_files(logger)
    
    # Step 6: Generate requirements
    generate_requirements(logger)
    
    # Step 7: Print build summary
    success = print_build_summary(logger, copy_operations)
    
    return 0 if success else 1

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)