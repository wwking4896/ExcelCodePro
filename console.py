"""
Console version of ExcelCode Pro.
Allows generating code from Excel files using a config file without GUI.
"""

import argparse
import json
import os
import sys
import logging
import pandas as pd
from code_generator import CodeGenerator
from utils import load_config, excel_notation_to_index

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger("ExcelCode-Console")

class ConsoleModeHandler:
    """Handle code generation in console mode"""
    
    def __init__(self):
        """Initialize the console mode handler"""
        self.excel_files = []
        self.dfs = {}
        self.selected_sheet = None
        self.selected_range = None
        self.selected_ranges = []
        self.named_ranges = {}
        self.code_template = None
        self.template_direction = "row"  # Default direction
        self.code_generator = CodeGenerator(self)
    
    def log(self, message):
        """Log messages (compatibility with GUI version)"""
        logger.info(message)
    
    def load_config_file(self, config_path):
        """Load config from JSON file"""
        logger.info(f"Loading config from: {config_path}")
        
        try:
            # Load the config
            with open(config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            
            # Process named ranges first
            if "named_ranges" in config_data:
                self.named_ranges = config_data["named_ranges"]
                logger.info(f"Loaded {len(self.named_ranges)} named ranges")
            
            # Process Excel files
            if "excel_files" in config_data and config_data["excel_files"]:
                # Filter out files that don't exist
                existing_files = []
                for file_path in config_data["excel_files"]:
                    if os.path.exists(file_path):
                        existing_files.append(file_path)
                    else:
                        logger.warning(f"Excel file not found: {file_path}")
                
                if not existing_files:
                    logger.error("No valid Excel files found in config")
                    return False
                
                self.excel_files = existing_files
                logger.info(f"Found {len(self.excel_files)} valid Excel files")
            else:
                logger.error("No Excel files specified in config")
                return False
            
            # Process selected sheet
            if "selected_sheet" in config_data:
                self.selected_sheet = config_data["selected_sheet"]
                logger.info(f"Selected sheet: {self.selected_sheet}")
            else:
                logger.error("No sheet specified in config")
                return False
            
            # Process selected ranges
            if "selected_ranges" in config_data and config_data["selected_ranges"]:
                self.selected_ranges = config_data["selected_ranges"]
                self.selected_range = self.selected_ranges[0]  # For compatibility
                range_strs = [r['range_str'] for r in self.selected_ranges]
                logger.info(f"Loaded {len(self.selected_ranges)} ranges: {', '.join(range_strs)}")
            elif self.named_ranges:
                # Convert named ranges to selected ranges
                logger.info("Creating ranges from named ranges...")
                for name, range_str in self.named_ranges.items():
                    try:
                        start, end = range_str.split(":")
                        start_row, start_col = excel_notation_to_index(start)
                        end_row, end_col = excel_notation_to_index(end)
                        
                        range_info = {
                            'start_row': start_row,
                            'start_col': start_col,
                            'end_row': end_row,
                            'end_col': end_col,
                            'range_str': range_str
                        }
                        self.selected_ranges.append(range_info)
                        logger.info(f"Created range from '{name}': {range_str}")
                    except Exception as e:
                        logger.error(f"Error processing range {name}: {str(e)}")
                
                if self.selected_ranges:
                    self.selected_range = self.selected_ranges[0]
            else:
                logger.error("No ranges specified in config and no named ranges available")
                return False
            
            # Process template
            if "template_type" in config_data:
                template_type = config_data["template_type"]
                
                if template_type == "custom" and "code_template" in config_data:
                    self.code_template = config_data["code_template"]
                    logger.info("Loaded custom template")
                elif template_type == "preset" and "preset_template" in config_data:
                    preset_name = config_data["preset_template"]
                    self.code_template = self.code_generator.get_default_template(preset_name)
                    logger.info(f"Using preset template: {preset_name}")
                else:
                    logger.error("Invalid template configuration")
                    return False
            elif "code_template" in config_data:  # Backwards compatibility
                self.code_template = config_data["code_template"]
                logger.info("Loaded template from config")
            else:
                logger.error("No template specified in config")
                return False
            
            # Process direction if available
            if "template_direction" in config_data:
                self.template_direction = config_data["template_direction"]
                logger.info(f"Template direction set to: {self.template_direction}")
            
            return True
            
        except Exception as e:
            logger.error(f"Error loading config: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def load_excel_data(self):
        """Load Excel files and selected sheet data"""
        logger.info("Loading Excel data...")
        
        try:
            for file_path in self.excel_files:
                logger.info(f"Processing file: {os.path.basename(file_path)}")
                
                # Read all data as object first, then convert to numeric where possible
                df = pd.read_excel(file_path, sheet_name=self.selected_sheet, dtype=object, header=None)
                
                # Try to convert numeric columns
                df = df.apply(pd.to_numeric, errors='ignore')
                
                # Reset index to ensure it starts from 0
                df = df.reset_index(drop=True)
                
                logger.info(f"DataFrame shape: {df.shape}")
                self.dfs[file_path] = df
            
            return True
            
        except Exception as e:
            logger.error(f"Error loading Excel data: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def generate_code(self):
        """Generate code based on loaded data and template"""
        logger.info("Generating code...")
        
        try:
            # Check if we have everything we need
            has_valid_range = (self.selected_range is not None) or (len(self.selected_ranges) > 0)
            
            if not has_valid_range or not self.code_template or not self.excel_files:
                logger.error("Missing required data for code generation")
                return None
            
            # Generate the code
            final_code = self.code_generator.generate_code(
                self.excel_files, 
                self.dfs, 
                self.selected_ranges, 
                self.code_template, 
                self.selected_range
            )
            
            logger.info("Code generation completed successfully")
            return final_code
            
        except Exception as e:
            logger.error(f"Error generating code: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return None

def parse_arguments():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(description='ExcelCode Pro - Console Version')
    parser.add_argument('--config', '-c', required=True, help='Path to config JSON file')
    parser.add_argument('--output', '-o', help='Output file path (if not specified, prints to stdout)')
    parser.add_argument('--verbose', '-v', action='store_true', help='Enable verbose logging')
    
    return parser.parse_args()

def main():
    """Main function for console operation"""
    args = parse_arguments()
    
    # Set log level based on verbose flag
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    
    # Check if config file exists
    if not os.path.exists(args.config):
        logger.error(f"Config file not found: {args.config}")
        return 1
    
    # Initialize handler
    handler = ConsoleModeHandler()
    
    # Load config
    if not handler.load_config_file(args.config):
        logger.error("Failed to load configuration")
        return 1
    
    # Load Excel data
    if not handler.load_excel_data():
        logger.error("Failed to load Excel data")
        return 1
    
    # Generate code
    generated_code = handler.generate_code()
    if generated_code is None:
        logger.error("Failed to generate code")
        return 1
    
    # Output the generated code
    if args.output:
        try:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(generated_code)
            logger.info(f"Code saved to {args.output}")
        except Exception as e:
            logger.error(f"Error saving code to file: {str(e)}")
            return 1
    else:
        # Print to stdout
        print(generated_code)
    
    logger.info("Done")
    return 0

if __name__ == "__main__":
    sys.exit(main())