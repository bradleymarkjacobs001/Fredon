#!/usr/bin/env python3
"""
Check the number format of financial columns in the generated Excel files.
"""

import openpyxl

def check_number_formats(file_path):
    print(f"=== CHECKING NUMBER FORMATS: {file_path} ===")
    
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb[wb.sheetnames[0]]  # Check first sheet
        
        print(f"Sheet: {sheet.title}")
        
        # Check first data row (row 22) for financial columns
        financial_cols = ['B', 'D', 'E', 'F', 'G', 'H']
        
        for col in financial_cols:
            cell = sheet[f'{col}22']
            print(f"Column {col} (cell {col}22):")
            print(f"  Value: {cell.value}")
            print(f"  Number Format: '{cell.number_format}'")
            print(f"  Data Type: {type(cell.value)}")
            print()
            
    except Exception as e:
        print(f"Error: {e}")
    
    print("=== END ANALYSIS ===\n")

if __name__ == "__main__":
    check_number_formats("portfolio_output_Calibrate.xlsx")
    check_number_formats("portfolio_output_Operational.xlsx")
