#!/usr/bin/env python3

import pandas as pd
import Fredon_projects_fields as pf

def debug_percentage_calculation():
    """Debug script to test the percentage calculation logic"""
    
    # Let's try to load the actual data file to test
    try:
        # Try to read the Excel file directly with different approaches
        print("Trying to read Excel file with different header options...")
        
        # First, try without header specification
        print("\n=== Reading with header=0 ===")
        df0 = pd.read_excel("Fredon calibration data.xlsx", header=0)
        print(f"Shape: {df0.shape}")
        print("Columns:", df0.columns.tolist()[:10])  # First 10 columns
        
        # Then try with header=1 (like in the app)
        print("\n=== Reading with header=1 ===")
        df1 = pd.read_excel("Fredon calibration data.xlsx", header=1)
        print(f"Shape: {df1.shape}")
        print("Columns:", df1.columns.tolist()[:10])  # First 10 columns
        
        # Try with header=None to see raw data
        print("\n=== Reading with header=None (raw data) ===")
        df_raw = pd.read_excel("Fredon calibration data.xlsx", header=None)
        print(f"Shape: {df_raw.shape}")
        print("First few rows:")
        print(df_raw.head())
        
        # Look for any cell containing percentage data
        print("\n=== Searching for percentage data in all cells ===")
        for row_idx in range(min(5, len(df_raw))):
            for col_idx in range(min(10, len(df_raw.columns))):
                cell_value = df_raw.iloc[row_idx, col_idx]
                if pd.notna(cell_value):
                    cell_str = str(cell_value).lower()
                    if '%' in cell_str or 'complete' in cell_str or 'percent' in cell_str:
                        print(f"  Found percentage-related content at ({row_idx}, {col_idx}): '{cell_value}'")
        
        # Check if we have the right Excel file
        print("\n=== Available Excel files ===")
        import os
        excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
        print("Excel files in directory:", excel_files)
        
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        print("Make sure 'Fredon calibration data.xlsx' exists in the current directory")

if __name__ == "__main__":
    debug_percentage_calculation()
