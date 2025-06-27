"""
Simple test to create an Excel file with our enhanced formatting
"""
import pandas as pd
import sys
import os

# Add current directory to path
sys.path.append('.')

# Import our dataclasses
import Dataclasses as dc
import Fredon_projects_fields as pf

def create_simple_test_portfolio():
    """Create a simple test portfolio with sample data"""
    
    # Create a sample project
    project = dc.Projects(
        Project_Name="Test Project Alpha",
        Project_ID="TP001",
        Location="Test Location",
        Post_Code="12345",
        Sector="Technology",
        Portfolio_Bus_Unit_Dept_ID="TECH001",
        Asset_Type="Commercial",
        Contract_Type="Fixed Price",
        Contract_Financial="5M",
        Client="Test Client Corp",
        Stage_of_Work="In Progress",
        Comments="This is a test project for formatting validation",
        Template_Version="1.0"
    )
    
    # Add some monthly records
    monthly_records = [
        dc.MonthlyRecord(
            Date="01/01/2024",
            approved_budget=1000000.0,
            forecast_end_date=20241201,
            forecast_final_cost=950000.0,
            contingency_remaining=50000.0,
            actual_cost_to_date=600000.0,
            forecast_final_revenue=1200000.0,
            actual_revenue_to_date=700000.0,
            notes="85% Cost Complete"
        ),
        dc.MonthlyRecord(
            Date="01/02/2024",
            approved_budget=1000000.0,
            forecast_end_date=20241201,
            forecast_final_cost=980000.0,
            contingency_remaining=20000.0,
            actual_cost_to_date=750000.0,
            forecast_final_revenue=1200000.0,
            actual_revenue_to_date=850000.0,
            notes="92% Cost Complete"
        ),
        dc.MonthlyRecord(
            Date="01/03/2024",
            approved_budget=1000000.0,
            forecast_end_date=20241201,
            forecast_final_cost=990000.0,
            contingency_remaining=10000.0,
            actual_cost_to_date=950000.0,
            forecast_final_revenue=1200000.0,
            actual_revenue_to_date=1100000.0,
            notes="98% Cost Complete"
        )
    ]
    
    project.Monthly_data = monthly_records
    
    # Create portfolio and add project
    portfolio = dc.Portfolio()
    portfolio.add_project(project)
    
    return portfolio

def test_excel_export():
    """Test our enhanced Excel export"""
    try:
        print("Creating test portfolio...")
        portfolio = create_simple_test_portfolio()
        
        print(f"Portfolio has {len(portfolio.projects)} project(s)")
        
        # Import the Excel export function
        from Fredon_Methods_test import create_excel_file_with_portfolio_data
        
        print("Exporting to Excel...")
        files_created = create_excel_file_with_portfolio_data(portfolio, "enhanced_format_test")
        
        if files_created:
            print(f"Successfully created files: {files_created}")
            
            # Analyze the created file
            import openpyxl as px
            filename = files_created[0]
            print(f"\nAnalyzing created file: {filename}")
            
            wb = px.load_workbook(filename)
            sheet_name = wb.sheetnames[0]
            sheet = wb[sheet_name]
            
            print(f"Sheet: {sheet_name}")
            
            # Check key cells for formatting
            test_cells = [
                ("A1", "Template Version"),
                ("B1", "Version Value"),
                ("H1", "Main Header"),
                ("F2", "Project Summary"),
                ("A4", "Project Name Label"),
                ("C4", "Project Name Value"),
                ("F20", "Monthly Data Header"),
                ("A21", "Reporting Date Header"),
                ("A22", "First Data Row")
            ]
            
            for cell_ref, description in test_cells:
                cell = sheet[cell_ref]
                print(f"\n{description} ({cell_ref}):")
                print(f"  Value: '{cell.value}'")
                if cell.font:
                    print(f"  Font: {cell.font.name}, Size: {cell.font.size}, Bold: {cell.font.bold}")
                if cell.fill and hasattr(cell.fill, 'start_color') and cell.fill.start_color and cell.fill.start_color.rgb != '00000000':
                    print(f"  Fill: {cell.fill.start_color.rgb}")
                if cell.alignment and (cell.alignment.horizontal or cell.alignment.vertical):
                    print(f"  Alignment: H={cell.alignment.horizontal}, V={cell.alignment.vertical}")
                if cell.border and any([
                    cell.border.left.style if cell.border.left else None, 
                    cell.border.right.style if cell.border.right else None,
                    cell.border.top.style if cell.border.top else None, 
                    cell.border.bottom.style if cell.border.bottom else None]):
                    left_style = cell.border.left.style if cell.border.left else "None"
                    right_style = cell.border.right.style if cell.border.right else "None"
                    top_style = cell.border.top.style if cell.border.top else "None"
                    bottom_style = cell.border.bottom.style if cell.border.bottom else "None"
                    print(f"  Border: L={left_style}, R={right_style}, "
                          f"T={top_style}, B={bottom_style}")
            
            # Check merged cells
            print(f"\nMerged cells in output:")
            for merged_range in sheet.merged_cells.ranges:
                print(f"  {merged_range}")
            
            print(f"\nTest completed successfully!")
            
        else:
            print("No files were created")
            
    except Exception as e:
        print(f"Error in test: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_excel_export()
