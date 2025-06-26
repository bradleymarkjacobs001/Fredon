#import openpyxl as px
import pandas as pd
import Dataclasses as dc
import Fredon_projects_fields as pf
import streamlit as st
import datetime
import openpyxl as px
from openpyxl import Workbook

def get_data_columns(df):
    return df.columns.tolist()

def remove_columns(df, columns_to_remove):
    return df.drop(columns=columns_to_remove)

def get_list_of_projects(df):
    # Get the column name by index (column 3 = PROJECT_NAME)
    project_col = df.columns[pf.PROJECT_NAME] if pf.PROJECT_NAME < len(df.columns) else None
    
    if project_col is None:
        print(f"Column index {pf.PROJECT_NAME} not found in DataFrame")
        return []
    
    df_clean = df.dropna(subset=[project_col])
    df_clean = df_clean[df_clean[project_col].astype(str).str.strip() != ""]
    return df_clean[project_col].unique().tolist()

def get_projects_with_highest_complete(df):
    # Get column names by index
    project_col = df.columns[pf.PROJECT_NAME] if pf.PROJECT_NAME < len(df.columns) else None
    
    # Find the "%Complete" or "% Complete" column
    complete_col = None
    for col in df.columns:
        col_str = str(col).lower().strip()
        if "% complete" in col_str or "%complete" in col_str or "complete" in col_str:
            complete_col = col
            # st.write(f"Found complete column: '{col}' at index {df.columns.get_loc(col)}")
            break
    
    if project_col is None or complete_col is None:
        # st.write(f"Debug: project_col={project_col}, complete_col={complete_col}")
        # st.write("Available columns with 'complete' in name:")
        # for i, col in enumerate(df.columns):
        #     if "complete" in str(col).lower():
        #         st.write(f"  {i}: '{col}'")
        return []
    
    # st.write(f"Debug: Using project column '{project_col}' and complete column '{complete_col}'")
    
    # Clean the data
    df_clean = df.dropna(subset=[project_col, complete_col])
    
    # Convert %Complete to numeric if it's not already
    if df_clean[complete_col].dtype == 'object':
        # Remove % signs and convert to float
        df_clean = df_clean.copy()
        df_clean[complete_col] = df_clean[complete_col].astype(str).str.replace('%', '').str.replace(',', '').str.strip()
        df_clean[complete_col] = pd.to_numeric(df_clean[complete_col], errors='coerce')
        df_clean = df_clean.dropna(subset=[complete_col])
    
    # If values are between 0-1 (like 0.95), convert to percentage (like 95)
    max_val = df_clean[complete_col].max()
    if max_val <= 1.0:
        df_clean[complete_col] = df_clean[complete_col] * 100
    
    if df_clean.empty:
        # st.write("Debug: No valid data after cleaning")
        return []
    
    # st.write(f"Debug: Sample %Complete values: {df_clean[complete_col].head().tolist()}")
    # st.write(f"Debug: Min: {df_clean[complete_col].min():.2f}, Max: {df_clean[complete_col].max():.2f}")
    
    # Group by project and get the row with maximum % Complete for each project
    idx = df_clean.groupby(project_col)[complete_col].idxmax()
    result = df_clean.loc[idx, [project_col, complete_col]].copy()
    
    # st.write(f"Debug: Found {len(result)} projects with max %Complete")
    
    # Show the results for debugging
    # for _, row in result.head().iterrows():
    #     st.write(f"  {row[project_col]}: {row[complete_col]:.2f}%")
    
    # Add status column
    result["Status"] = result[complete_col].apply(lambda x: "Calibrate" if x > 0.95 else "Operational")
    
    # Format % Complete as a percentage string (but keep original numeric value for comparison)
    result["% Complete Formatted"] = result[complete_col].apply(lambda x: f"{x:.2f}%" if pd.notnull(x) else "")
    
    # Return as list of tuples: (Project_Name, Max_%, Status)
    projects = []
    for _, row in result.iterrows():
        projects.append((row[project_col], row["% Complete Formatted"], row["Status"]))
    
    return projects

def create_data_objects(df):
    project_names = get_list_of_projects(df)
    
    if not project_names:
        st.write("No project names found")
        return pd.DataFrame()
    
    # Get the project column name by index
    project_col = df.columns[pf.PROJECT_NAME] if pf.PROJECT_NAME < len(df.columns) else None
    
    if project_col is None:
        st.write(f"Column index {pf.PROJECT_NAME} not found in DataFrame")
        return pd.DataFrame()
    
    # Create portfolio to hold all projects
    Portfolio = dc.Portfolio()
    all_project_data = pd.DataFrame()
    
    st.write(f"Processing {len(project_names)} projects...")
    
    # Loop through all project names
    for project_idx, project_name in enumerate(project_names):
        st.write(f"Processing project {project_idx + 1}/{len(project_names)}: {project_name}")
        
        # Filter data for current project
        project_data = df[df[project_col] == project_name]
        
        if project_data.empty:
            st.write(f"No data found for project: {project_name}")
            continue
        
        try:
            project_object = dc.Projects(
                Project_Name=project_data.iloc[0, pf.PROJECT_NAME],
                Project_ID=project_data.iloc[0, pf.PROJECT_ID] if pf.PROJECT_ID < len(df.columns) else "Unknown",
                Location=project_data.iloc[0, pf.LOCATION] if pf.LOCATION and pf.LOCATION < len(df.columns) else "Unknown",
                Post_Code=project_data.iloc[0, pf.POST_CODE] if pf.POST_CODE and pf.POST_CODE < len(df.columns) else "Unknown",
                Sector=project_data.iloc[0, pf.CLIENT_SECTOR] if pf.CLIENT_SECTOR < len(df.columns) else "Unknown",
                Portfolio_Bus_Unit_Dept_ID=project_data.iloc[0, pf.PORTFOLIO_BUS_UNIT_DEPT_ID] if pf.PORTFOLIO_BUS_UNIT_DEPT_ID < len(df.columns) else "Unknown",
                Asset_Type=project_data.iloc[0, pf.ASSET_TYPE] if pf.ASSET_TYPE and pf.ASSET_TYPE < len(df.columns) else "Unknown",
                Contract_Type=project_data.iloc[0, pf.CONTRACT_TYPE] if pf.CONTRACT_TYPE and pf.CONTRACT_TYPE < len(df.columns) else "Unknown",
                Contract_Financial=project_data.iloc[0, pf.CONTRACT_FINANCIAL] if pf.CONTRACT_FINANCIAL < len(df.columns) else "Unknown",
                Client=project_data.iloc[0, pf.CLIENT] if pf.CLIENT < len(df.columns) else "Unknown",
                Stage_of_Work=project_data.iloc[0, pf.STAGE_OF_WORK] if pf.STAGE_OF_WORK < len(df.columns) else "Unknown",
                Template_Version="0.2",
                Status="Operational"
            )
            
            # Process monthly data for this project
            number_of_months = project_data.shape[0]
            for mon in range(number_of_months):
                try:
                    # Handle date conversion properly
                    if pf.DATE and pf.DATE < len(df.columns):
                        date_value = project_data.iloc[mon, pf.DATE]
                        if pd.isna(date_value):
                            record_date = ""  # Empty string if null
                        else:
                            # Convert pandas Timestamp to formatted string
                            if hasattr(date_value, 'date'):
                                record_date = date_value.strftime("%d/%m/%Y")  # Format as dd/mm/yyyy
                            elif isinstance(date_value, str):
                                # Handle string format like "2019-11-30 00:00:00"
                                try:
                                    parsed_date = pd.to_datetime(date_value)
                                    record_date = parsed_date.strftime("%d/%m/%Y")  # Format as dd/mm/yyyy
                                except:
                                    record_date = ""  # Empty string if parsing fails
                            else:
                                record_date = ""  # Empty string for other types
                    else:
                        record_date = ""  # Empty string if no date column
                    
                    monthly_record = dc.MonthlyRecord(
                        Date=record_date,
                        approved_budget=float(project_data.iloc[mon, pf.APPROVED_BUDGET]) if pf.APPROVED_BUDGET and pf.APPROVED_BUDGET < len(df.columns) else 0.0,
                        forecast_end_date=float(project_data.iloc[mon, pf.FORECAST_END_DATE]) if pf.FORECAST_END_DATE and pf.FORECAST_END_DATE < len(df.columns) else 0.0,
                        forecast_final_cost=float(project_data.iloc[mon, pf.FORECAST_FINAL_COST]) if pf.FORECAST_FINAL_COST and pf.FORECAST_FINAL_COST < len(df.columns) else 0.0,
                        contingency_remaining=float(project_data.iloc[mon, pf.CONTINGENCY_REMAINING]) if pf.CONTINGENCY_REMAINING and pf.CONTINGENCY_REMAINING < len(df.columns) else 0.0,
                        actual_cost_to_date=float(project_data.iloc[mon, pf.ACTUAL_COST_TO_DATE]) if pf.ACTUAL_COST_TO_DATE and pf.ACTUAL_COST_TO_DATE < len(df.columns) else 0.0,
                        forecast_final_revenue=float(project_data.iloc[mon, pf.FORECAST_FINAL_REVENUE]) if pf.FORECAST_FINAL_REVENUE and pf.FORECAST_FINAL_REVENUE < len(df.columns) else 0.0,
                        actual_revenue_to_date=float(project_data.iloc[mon, pf.ACTUAL_REVENUE_TO_DATE]) if pf.ACTUAL_REVENUE_TO_DATE and pf.ACTUAL_REVENUE_TO_DATE < len(df.columns) else 0.0,
                        notes=str(project_data.iloc[mon, pf.NOTES]) if hasattr(pf, 'NOTES') and pf.NOTES and pf.NOTES < len(df.columns) else None,
                        Comments=str(project_data.iloc[mon, pf.COMMENTS]) if hasattr(pf, 'COMMENTS') and pf.COMMENTS and pf.COMMENTS < len(df.columns) else None
                    )
                    project_object.Monthly_data.append(monthly_record)
                except (ValueError, IndexError, TypeError) as e:
                    st.write(f"Error processing month {mon+1} for project {project_name}: {e}")
                    continue
            
            # Add project to portfolio
            Portfolio.add_project(project_object)
            
            # Add column numbers to project data and append to all_project_data
            numbered_columns = {col: f"Col {i+1}: {col}" for i, col in enumerate(project_data.columns)}
            project_data_numbered = project_data.rename(columns=numbered_columns)
            all_project_data = pd.concat([all_project_data, project_data_numbered], ignore_index=True)
            
        except Exception as e:
            st.write(f"Error processing project {project_name}: {e}")
            continue
    
    st.write(f"Successfully processed {len(Portfolio.projects)} projects")
    
    # Display project summary
    st.write("Portfolio Summary:")
    st.data_editor(Portfolio.to_dataframe(), use_container_width=True)
    
    # Display monthly data
    st.write("All Monthly Data:")
    st.data_editor(Portfolio.monthly_data_to_dataframe(), use_container_width=True)
    
    # Add Excel export button
    if st.button("Export Portfolio to Excel"):
        filename = create_excel_file_with_portfolio_data(Portfolio)
        st.success(f"Excel file created with {len(Portfolio.projects)} projects: {filename}")
        
        # Offer download
        with open(filename, "rb") as file:
            st.download_button(
                label="Download Portfolio Excel File",
                data=file.read(),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    return all_project_data


  
    
    
def create_excel_file_with_portfolio_data(portfolio, filename="portfolio_output.xlsx"):
    # Create a new workbook
    workbook = px.Workbook()
    
    # Remove the default sheet
    workbook.remove(workbook.active)
    
    # Create a worksheet for each project
    for idx, project_object in enumerate(portfolio.projects):
        # Create worksheet with project name (sanitized for Excel)
        sheet_name = project_object.Project_Name[:31] if len(project_object.Project_Name) <= 31 else project_object.Project_Name[:28] + "..."
        # Remove invalid characters for Excel sheet names
        invalid_chars = ['\\', '/', '*', '[', ']', ':', '?']
        for char in invalid_chars:
            sheet_name = sheet_name.replace(char, '_')
        
        sheet = workbook.create_sheet(title=sheet_name)
        
        # Template headers
        sheet["A1"] = "TEMPLATE VERSION:"
        sheet["B1"] = project_object.Template_Version
        sheet["H1"] = "OCTANT BUSINESS - PROJECT DATA UPLOAD TEMPLATE"
        sheet["F2"] = "PROJECT SUMMARY"
        
        # Project summary labels
        sheet["A4"] = "Project Name*"
        sheet["A5"] = "Project ID*"
        sheet["A6"] = "Location*" 
        sheet["A7"] = "Post Code"   
        sheet["A8"] = "Sector*"
        sheet["A9"] = "Portfolio/ Bus Unit/ Dept ID*"
        sheet["A11"] = "Comments this period"
        sheet["A12"] = "(max 200 Characters )"
        sheet["H4"] = "Asset Type*"
        sheet["H5"] = "Contract Type"
        sheet["H6"] = "Contract Financial"
        sheet["H7"] = "Client*"
        sheet["H8"] = "Stage of Work*"
        
        # Project summary data
        sheet["C4"] = project_object.Project_Name
        sheet["C5"] = project_object.Project_ID
        sheet["C6"] = project_object.Location
        sheet["C7"] = project_object.Post_Code
        sheet["C8"] = project_object.Sector
        sheet["C9"] = project_object.Portfolio_Bus_Unit_Dept_ID
        sheet["J4"] = project_object.Asset_Type
        sheet["J5"] = project_object.Contract_Type
        sheet["J6"] = project_object.Contract_Financial
        sheet["J7"] = project_object.Client
        sheet["J8"] = project_object.Stage_of_Work
        
        # Instructions
        sheet["A15"] = "Instructions: Please overwrite all mandatory cells with your own business data"
        sheet["A16"] = "You may also upload this template directly as sample data to experience"
        sheet["A17"] = "Octant AI with sample operational data"
        
        # Monthly data headers
        sheet["F20"] = "MONTHLY DATA"
        sheet["A21"] = "Reporting Date*"
        sheet["B21"] = "Approved Budget (including contingency)"
        sheet["C21"] = "Forecast End Date (practical completion)"
        sheet["D21"] = "Forecast Final Cost*"
        sheet["E21"] = "Contingency Remaining"
        sheet["F21"] = "Actual Cost to Date*"
        sheet["G21"] = "Forecast Final Revenue*"
        sheet["H21"] = "Actual Revenue to Date*"
        sheet["I21"] = "Notes"
        sheet["J21"] = "Comments"
        
        # Monthly data rows
        start_row = 22
        for record_idx, monthly_record in enumerate(project_object.Monthly_data):
            row = start_row + record_idx
            sheet[f"A{row}"] = monthly_record.Date
            sheet[f"B{row}"] = monthly_record.approved_budget
            sheet[f"C{row}"] = monthly_record.forecast_end_date
            sheet[f"D{row}"] = monthly_record.forecast_final_cost
            sheet[f"E{row}"] = monthly_record.contingency_remaining
            sheet[f"F{row}"] = monthly_record.actual_cost_to_date
            sheet[f"G{row}"] = monthly_record.forecast_final_revenue
            sheet[f"H{row}"] = monthly_record.actual_revenue_to_date
            sheet[f"I{row}"] = monthly_record.notes
            sheet[f"J{row}"] = monthly_record.Comments
    
    # Save the workbook
    workbook.save(filename)
    return filename

def debug_dataframe_structure(df):
    """Debug function to analyze DataFrame structure and find percentage columns"""
    st.write("=== DEBUGGING DATAFRAME STRUCTURE ===")
    st.write(f"DataFrame shape: {df.shape}")
    st.write(f"Column count: {len(df.columns)}")
    
    # Show first few columns with their indices
    st.write("Column indices and names:")
    for i, col in enumerate(df.columns[:20]):  # Show first 20 columns
        st.write(f"  {i}: '{col}'")
    
    # Look for Complete columns specifically
    st.write("\nColumns containing 'complete':")
    complete_cols = []
    for i, col in enumerate(df.columns):
        if "complete" in str(col).lower():
            complete_cols.append((i, col))
            st.write(f"  {i}: '{col}'")
    
    if complete_cols:
        # Show sample data from percentage columns
        for col_idx, col_name in complete_cols[:3]:  # Show first 3 matching columns
            st.write(f"\nSample data from column '{col_name}':")
            sample_data = df[col_name].dropna().head(10)
            st.write(sample_data.tolist())
            st.write(f"Data type: {df[col_name].dtype}")
    else:
        st.write("No columns found containing 'complete'")
    
    # Show project column data
    if pf.PROJECT_NAME < len(df.columns):
        project_col = df.columns[pf.PROJECT_NAME]
        st.write(f"\nProject column (index {pf.PROJECT_NAME}): '{project_col}'")
        unique_projects = df[project_col].dropna().unique()
        st.write(f"Unique projects ({len(unique_projects)}): {unique_projects[:5].tolist()}")
    
    st.write("=== END DEBUGGING ===")