import streamlit as st
import openpyxl as px
import pandas as pd
import Fredon_Methods_test as fm
import Dataclasses as dc

   
def main():
    
    st.title("Excel to DataFrame Converter")
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])
    
    # Initialize df as None
    df = None
    
    if uploaded_file is not None:
        try:
            # Load the Excel file
            workbook = px.load_workbook(uploaded_file, data_only=True,)
            sheet_names = workbook.sheetnames
            
            # Select a sheet
            selected_sheet = st.selectbox("Select a sheet", sheet_names)
            
            # Read the selected sheet into a DataFrame
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=1)
            
            # Display the DataFrame
            st.write(f"Data from sheet: {selected_sheet}")
            st.dataframe(df)
            
            # Debug the data structure (commented out for production)
            # fm.debug_dataframe_structure(df)
        except Exception as e:
            st.error(f"Error reading the file: {e}")
            return
    
    # Only show these options if df is loaded
    if df is not None:
        # df_updated = st.multiselect("Select columns to remove", options=fm.get_data_columns(df), key="columns_to_remove")  
        # if df_updated:
        #     df = fm.remove_columns(df, df_updated)
        #     st.write("Updated DataFrame after removing selected columns:")
        #     st.dataframe(df)  
        
        # df_projects = fm.get_list_of_projects(df)
        # if df_projects: 
        #     st.write("List of unique project names:")
        #     st.data_editor(pd.DataFrame(df_projects,columns=["Project Name"]), use_container_width=True)  
        
        df_highest_complete = fm.get_projects_with_highest_complete(df) 
        if df_highest_complete:
            st.write("Projects with the highest % Complete:")
            st.data_editor(pd.DataFrame(df_highest_complete, columns=["Project ID", "Project Name", "% Complete", "Status"]), use_container_width=True)
        else:
            st.write("No projects found with complete data.")
        
        # st.dataframe(fm.create_data_objects(df), use_container_width=True)
        fm.create_data_objects(df)
if __name__ == "__main__":
    main()
    
    
    

    
