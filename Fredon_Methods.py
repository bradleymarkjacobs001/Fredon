#import openpyxl as px
import pandas as pd

def get_data_columns(df):
    return df.columns.tolist()

def remove_columns(df, columns_to_remove):
    return df.drop(columns=columns_to_remove)

def get_list_of_projects(df):
    
    df = df.dropna(subset=["Contract Name (Project Name)"])
    df = df[df["Contract Name (Project Name)"].str.strip() != ""]
    return df["Contract Name (Project Name)"].unique().tolist()

def get_projects_with_highest_complete(df):
    df = df.dropna(subset=["Contract Name (Project Name)", "% Complete"])
    idx = df.groupby("Contract Name (Project Name)")["% Complete"].idxmax()
    result = df.loc[idx, ["Contract Name (Project Name)", "% Complete"]].copy()
    # Add status column
    result["Status"] = result["% Complete"].apply(lambda x: "Calibrate" if x > .95 else "Operational")
    # Format % Complete as a percentage string
    result["% Complete"] = result["% Complete"].apply(lambda x: f"{x:.2f}%" if pd.notnull(x) else "")
    projects = list(result.itertuples(index=False, name=None))
    return projects

def create_data_objects(df):
    project_names = get_list_of_projects(df)
    
    #for project in project_names:
    project_data = df[df["Contract Name (Project Name)"] == project_names[0]]
    return project_data    

def create_excel_file(df, filename="output.xlsx", sheet_name="Sheet1"):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        
    return filename