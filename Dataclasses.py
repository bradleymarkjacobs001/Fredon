import datetime
from dataclasses import dataclass, field
from pandas import DataFrame
from typing import List, Optional

@dataclass
class MonthlyRecord:
    Date: str
    approved_budget: float
    forecast_final_cost: float
    actual_cost_to_date: float
    forecast_final_revenue: float
    actual_revenue_to_date: float
    forecast_end_date: Optional[str] = ""
    contingency_remaining: Optional[float] = None
    notes: Optional[str] = None
    
@dataclass
class Projects:
    Project_Name: str
    Project_ID: str
    Location: str
    
    Sector: str
    Portfolio_Bus_Unit_Dept_ID: str
    Asset_Type: str
    Client: str
    Stage_of_Work: str
    Template_Version: str = "0.2"
    Status: str = "Operational"
    Monthly_data: List[MonthlyRecord] = field(default_factory=list)
    Contract_Type: Optional[str] = None
    Contract_Financial: Optional[str] = None
    Post_Code: Optional[str] = None
    Comments: Optional[str] = None

    

class Portfolio:
    def __init__(self, projects: List[Projects] = []):
        self.projects = projects or []

    def add_project(self, project: Projects):
        self.projects.append(project)

    def to_dataframe(self) -> DataFrame:
        project_dicts = []
        for project in self.projects:
            project_dict = {
                'Project_Name': project.Project_Name,
                'Project_ID': project.Project_ID,
                'Location': project.Location,
                'Post_Code': project.Post_Code,
                'Sector': project.Sector,
                'Portfolio_Bus_Unit_Dept_ID': project.Portfolio_Bus_Unit_Dept_ID,
                'Asset_Type': project.Asset_Type,
                'Contract_Type': project.Contract_Type,
                'Contract_Financial': project.Contract_Financial,
                'Client': project.Client,
                'Stage_of_Work': project.Stage_of_Work,
                'Template_Version': project.Template_Version,
                'Status': project.Status,
                'Number_of_Monthly_Records': len(project.Monthly_data),
                'Comments': project.Comments
            }
            project_dicts.append(project_dict)
        return DataFrame(project_dicts)
    
    def monthly_data_to_dataframe(self) -> DataFrame:
        monthly_records = []
        for project in self.projects:
            for record in project.Monthly_data:
                record_dict = {
                    'Project_Name': project.Project_Name,
                    'Date': record.Date,
                    'Approved_Budget': record.approved_budget,
                    'Forecast_End_Date': record.forecast_end_date,
                    'Forecast_Final_Cost': record.forecast_final_cost,
                    'Contingency_Remaining': record.contingency_remaining,
                    'Actual_Cost_To_Date': record.actual_cost_to_date,
                    'Forecast_Final_Revenue': record.forecast_final_revenue,
                    'Actual_Revenue_To_Date': record.actual_revenue_to_date,
                    'Notes': record.notes,
                      # Assuming Comments is a project-level attribute
                }
                monthly_records.append(record_dict)
        return DataFrame(monthly_records)