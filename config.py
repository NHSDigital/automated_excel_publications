import datetime
from pathlib import Path
import pandas

"""
This config file exists to set configuration parameters for your project. 
For the sake of simplicity, we have 'hard-coded' parameters into functions here; for example, the date-time is set within the function below. 
If you are adapting this template out into a project, we recommend parametrising your variables into a separate file. 
"""

def get_report_month():
    return datetime.date(year=2022, month=4, day=1)

def get_number_of_months():
    return 12

def get_easy_a_data():
    filepath = Path('data/data_for_sheet_easy_a.csv')
    return pandas.read_csv(filepath)

def get_easy_b_data():
    filepath = Path('data/data_for_sheet_easy_b.csv')
    return pandas.read_csv(filepath)

def get_appointments_data():
    filepath = Path('data/appointment_data.csv')
    return pandas.read_csv(filepath)

def get_practices_data():
    filepath = Path('data/practices_data.csv')
    return pandas.read_csv(filepath)

def get_table1_data():
    filepath = Path('data/table1_data.csv')
    return pandas.read_csv(filepath)
