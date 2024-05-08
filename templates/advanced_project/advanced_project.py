"""
Here we define the functions for creating the 'advanced' project. 

This is more complicated than the medium project in that it includes a heavily formatted summary sheet - here, this is the sheet titled 'Table 1'. 
Many NHSD publications have a sheet of this type, and they do not lend themselves to simple automation. 
The solution we have landed on here involved writing data cell-by-cell, to ensure accuracy. 

We only recommend adapting this template if you are working to automate a publication which already includes a sheet similar to the 'Table 1' sheet produced here. 
If your publication does not include a sheet like this, we recommend sticking to data formatted in a simpler manner, such as in the other two example projects. 

Advice on how to adapt this project to your own can be found in the README. 
"""
from pathlib import Path
import openpyxl

import utils
from templates.advanced_project import table_1



template_path = Path('templates/advanced_project/advanced_template.xlsx')

def make_excel_output() -> None:
    """Creates and writes the Excel file for the 'advanced' project. 
    """    
    # Set Up
    output_path = Path('outputs/advanced_output.xlsx')
    wb = openpyxl.load_workbook(template_path) # Make and Write Each Sheet

    # Make the Excel file
    wb = table_1.make_and_write_table1(wb=wb)
    wb = utils.make_and_write_2a(wb=wb)
    wb = utils.make_and_write_2b(wb=wb)
    wb = utils.make_and_write_2c(wb=wb)
    wb = utils.make_and_write_2d(wb=wb)
    wb = utils.make_and_write_3a(wb=wb)
    wb = utils.make_and_write_3b(wb=wb)
    wb = utils.make_and_write_3c(wb=wb)
    wb = utils.make_and_write_3d(wb=wb)

    wb = utils.make_and_write_table_5(wb=wb)

    # Save

    wb.save(output_path)
    print("Advanced Project: Excel file written")
