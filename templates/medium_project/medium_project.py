"""
This is to create the 'medium', in which the individual sheets contain a reasonable amount of logic each. 

This is the appropriate template to adapt when you want to write and Excel file in which individual sheets draw on data from two or more different sources.

This builds on the logic and functionality from the 'Easy' project. 

Advice on how to adapt this to your own project can be found in the README. 
"""
from pathlib import Path
import openpyxl
import utils



template_path = Path('templates/medium_project/medium_template.xlsx')

def make_excel_output() -> None:
    # Set Up
    output_path = Path('outputs/medium_output.xlsx')
    wb = openpyxl.load_workbook(template_path) # Make and Write Each Sheet

    wb = utils.make_and_write_2a(wb=wb)
    wb = utils.make_and_write_2b(wb=wb)
    wb = utils.make_and_write_2c(wb=wb)
    wb = utils.make_and_write_2d(wb=wb)
    wb = utils.make_and_write_3a(wb=wb)
    wb = utils.make_and_write_3b(wb=wb)
    wb = utils.make_and_write_3c(wb=wb)
    wb = utils.make_and_write_3d(wb=wb)

    wb = utils.make_and_write_table_5(wb=wb)


    wb.save(output_path)
    print("Medium project: Excel file written")
