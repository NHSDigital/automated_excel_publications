"""
This is the `easy` project file. 

This is the most template project with the most straightforward logic: the data written to the sheets in the output Excel file directly represents a filtered selection from a CSV file. We are simply loading a CSV file; selecting some data, and writing that data to a sheet. There is very little logical manipulation of the data, and very little additional formatting being applied. 

For info on how to adapt this template to your own project, see the README.
"""
from pathlib import Path
import openpyxl
import utils


template_path = Path("templates/easy_project/easy_template.xlsx")


def make_excel_output() -> None:
    """Creates and writes the output Excel file for the easy project

    Returns:
        None:
    """

    # Set Up
    output_path = Path("outputs/easy_output.xlsx")
    wb = openpyxl.load_workbook(template_path)  # Make and Write Each Sheet

    # Make the sheets
    wb = utils.make_and_write_easy_a(wb=wb)
    wb = utils.make_and_write_easy_b(wb=wb)

    # Write the workbook
    wb.save(output_path)
    print("Easy project: Excel file written")
    return None
