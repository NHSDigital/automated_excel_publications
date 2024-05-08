from templates.advanced_project import advanced_project 
from templates.medium_project import medium_project
from templates.easy_project import easy_project

def main():
    easy_project.make_excel_output()
    medium_project.make_excel_output()
    advanced_project.make_excel_output()


if __name__ == "__main__":
    main()