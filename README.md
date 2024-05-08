# Excel Automation

This is an example of a module used to automate the production of Excel files.
NHS England publishes a number of Excel files, and does so with a significant amount of formatting.

The repo is organised around three example projects, in progressing levels of complexity. The `advanced` project is very close to the real publication [here](https://digital.nhs.uk/data-and-information/publications/statistical/appointments-in-general-practice/march-2022); the other two examples pare this project down to make it more simple, straightforward, and concise.

> **Note**
> Being adapted from the code used to produce the [Appointments in General Practice](https://digital.nhs.uk/data-and-information/publications/statistical/appointments-in-general-practice) publications, the sample templates used in this repo work with publicly available historical data.

## Ownership

NHS England and NHS Digital have [merged](https://digital.nhs.uk/about-nhs-digital/nhs-digital-merger-with-nhs-england).

**Repository owner**: [NHS England Data Science](https://nhsengland.github.io/datascience/)

**Email**: datascience@nhs.net

_To contact us raise an issue on Github or via email._

## File Structure and Overview

```text
README.md
templates
   |-- easy_project
   |   |-- __init__.py
   |   |-- easy_project.py
   |   |-- easy_template.xlsx
   |-- medium_project
   |   |-- __init__.py
   |   |-- medium_project.py
   |   |-- medium_template.xlsx
   |-- advanced_project
   |   |-- __init__.py
   |   |-- advanced_project.py
   |   |-- advanced_template.xlsx
   |   |-- table1_cell_tags.py
   |   |-- table_1.py
data
   |-- appointment_data.csv
   |-- practices_data.csv
   |-- table1_data.csv
outputs
   |-- .gitkeep
   |-- advanced_output.xlsx
   |-- easy_output.xlsx
   |-- medium_output.xlsx
main.py
config.py
requirements.txt
utils.py
```



## Installation

The [RAP Community of Practice resources](https://nhsdigital.github.io/rap-community-of-practice/) can help you if you are unsure about any of the steps below, including [cloning repositories](https://nhsdigital.github.io/rap-community-of-practice/training_resources/git/introduction-to-git/#common-git-commands), and [virtual environments](https://nhsdigital.github.io/rap-community-of-practice/training_resources/python/virtual-environments/why-use-virtual-environments/).

1. Clone this repo to your preferred working destination, and point your terminal to the folder you've cloned.
2. Next, create and activate a python virtual environment using your preferred tool.
3. Next we need to install the packages which this project requires. In your terminal, enter

```bash
    pip install -r requirements.txt
```

and run the command. If you run into any problems, check that your terminal is pointed to the right place, and that your virtual environment is activated.

Your installation should now be complete, and you should be able to run the project.

The basic logic of all three example projects is the same:

1. `main.py` calls a `make_excel` function in the relevant project's script.
2. This function loads in the template `.xlsx` file for the example project.
3. It then uses various functions from `excel_functions`; one for each sheet of the target template publication.

These functions all follow the same basic logic, they:

1. Open the relevant data file, as specified in the `config` file.
2. Select the relevant columns, join and re-order if necessary.
3. The function specifies a sheet name: open this sheet in the template, and find the cell with `<start>` in it.
4. Write the data to this sheet, starting at this cell.
5. Find the cell with `<end>` written in it.
6. Delete all rows between the last row which we've written data to, and this `<end>` cell.

> This last step needs further explanation. The nature of the publication is such that the number of rows printed each publication might vary; different months have different numbers of days, new regions might be added to the scope, etc. Given this, the most practical solution is to allow for an over-abundance of white space in the `template` document, and then delete as appropriate.

## Easy Project

This project writes two simple sheets: `2a` and `2b`. The functions for writing these sheets are straightforward: select the relevant data, and write it to the workbook.

### How To Run the Easy Project

1. Open `main.py`
2. Comment out the lines which call the other projects:

    ```python
    #medium_project.make_excel_output()
    #advanced_project.make_excel_output()
    ```

    by adding `#` to the beginning of each line
3. Run

    ```bash
    python main.py
    ```

    in the terminal.

### What the Easy Project is Appropriate For

This project is the appropriate starting point when you have a number of separate CSV files, each of which you want to write to a separate sheet.

### How To Adapt the Easy Project

The easy project is very simple to adapt. The `data_for_sheet_easy_a.csv` file is written, almost directly, to sheet `Easy A` in the output file.

1. Copy the template .xlsx file, and duplicate the example sheets within your new version, naming the sheets and their columns appropriately
2. Add your CSV files to the `data` folder, and add functions to the `config.py` file to load these
3. Copy and rename the `make_and_write_easy_a` function in `utils.py`, and adapt it to your sheet and data source.
4. Open `easy_project.py` and replace the functions called within `make_excel_output` with your new functions, and change the template path to your new template.

## Medium Project

This adds a little complexity, in that we will now be handling data from multiple sources.
Sheets 3.x require data from the `appointments` and `practices` CSV files to be joined, according to region  - This is handled by the `combine_practices_with_appts` function. Once the data from these has been joined, we can simply write it to the appropriate sheet.

### How To Run the Medium Project

1. Open `main.py`
2. Comment out the lines which call the other projects:

    ```python
    #easy_project.make_excel_output()
    #advanced_project.make_excel_output()
    ```

    by adding `#` to the beginning of each line
3. Run

    ```bash
    python main.py
    ```

    in the terminal.

### What the Medium Project is Appropriate For

This project is the appropriate starting point when you have CSV files which contain more information than you want o appear on any given individual sheet. The functions used in this project do more manipulation on the data once it's been loaded in. Some of the sheets produced here also require that data be joined from two separate CSV files. As such, the functions for producing each sheet here contain more logic than in the 'easy' project.

### How To Adapt the Medium Project

As in the easy project, load your data into the `data` folder, and create the corresponding functions in `config` to load your data in.

Then, in `utils.py`, create the functions which curates (and possibly join) the data you need for your sheets.

As you can see in the functions we've got here, we've given the data in the CSV files a `breakdown` column: this means that we're able to easily identify the relevant rows from a 'long' dataset without much logic.

Now, place your template `.xlsx` file in the project folder, making sure that your target sheets contain the `<start>` and `<end>` tags, as in the example template.

Once this is done, you can adapt or replace the `medium_project.py` file, and call it from `main.py` in the usual way.

## Advanced Project

This involved a significant step up in complexity from the medium project.
Many NHS England publications contain a summary sheet which is heavily formatted, and which contains information about data which will appear in the rest of the sheets in the publication at a more granular level.

The other sheets have a data layout which more-or-less echoes a Pandas dataframe, and as such a single dataframe can be written without much manipulation.
Table 1 on the other hand has a number of specific formatting and presentation criteria.
These include; including a range of months in the sheet, giving most information twice (once as a count, once as a percentage), displaying the information in a form which is the transpose of the convention (here the data **fields** are the **rows**, and the **entries** are the **columns**), having multiple blank spaces, and so on.

All of these features present us with difficulties for formatting this sheet. Should we create a very wide dataframe, and write its transpose? Or many small ones, and write them one by one?

The solution we have implemented is to assign each and every cell in a given 'column' a unique cell identifier, and to provide a matching CSV file specifically for the purpose of populating Table 1.
The functions in `table_1.py` then search through for these specific tags and replace them with the appropriate values.
This allows us to minimise the risk of introducing errors, and use a simple function in a loop to write each cell of the sheet.

### How To Run the Advanced Project

1. Open `main.py`
2. Comment out the lines which call the other projects:

    ```python
    #easy_project.make_excel_output()
    #medium_project.make_excel_output()
    ```

    by adding `#` to the beginning of each line
3. Run

    ```bash
    python main.py
    ```

    in the terminal.

### What the Advanced Project is Appropriate For

This project is more specific than the other two. It is appropriate for adapting to produce NHS England publications which contain a heavily formatted, high-level, summary sheet.

### How To Adapt the Advanced Project

This involved more work than adapting the other two projects.

We will deal only with how to adapt the summary page, the `Table 1` sheet, since the other sheets are equivalent to those in the 'medium' project.

1. Make a copy of the publication which you are trying to automate. Leave all of the headers and column names intact, but delete all of the numbers/data from the cells.
2. Next, fill those empty cells with written tags which correspond meaningfully to the data which will be written there. This is less laborious than it might look; using the column names to populate the tags, and copious amounts of copy-pasting, you can populate your template in < 10 minutes.

The data used to populate `Table 1` in the example project is specifically written for this purpose; every cell tag combined with a `month` yields a single value. Writing your CSV data in this format makes it much easier to incorporate into this pipeline, and save you from having to write much logic here.

If your summary sheet is formatted in a way similar to the example, then you will be able to implement it without changing too much of the code. You will need to create an equivalent to `cell_1_tags.py`, and populate it with the tags from your project.

You might also need to replace `month` with whichever index it is that you are iterating over, and adjust the logic accordingly.
