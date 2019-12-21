# Excel to Database converter
This is a python program that converts excel files into database files.

This program reads the data to be inserted from the excel file. After reading the data, it creates the table in the database and inserts the elements into the database.

# Requirements
This program uses one external module and one in-built module:
  * openpyxl
  * sqlite3

# Instructions for use - How the format of the Excel of the file should be
* The name of the table should be in the first cell of the worksheet (i.e.) in the cell "A1"
* The names of the columns should be in the next row (i.e.) the row number 2
* All the columns must have a corresponding colummn name in row number 2
* All data to be entered normally under each column

# How to use the converter
* Run the *main.py* file to open the GUI application.
* Enter the name of the excel file to be converted in the field provided (the excel file should be present in the same directory as the main.py file).
* Enter the name of the database file in which the data will be input in the corresponding field provided.

**NOTE**: Do not leave these fields empty

* Click on the **Start Conversion** button to covert the excel file.
* The resulting database file will be present in the same directory as *main.py*

# Misc. Info
* To use this program, the user should have the excel file to be placed in the same directory as the python file (main.py) is present.
* The only file formats supported are *xlsx / xlsm / xltx / xltm*.
* The program will output the database file in the same directory as the python file (main.py) is present.
* Currently only the first worksheet will be converted into corresponding database
