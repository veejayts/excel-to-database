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

# Misc. Info
To use this program, the user should have the excel file to be placed in the same directory as the python file is present.
The only file formats supported are *xlsx / xlsm / xltx / xltm*.
The program will output the database file in the same directory as the python file is present.
