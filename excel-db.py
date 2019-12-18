import openpyxl

import sqlite3

def readData():
    wb = openpyxl.load_workbook('test.xlsx')

    # To get all the worksheets
    sheets = wb.sheetnames
    # To select the individual sheet
    sheet = wb['Sheet1']

    # Number of rows and columns
    no_of_row = sheet.max_row + 1
    no_of_col = sheet.max_column + 1

    # The start values for the rows and columns respectively
    row_start = 2
    col_start = 1

    # Loop variables
    row_counter = 1
    col_counter = 1

    # List for storing the names of the columns
    col_names = list()

    # Getting the column names
    for col_counter in range(col_start, no_of_col):
        col_names.append(sheet.cell(col_start, col_counter).value)

    # List for storing the data of each rows
    row_data = list()

    # Loop to iterate over each row
    for row_counter in range(row_start, no_of_row):
        # Temporary list to store the data of the individual row and then append it to the row_data
        # so that temp_row_data can be rewritten
        temp_row_data = list()
        # Loop to iterate over the columns
        for col_counter in range(col_start, no_of_col):
            # Appending the data of the particular row into the temporary list
            temp_row_data.append(sheet.cell(row_counter, col_counter).value)
        # Appending the temp_row_data to row_data
        row_data.append(temp_row_data)

conn = sqlite3.connect('test.db')

c = conn.cursor()

c.execute("""create table test (
    first text,
    last text,
    pay integer
)""")

# c.execute("""insert into test values('vijay', 't s', 50000)""")

conn.commit()

conn.close()