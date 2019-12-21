import tkinter as tk
from tkinter import Label, Button, Frame, Tk, Entry

from modules.excel_db import ExcelToDb

class GUI:
    def __init__(self, master):
        self.master = master
        self.title = Label(text="Excel to Database converter")
        self.file_name_label = Label(text="Enter Excel file name")
        self.file_entry = Entry()
        self.db_name_label = Label(text="Enter Database name")
        self.db_entry = Entry()
        self.button = Button(text="Start conversion", command=self.start)
        self.packWidgets()

    def packWidgets(self):
        self.title.grid(row=0, column=0, columnspan=2)
        self.file_name_label.grid(row=1, column=0)
        self.file_entry.grid(row=1, column=1)
        self.db_name_label.grid(row=2, column=0)
        self.db_entry.grid(row=2, column=1)
        self.button.grid(row=3, column=0, columnspan=2)
    
    def start(self):
        file_name = self.file_entry.get()
        db_name = self.db_entry.get()
        ex_to_db = ExcelToDb(file_name, db_name)

if __name__ == "__main__":
    root = Tk()
    root.title("Excel to DB")

    frame = Frame(root)

    root.geometry("300x100")
    app = GUI(frame)

    root.mainloop()
