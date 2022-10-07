'''
HOA Trial Balance Converter

Description: A tool to help me convert HOA's TB PDF pages into Excel table,
so I can build a summary of monthly TB easily using vlookup in Excel.

After run the codes:
Pump-up window #1: Select the PDF FS file (Note, the FS file should only include the TB pages)
Pump-up window #2: Select the summary Excel file that I want the monthly table to be created

Output: A new tab called "08.22" (month.year) in the summary Excel file.

'''

import tkinter as tk
from tkinter import filedialog
from os import getcwd
from glob import glob
from tika import parser
import re
import openpyxl

'''
Use a tk window to obtain PDF file path
Input: none
Output: file path in string like this (on mac) "/Users/xxxx/Documents/Project/13F/13F_PDF_files/13flist2021q2.pdf
'''


def obtain_pdf_file_path():
    root = tk.Tk()
    root.withdraw()  # hides the root window

    # create window to select PDF file
    root.filename =  filedialog.askopenfilename(initialdir=getcwd(),\
                                                title = "Select file",\
                                                filetypes = (("PDF file","*.pdf"),\
                                                             ("all files","*.*")))
    return root.filename

def obtain_xlsx_file_path():
    root = tk.Tk()
    root.withdraw()  # hides the root window

    # create window to select PDF file
    root.filename =  filedialog.askopenfilename(initialdir=getcwd(),\
                                                title = "Select file",\
                                                filetypes = (("PDF file","*.xlsx"),\
                                                             ("all files","*.*")))
    return root.filename

def save_excel_file_path():
    root = tk.Tk()
    root.withdraw()  # hides the root window

    # create window to select PDF file
    root.filename = filedialog.asksaveasfilename(initialdir=getcwd(), \
                                                 title="Select file", \
                                                 filetypes=(("Excel files", "*.xlsx"), \
                                                            ("all files", "*.*")))
    return root.filename

'''
Read PDF into string and organize the string by line into a list
Input: file path
Output: a split list of strings by PDF line without empty lines
'''


def parsePDF(input_path):
    for input_file in glob(input_path):
        # make file into a string
        parserPDF = parser.from_file(input_file)
        # make string more readable separating by line (it is still a str)
        pdf = parserPDF['content']
        # split the str by line break (\n) and add each block of string in a list
        split = pdf.splitlines()
        # remove empty line
        split2 = [x for x in split if x.strip()]

        return split2

def cut_inrelevant_rows(el, table):

    if len(el) > 6 and el[6] == '-': # Check if the string contains "-" at index 6
        table.append(el)
        return table
    else:
        return None

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

def split_row(el, table):

    row = re.split('\s', el)
    new_row_1 = []

    count = 0
    for input in row:
        if count < len(row) - 1 and input == '-' and not is_number(row[count + 1][0]) and row[count + 1] != "-" \
                and row[count+1][0] != "(":
            pass
        elif re.search('\d', input) == None and input != "-" and row.index(input) > 1:
            if len(new_row_1) < 2:
                new_row_1.append(input)
            else:
                new_row_1[1] = new_row_1 [1] + ' ' + input
        else:
            new_row_1.append(input)
        count += 1

    table.append(new_row_1)



def workbook_insert_sheet(file_path, FS_file_name, pdf_data_list):

    # open the summary file (set up originally in the folder) and add new sheet called FS file's name
    wb = openpyxl.load_workbook(file_path)
    wb.create_sheet(FS_file_name)
    ws = wb[FS_file_name]

    # create heading names
    headings = ["Acct#", "Acct", "PTD Actual", "PTD Budget", "PTD Variance", "YTD Actual", "YTD Budget", "YTD Variance",
                "Annual Budget"]

    # insert heading names to the top of the new sheet
    ws.append(headings)

    # input PDF data into the new sheet
    for row in pdf_data_list:

        for i in range(len(row)):
            if row[i] == "-":
                row[i] = "0"
        ws.append(row)


    wb.save(file_path)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # Obtain PDF file path
    input_path = obtain_pdf_file_path()

    # Read and organize PDF file into a list
    split_list = parsePDF(input_path)

    # Create a list to input all valid rows
    valid_rows = []

    # Create a list to input row elements in sublist
    final_table = []

    for el in split_list:
        cut_inrelevant_rows(el, valid_rows)

    for el in valid_rows:
        split_row(el, final_table)

    # Obtain PDF file name (used to set up sheet name)
    file_name = input_path[-12:-7]

    # Obtain Summary Excel file path
    output_path = obtain_xlsx_file_path()

    # Input FS data into Summary Excel file in new sheet
    workbook_insert_sheet(output_path, file_name, final_table)

    print(final_table)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
