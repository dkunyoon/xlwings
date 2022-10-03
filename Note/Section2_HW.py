''' Homework '''

# Import the required Library

import pandas as pd
import numpy as np
import xlwings as xw


# Open a new workbook and create a connection to that workbook with xlwings.
# Create a sheet object.

wb = xw.Book()
ws = wb.sheets[0]


# Type two numbers of your choice into the cells A1 and B1 (no Python Code required)
# Create named ranges for both cells and read the values into Python with xlwings.

ws["A1"].name = "Cell_A1"
ws["B1"].name = "Cell_B1"


# Move the cells in Excel and change their values (no Python Code required)
# Read the values into Python, calculate the sum in Python,
# and write the sum into any other cell with xlwings.

ws["Cell_A1"].value  # 120.0
ws["Cell_B1"].value  # 130.0

ws["Cell_A1"].value + ws["Cell_B1"].value  # 250.0


# Open the Excel file "Test2.xlsx" and change the value in cell C1 from 2 to "two"

wb2 = xw.Book("Test2.xlsx")
ws2 = wb2.sheets("Tabelle1")

ws2["C1"].value = "two"


# Try to create an xlwings connection to any excel workbook of your choice
# on your computer and read and write values

ws2.range("A2").value = "Hi"
ws2.range("A3:F3").value = [1,2,3,4,5,6]

