import xlwings as xw
import pandas as pd
import numpy as np

# numpy array

np.random.seed(123)
data1 = np.random.randint(1, 100, 10000).reshape(100,100)
data1

# array([[67, 93, 99, ..., 88, 15, 84],
#        [71, 13, 55, ..., 46, 59, 60],
#        [24, 26, 28, ..., 90, 75,  1],
#        ...,
#        [17, 13, 84, ..., 81, 44,  3],
#        [44, 71, 91, ..., 18, 67, 44],
#        [36, 61, 11, ..., 58, 58, 90]])


data1.shape
# (100, 100)

xw.view(data1)

np.random.seed(123)
data2 = np.random.randint(1, 100, 100).reshape(10,10)
data2

# array([[67, 93, 99, 18, 84, 58, 87, 98, 97, 48],
#        [74, 33, 47, 97, 26, 84, 79, 37, 97, 81],
#        [69, 50, 56, 68,  3, 85, 40, 67, 85, 48],
#        [62, 49,  8, 93, 53, 98, 86, 95, 28, 35],
#        [98, 77, 41,  4, 70, 65, 76, 35, 59, 11],
#        [23, 78, 19, 16, 28, 31, 53, 71, 27, 81],
#        [ 7, 15, 76, 55, 72,  2, 44, 59, 56, 26],
#        [51, 85, 57, 50, 13, 19, 82,  2, 52, 45],
#        [49, 57, 92, 50, 87,  4, 68, 12, 22, 90],
#        [99,  4, 12,  4, 95,  7, 10, 88, 15, 84]])

xw.view(data2)
xw.view(data2, sheet=xw.sheets.active)

# =======================================================
# Pandas DataFrame
# =======================================================

df = pd.read_csv("titanic.csv")
df

pd.options.display.max_rows = 900

xw.view(df)


# =======================================================
# Connecting to a Book
# =======================================================

import xlwings as xw

# 1. Open a new Book

xw.Book() # creates an object
# <Book [통합 문서1]>

xw.books.active
# <Book [통합 문서1]>

xw.books.active.close()

wb = xw.Book()
wb.close()


# 2. Connect to an open, unsaved Book

wb = xw.Book("통합 문서1")
wb.close()


# 3. Connect to a saved Book

path = "G:\My Drive\Python\Python_for_Excel_xlwings_Udemy\Test2.xlsx"
wb = xw.Book(path)

# or

wb = xw.Book("Test2.xlsx")
wb
# <Book [test2.xlsx]>


# =======================================================
# Reading and Writing single Values
# =======================================================

import xlwings as xw

wb = xw.Book()
sheet = wb.sheets[0]

sheet
# <Sheet [통합 문서1]Sheet1>

xw.books.active
# <Book [통합 문서1]>

xw.sheets.active
# <Sheet [통합 문서1]Sheet1>


# =======================================================
# Writing a single Value
# =======================================================

sheet.range("A1").value = "Hello World"


# =======================================================
# Reading a single Value
# =======================================================

sheet.range("A1").value
# 'Hello World'

sheet.range("A1")
# <Range [통합 문서1]Sheet1!$A$1>


'''Index numbers in Tuple approach'''

sheet.range((1,1)).value
# 'Hello World'

sheet.range((1,2)).value = 123
sheet.range("B1").value
# 123.0

int(sheet.range("B1").value)
# 123

'''datetime objects'''

import datetime as dt

date = dt.datetime(2020, 3, 20, 15, 30, 45)
date
# datetime.datetime(2020, 3, 20, 15, 30, 45)

sheet.range("A2").value = date
# 2020-03-20  3:30:45 PM

sheet.range("A2").value
# datetime.datetime(2020, 3, 20, 15, 30, 45)

sheet.range("A1").clear_contents()


# =======================================================
# Assigning a Name (named ranges)
# =======================================================

sheet.range("A1").value = 1000
sheet.range("A1").value
# 1000.0
sheet.range("A1").name = "Number"

### after inserting a column ###
sheet.range("Number").value
# 1000.0
sheet.range("Number").name
# <Name 'Number': =Sheet1!$B$1>

wb.close()















