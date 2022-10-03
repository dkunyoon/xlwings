# ==================================================
# Reading and Writing Many Values
# ==================================================

# One-dimensional(1d) Structures

import xlwings as xw

wb = xw.Book()
wb
# <Book [통합 문서1]>

sheet = wb.sheets[0]
sheet
# <Sheet [통합 문서1]Sheet1>

sheet.range("A1").value = 123
sheet.range("A1").value

l = [1, 2, 3, 4]
sheet.range("A1").value = l  # writes a list horizontally
sheet.range("A1").value  # 1.0
sheet.range("B1").value  # 2.0

sheet.range("A1:D1").value  
# reading an excel row into a list
# [1.0, 2.0, 3.0, 4.0]

sheet.range("A2:D2").value = 765  # writes a scalar value multiple times

l2 = [5, 6, 7]

# writing a list vertically does not work
sheet.range("A3:A5").value = l2  # this does not work!
sheet["A3:A5"].value = l2  # this does not work!

# writing a scalar value multiple times vertically
sheet.range("A4:A6").value = 423
sheet["A4:A6"].value = 567

# reading an excel column into a list
sheet.range("A1:A6").value
sheet["A1:A6"].value
# [1.0, 5.0, 5.0, 567.0, 567.0, 567.0]


# ==================================================
# How to write lists vertically
# ==================================================

l2 = [5, 6, 7]

# this does not work
sheet.range("A7:A9").value = l2

# this works
sheet.range("A7").options(transpose=True).value = l2
sheet["B7"].options(transpose=True).value = l2


# ==================================================
# Reading Rows and Columns (1dim vs 2dim)
# ==================================================

sheet.range("A1:D1").value  # 1dim
# [1.0, 2.0, 3.0, 4.0]

sheet.range("A1:A9").value  # 1dim
# [1.0, 765.0, 5.0, 567.0, 567.0, 567.0, 5.0, 6.0, 7.0]

sheet.range("A1:A9").options(ndim=1).value  # 1dim
# [1.0, 765.0, 5.0, 567.0, 567.0, 567.0, 5.0, 6.0, 7.0]

sheet.range("A1:A9").options(ndim=2).value  # 2dim
# [[1.0], [765.0], [5.0], [567.0], [567.0], [567.0], [5.0], [6.0], [7.0]]
# nested list

sheet.range("A1:D1").options(ndim=1).value  # 1dim
# [1.0, 2.0, 3.0, 4.0]

sheet.range("A1:D1").options(ndim=2).value  # 2dim
# [[1.0, 2.0, 3.0, 4.0]]

sheet.range("A1:D6").options(ndim=2).value  # 2dim
# [[1.0, 2.0, 3.0, 4.0],
#  [765.0, 765.0, 765.0, 765.0],
#  [5.0, 6.0, 7.0, None],
#  [567.0, None, None, None],
#  [567.0, None, None, None],
#  [567.0, None, None, None]]


# ==================================================
# How to read two-dimensional (2d) Structures
# ==================================================

sheet.range("A4:A9").clear()

sheet.range("A1:C3").value
# [[1.0, 2.0, 3.0], [765.0, 765.0, 765.0], [5.0, 6.0, 7.0]]
# nested list

sheet["A1:E4"].value
# [[1.0, 2.0, 3.0, 4.0, None],
#  [765.0, 765.0, 765.0, 765.0, None],
#  [5.0, 6.0, 7.0, None, None],
#  [None, None, None, None, None]]

sheet["A1:C3"].options(transpose=True).value  # transposed
# [[1.0, 765.0, 5.0], [2.0, 765.0, 6.0], [3.0, 765.0, 7.0]]


# ==================================================
# Advanced Reading with expand
# ==================================================

sheet["A1:D1"].value
# [1.0, 2.0, 3.0, 4.0]

sheet.range("A1").expand("right").value
# [1.0, 2.0, 3.0, 4.0]

sheet["A1"].expand("down").value
# [1.0, 765.0, 5.0]

sheet["A1"].expand("table").value
# [[1.0, 2.0, 3.0, 4.0], [765.0, 765.0, 765.0, 765.0], [5.0, 6.0, 7.0, None]]

sheet["A1"].expand().value
# [[1.0, 2.0, 3.0, 4.0], [765.0, 765.0, 765.0, 765.0], [5.0, 6.0, 7.0, None]]

# only expand="down" >>> list
sheet.range("A1").options(expand="down").value

# ndim=2 and expand="down" >>> nested list
sheet.range("A1").options(ndim=2, expand="down").value


# ==================================================
# How to write two-dimensional (2d) Structures
# ==================================================

sheet["A4"].value = [[1,2],[3,4]]

sheet.range("A7").options(transpose=True).value = [5,6,7]

sheet["A7"].value = [[5],[6],[7]]


# ==================================================
# Range Indexing & Slicing
# ==================================================

sheet.range("A1:D1").value
# [1.0, 2.0, 3.0, 4.0]

sheet["A1:D1"][0].value
# 1.0

sheet["A1:D1"][0:2].value
# [1.0, 2.0]

sheet["A1:D1"][-2:].value
# [3.0, 4.0]

sheet["A1:D3"].value
# [[1.0, 2.0, 3.0, 4.0], [765.0, 765.0, 765.0, 765.0], [5.0, 6.0, 7.0, None]]

sheet["A1:D3"][9].value
# 6.0

sheet["A1:D3"][1, 1].value
# 765.0

sheet.range("A1:D3")[:, 3].value # fourth column
# [4.0, 765.0, None]

sheet.range("A1").expand()[:2, -2:].value # first 2 rows, last 2 columns
# [[3.0, 4.0], [765.0, 765.0]]

sheet.range("A1").expand()[:, :2].value  # first 2 columns
# [[1.0, 2.0], [765.0, 765.0], [5.0, 6.0], [1.0, 2.0], [3.0, 4.0]]


wb.close()











