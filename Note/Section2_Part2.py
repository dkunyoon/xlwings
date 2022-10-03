# =======================================================
# Writing Excel Functions with Python
# =======================================================

import xlwings as xw

wb = xw.Book()
wb
# <Book [통합 문서1]>

sheet = wb.sheets[0]
sheet
# <Sheet [통합 문서1]Sheet1>

sheet.range("A3").formula = "=SUM(A1:A2)"
sheet.range("A3").value

sheet.range("A3").formula
# '=SUM(A1:A2)'

sheet.range("B1").formula = "=A1"
sheet.range("B1").formula
# '=A1'


# =======================================================
# Range Shortcuts
# =======================================================

wb = xw.Book()
wb
# <Book [통합 문서2]>

sheet = wb.sheets[0]
sheet
# <Sheet [통합 문서2]Sheet1>

sheet.range("A1").value = 100
sheet.range("A1").value
# 100.0
sheet["B1"].value = 100
sheet["B1"].value
# 100.0

sheet["A3:C3"].value = [10,20,30]
sheet["A3:C3"].value
# [10.0, 20.0, 30.0]

wb.close()










