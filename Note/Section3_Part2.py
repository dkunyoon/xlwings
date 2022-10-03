# ==================================================
# Efficiency
# ==================================================

import xlwings as xw
import numpy as np

wb = xw.Book()
sheet = wb.sheets[0]

list(enumerate(sheet["A1:E50"]))

for i, cell in enumerate(sheet.range("A1:E50")):
    cell.value = i

sheet.range("A1:E50").clear()


# numpy approach
array = np.arange(5*50).reshape(50,5)
sheet.range("A1").value = array







