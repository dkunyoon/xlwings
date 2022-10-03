import xlwings as xw
import numpy as np

wb = xw.Book("CaseStudy_1.xlsx")
wb
# <Book [CaseStudy_1.xlsx]>

ws = wb.sheets("Sheet1")
ws
# <Sheet [CaseStudy_1.xlsx]Sheet1>

ws["B1"].name = "Saving"
ws["B2"].name = "Interest_Rate"
ws["B3"].name = "Periods"
ws["B4"].name = "Future_Value"

def fut_value(rate, nper, pv):
    return pv * ((1 + rate) ** nper)

fv = fut_value(rate=ws["Interest_Rate"].value, nper=ws["Periods"].value,
               pv=ws["Saving"].value)
fv
# 121.55062500000003

ws["Future_Value"].value = fv

wb.close()







