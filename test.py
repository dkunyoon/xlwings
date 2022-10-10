import pandas as pd
import xlwings as xw

wb = xw.Book("hello.xlsm")
ws = wb.sheets[0]

df = ws.range("B18:E22").options(pd.DataFrame, expand="table", index=False, header=False).value
df
df.sum().sum()
df.add(1)
