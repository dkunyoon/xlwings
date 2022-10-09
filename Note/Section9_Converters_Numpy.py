# =================================================================
# (Default) Converters and Settings
# =================================================================

import datetime as dt
import pandas as pd
import numpy as np
import xlwings as xw

wb = xw.Book()
sheet = wb.sheets[0]
sheet
# <Sheet [통합 문서1]Sheet1>

sheet.range("A1").value = [[4.32, "Dog"],
                           [dt.datetime(2020, 3, 30), None]]

sheet.range("A1:B2").value
# [[4.32, 'Dog'], [datetime.datetime(2020, 3, 30, 0, 0), None]]

sheet.range("A1").expand().value
# [[4.32, 'Dog'], [datetime.datetime(2020, 3, 30, 0, 0), None]]

sheet.range("A1").options(expand="table").value
# [[4.32, 'Dog'], [datetime.datetime(2020, 3, 30, 0, 0), None]]

cells = sheet.range("A1:B2")

cells.options(transpose=True).value
# [[4.32, datetime.datetime(2020, 3, 30, 0, 0)], ['Dog', None]]

cells.options(numbers=int).value
# [[4, 'Dog'], [datetime.datetime(2020, 3, 30, 0, 0), None]]

cells.options(numbers=lambda x: round(x, 1)).value
# [[4.3, 'Dog'], [datetime.datetime(2020, 3, 30, 0, 0), None]]

cells.options(empty="No Data.").value
# [[4.32, 'Dog'], [datetime.datetime(2020, 3, 30, 0, 0), 'No Data.']]


# =================================================================
# The Numpy Converter
# =================================================================

my_array = np.random.normal(10, 2, (3, 4))
my_array
# array([[ 8.19184433, 10.60360001, 10.07971632,  5.31556643],
#        [ 9.1162136 ,  8.48579143,  7.2544749 ,  8.45792332],
#        [11.27209094,  9.69337611,  8.35061143, 11.21907133]])

sheet.range("A5").value = my_array

sheet.range("A5").options(expand="table").value
# [[8.191844327042382, 10.603600007546126, 10.079716324420447, 5.31556642808807],
#  [9.116213597393665, 8.485791432207222, 7.25447490015803, 8.457923321255919],
#  [11.272090939976652, 9.693376106395133, 8.350611429420892, 11.219071325657179]]

sheet.range("A5").options(convert=np.array, expand="table").value
# array([[ 8.19184433, 10.60360001, 10.07971632,  5.31556643],
#        [ 9.1162136 ,  8.48579143,  7.2544749 ,  8.45792332],
#        [11.27209094,  9.69337611,  8.35061143, 11.21907133]])

sheet.range("A5").options(np.array, expand="table",
                          transpose=True, numbers=int).value
# array([[ 8,  9, 11],
#        [11,  8, 10],
#        [10,  7,  8],
#        [ 5,  8, 11]])


# =================================================================
# The Dictionary Converter
# =================================================================

sheet.range("A9").value = [["a", 1], ["b", 2], ["c", 3]]

sheet.range("A9").options(expand="table").value  # nested list
# [['a', 1.0], ['b', 2.0], ['c', 3.0]]

sheet.range("A9").options(dict, expand="table").value  # dictionary
# {'a': 1.0, 'b': 2.0, 'c': 3.0}

my_dic = sheet.range("A9").options(dict, expand="table", numbers=int).value
my_dic
# {'a': 1, 'b': 2, 'c': 3}
type(my_dic)

my_dic2 = {"d": 4, "e": 5, "f": 6}
my_dic2

sheet.range("A13").value = my_dic2
sheet.range("A13").options(transpose=True).value = my_dic2

sheet.range("A5:B7").options(dict).value
# {8.191844327042382: 10.603600007546126,
#  9.116213597393665: 8.485791432207222,
#  11.272090939976652: 9.693376106395133}
