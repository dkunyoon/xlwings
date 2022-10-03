# ============================================================
# Excel Chart
# ============================================================

import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt

wb = xw.Book()
sheet = wb.sheets[0]

sheet.range("A1").options(transpose=True).value = [1, 3, 4, 6, 8, 10]
sheet.range("A1").expand().value

chart = sheet.charts.add()
chart

chart.set_source_data(sheet.range("A1").expand())

# chart.chart_type = "line"
chart.chart_type = "area"

xw.constants.chart_types
# ('3d_area',
#  '3d_area_stacked',
#  '3d_area_stacked_100',
#  '3d_bar_clustered',
# ...
#  'xy_scatter',
#  'xy_scatter_lines',
#  'xy_scatter_lines_no_markers',
#  'xy_scatter_smooth',
#  'xy_scatter_smooth_no_markers')

chart.left = sheet.range("D4").left
chart.top = sheet.range("D4").top

chart.height
# 211.0

chart.width
# 355.0

chart.height = 200
chart.width = 355