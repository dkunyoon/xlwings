# =============================================================================
# Importing from Excel Files with pd.read_excel()
# =============================================================================

# First Steps

import pandas as pd

sales = pd.read_excel("sales.xls")
sales
#   Unnamed: 0      City Country  Sales  Bonus
# 0       Mike  New York     USA     25   2.50
# 1        Jim    Boston     USA     43   4.30
# 2     Steven    London      UK     76   7.60
sales.info()

pd.read_excel("sales.xls", index_col=0, header=0)

#             City Country  Sales  Bonus
# Mike    New York     USA     25   2.50
# Jim       Boston     USA     43   4.30
# Steven    London      UK     76   7.60

pd.read_excel("sales.xls", index_col=0, header=0,
              names=["Name", "Loc_City", "Loc_Country", "Revenue", "Add_Comp"])

#         Loc_City Loc_Country  Revenue  Add_Comp
# Name
# Mike    New York         USA       25      2.50
# Jim       Boston         USA       43      4.30
# Steven    London          UK       76      7.60

pd.read_excel("sales.xls", index_col=0, header=0, usecols="A:C")

#             City Country
# Mike    New York     USA
# Jim       Boston     USA
# Steven    London      UK

pd.read_excel("sales.xls", index_col=0, header=0, usecols="C:E")
# C:E 에서 0번째인 C가 Index가 되고 D:E는 Columns로 인식

#          Sales  Bonus
# Country
# USA         25   2.50
# USA         43   4.30
# UK          76   7.60

pd.read_excel("sales.xls", index_col=0, header=0, usecols="A, C:E")

#        Country  Sales  Bonus
# Mike       USA     25   2.50
# Jim        USA     43   4.30
# Steven      UK     76   7.60

pd.read_excel("sales.xls", index_col=0, header=0, usecols=":C")

#             City Country
# Mike    New York     USA
# Jim       Boston     USA
# Steven    London      UK

pd.read_excel("sales.xls", index_col=0, header=0, usecols="C:") # does not work

pd.read_excel("sales.xls", index_col=0, header=0, usecols=[0,3,4])

#         Sales  Bonus
# Mike       25   2.50
# Jim        43   4.30
# Steven     76   7.60

pd.read_excel("sales.xls", index_col=0, header=0, usecols=["City","Sales"])

#           Sales
# City
# New York     25
# Boston       43
# London       76

# =============================================================================
# Customizing import with pd.read_excel()
# =============================================================================

# ↓ by default imports the first sheet
pd.read_excel("summer_raw.xls")
#   Unnamed: 0      City Country  Sales  Bonus
# 0       Mike  New York     USA     25    2.5
# 1        Jim    Boston     USA     43    4.3
# 2     Steven    London      UK     76    7.6
# 3        Joe    Madrid   Spain     12    1.8
# 4        Tom     Paris  France     89   13.4

# ↓ specify a sheet name to import
summer_raw = pd.read_excel("summer_raw.xls", sheet_name="summer")
#        Unnamed: 0  Unnamed: 1  ...                 Unnamed: 10 Unnamed: 11
# 0             NaN         NaN  ...                         NaN         NaN
# 1             NaN         NaN  ...                       Event       Medal
# 2             NaN         NaN  ...              100M Freestyle  Gold Medal
#            ...         ...  ...                         ...         ...
# 31170         NaN         NaN  ...                    Wg 96 KG      Bronze
# 31171         NaN         NaN  ...                    Wg 96 KG      Bronze

summer_raw = pd.read_excel("summer_raw.xls", sheet_name="summer", skiprows=2)
#        Unnamed: 0  Unnamed: 1  ...                       Event       Medal
# 0             NaN         NaN  ...              100M Freestyle  Gold Medal
# 1             NaN         NaN  ...              100M Freestyle      Silver
# 2             NaN         NaN  ...  100M Freestyle For Sailors      Bronze

pd.read_excel("summer_raw.xls", sheet_name="summer", skiprows=2, usecols="C:L")
#        Unnamed: 2  Year    City  ... Gender                       Event       Medal
# 0               0  1896  Athens  ...    Men              100M Freestyle  Gold Medal
# 1               1  1896  Athens  ...    Men              100M Freestyle      Silver
# 2               2  1896  Athens  ...    Men  100M Freestyle For Sailors      Bronze

pd.read_excel("summer_raw.xls", sheet_name="summer", skiprows=2, usecols="D:L")
#        Year    City      Sport  ... Gender                       Event       Medal
# 0      1896  Athens   Aquatics  ...    Men              100M Freestyle  Gold Medal
# 1      1896  Athens   Aquatics  ...    Men              100M Freestyle      Silver
# 2      1896  Athens   Aquatics  ...    Men  100M Freestyle For Sailors      Bronze

summer = pd.read_excel("summer_raw.xls", sheet_name="summer", skiprows=2, usecols="D:L")
summer.tail()
summer.info()

summer.to_csv("summer_imp.csv", index=False)
summer.to_excel("summer_imp.xlsx")

pd.read_csv("summer_imp.csv")





