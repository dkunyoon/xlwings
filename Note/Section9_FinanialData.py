# =============================================================================
# Financial Data - Advanced Analysis Techiniques
# Importing Financial Data from Excel
# =============================================================================

import pandas as pd

sp500 = pd.read_excel("SP500.xls")
sp500
#             Date         Open  ...    Adj Close      Volume
# 0     1970-12-31    92.269997  ...    92.150002    13390000
# 1     1971-01-04    92.150002  ...    91.150002    10010000
# 2     1971-01-05    91.150002  ...    91.800003    12600000

sp500.info()
# <class 'pandas.core.frame.DataFrame'>
# RangeIndex: 12107 entries, 0 to 12106
# Data columns (total 7 columns):
#  #   Column     Non-Null Count  Dtype
# ---  ------     --------------  -----
#  0   Date       12107 non-null  datetime64[ns]
# ...

sp500 = pd.read_excel("SP500.xls", parse_dates=["Date"], index_col="Date")
sp500
#                    Open         High  ...    Adj Close      Volume
# Date                                  ...
# 1970-12-31    92.269997    92.790001  ...    92.150002    13390000
# 1971-01-04    92.150002    92.190002  ...    91.150002    10010000

sp500.info()
# <class 'pandas.core.frame.DataFrame'>
# # DatetimeIndex: 12107 entries, 1970-12-31 to 2018-12-28
# # Data columns (total 6 columns):
# #  #   Column     Non-Null Count  Dtype
# # ---  ------     --------------  -----
# #  0   Open       12107 non-null  float64

pd.read_excel("SP500.xls", sheet_name="SP500")

sp500 = pd.read_excel("SP500.xls", parse_dates=["Date"], index_col="Date", usecols="A:E")
sp500.head()
#                    Open         High          Low        Close
# Date
# 1970-12-31    92.269997    92.790001    91.360001    92.150002
# 1971-01-04    92.150002    92.190002    90.639999    91.150002
# 1971-01-05    91.150002    92.279999    90.690002    91.800003
sp500.tail()
#                    Open         High          Low        Close
# Date
# 2018-12-21  2465.379883  2504.409912  2408.550049  2416.620117
# 2018-12-24  2400.560059  2410.340088  2351.100098  2351.100098
# 2018-12-26  2363.120117  2467.760010  2346.580078  2467.699951

sp500.info()
# <class 'pandas.core.frame.DataFrame'>
# DatetimeIndex: 12107 entries, 1970-12-31 to 2018-12-28
# Data columns (total 4 columns):
#  #   Column  Non-Null Count  Dtype
# ---  ------  --------------  -----
#  0   Open    12107 non-null  float64
#  1   High    12107 non-null  float64
#  2   Low     12107 non-null  float64
#  3   Close   12107 non-null  float64

sp500.to_csv("SP500.csv")
sp500.to_excel("SP500_red.xlsx")

























