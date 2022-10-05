'''
[Project : Creating a Stock Dashboard]
'''
# ============================================================
# Analyzing Stocks with Python and xlwings
# ============================================================

import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pandas_datareader import data
from statsmodels.formula.api import ols

plt.style.use("seaborn")

wb = xw.Book("../Xlwings/Stocks.xlsx")

db_s = wb.sheets[0]
prices_s = wb.sheets[1]

symbol = db_s.range("C7").value
start = db_s.range("F7").value
end = db_s.range("I7").value
benchmark = db_s.range("K20").value
freq = db_s.range("K39").value

print(symbol, start, end, benchmark, freq, sep="|")

df = data.DataReader(name=[symbol, benchmark], data_source="yahoo", start=start, end=end).Close
df.rename(columns={benchmark: benchmark.replace("^", "")}, inplace=True)
df

benchmark = benchmark.replace("^", "")

chart = plt.figure(figsize=(19, 8))
df[symbol].plot(fontsize=15)
plt.title(symbol, fontsize=20)
plt.xlabel("Date", fontsize=15)
plt.ylabel("Stock Price", fontsize=15)
plt.show()

# sizing picture is trial and error process
db_s.pictures.add(chart, name="Chart", update=True, left=db_s.range("C8").left, top=db_s.range("C8").top, scale=0.9)

first = df.iloc[0, 0]
high = df.iloc[:, 0].max()
low = df.iloc[:, 0].min()
last = df.iloc[-1, 0]

total_change = last / first - 1
total_change

db_s.range("H12").options(transpose=True).value = [first, high, low, last, total_change]

# normalize data
norm = df.div(df.iloc[0]).mul(100)
norm

chart2 = plt.figure(figsize=(19, 8))
norm[symbol].plot(fontsize=15)
norm[benchmark].plot(fontsize=15)
plt.title(symbol + " vs. " + benchmark, fontsize=20)
plt.xlabel("Date", fontsize=15)
plt.ylabel("Normalized Price (Base 100)", fontsize=15)
plt.legend(fontsize=20)
plt.show()

db_s.pictures.add(chart2, name="Chart2", update=True,
                  left=db_s.range("C21").left, top=db_s.range("C21").top, scale=0.9)

ret = df.resample(freq).last().dropna().pct_change().dropna()
ret

chart3 = plt.figure(figsize=(12.5, 10))
sns.regplot(data=ret, x=benchmark, y=symbol)
plt.title(symbol + " vs. " + benchmark, fontsize=20)
plt.xlabel(benchmark + " Returns", fontsize=15)
plt.ylabel(symbol + "Returns", fontsize=15)
plt.show()

db_s.pictures.add(chart3, name="Chart3", update=True,
                  left=db_s.range("C40").left, top=db_s.range("C40").top, scale=0.45)

model = ols(symbol + "~" + benchmark, data=ret)
results = model.fit()

results.summary()

"""
                            OLS Regression Results                            
==============================================================================
Dep. Variable:                   MSFT   R-squared:                       0.787
Model:                            OLS   Adj. R-squared:                  0.785
Method:                 Least Squares   F-statistic:                     453.6
Date:                Thu, 06 Oct 2022   Prob (F-statistic):           4.41e-43
Time:                        00:12:52   Log-Likelihood:                 341.17
No. Observations:                 125   AIC:                            -678.3
Df Residuals:                     123   BIC:                            -672.7
Df Model:                           1                                         
Covariance Type:            nonrobust                                         
==============================================================================
                 coef    std err          t      P>|t|      [0.025      0.975]
------------------------------------------------------------------------------
Intercept      0.0029      0.001      2.060      0.041       0.000       0.006
DJI            0.9737      0.046     21.297      0.000       0.883       1.064
==============================================================================
Omnibus:                        1.334   Durbin-Watson:                   1.991
Prob(Omnibus):                  0.513   Jarque-Bera (JB):                0.897
Skew:                           0.169   Prob(JB):                        0.639
Kurtosis:                       3.240   Cond. No.                         32.1
==============================================================================
Notes:
[1] Standard Errors assume that the covariance matrix of the errors is correctly specified.
"""

obs = len(ret)
corr_coef = ret.corr().iloc[0, 1]
beta = results.params[1]
r_sq = results.rsquared
t_stat = results.tvalues[1]
p_value = results.pvalues[1]
conf_left = results.conf_int().iloc[1, 0]
conf_right = results.conf_int().iloc[1, 1]
interc = results.params[0]

reg_list = [obs, corr_coef, beta, r_sq, t_stat, p_value, conf_left, conf_right, interc]

db_s.range("K41").options(transpose=True).value = reg_list

prices_s.range("A1").expand().clear_contents()
prices_s.range("A1").value = df
