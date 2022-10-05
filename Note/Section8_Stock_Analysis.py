'''
[Project : Creating a Stock Dashboard]
'''
# ============================================================
# Analyzing Stocks with Python (Part 1)
# ============================================================

import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pandas_datareader import data
# need installation: conda install pandas-datareader
from statsmodels.formula.api import ols
plt.style.use("seaborn")

symbol = "MSFT"
start = "2020-01-01"
end = "2020-06-30"
benchmark = "^DJI"

data.DataReader(name=[symbol, benchmark],
                data_source="yahoo", start=start, end=end)

df = data.DataReader(name=[symbol, benchmark],
                     data_source="yahoo", start=start, end=end).Close
df.rename(columns={benchmark: benchmark.replace("^", "")}, inplace=True)
df

plt.figure(figsize=(12, 8))
df[symbol].plot(fontsize=12)
plt.show()

first = df.iloc[0, 0]
first

high = df.iloc[:, 0].max()
high

low = df.iloc[:, 0].min()
low

last = df.iloc[-1, 0]
last

total_chg = (last/first)-1
total_chg


# ============================================================
# Analyzing Stocks with Python (Part 2)
# ============================================================

df

benchmark = benchmark.replace("^", "")

plt.figure(figsize=(12, 8))
df[symbol].plot(fontsize=12)
df[benchmark].plot(fontsize=12)
plt.show()

# need to Normalize

norm = df.div(df.iloc[0]).mul(100)
norm

plt.figure(figsize=(12, 8))
norm[symbol].plot(fontsize=12)
norm[benchmark].plot(fontsize=12)
plt.legend(fontsize=15)
plt.show()

df.pct_change().dropna()  # calculate daily returns

# Symbols	    MSFT	    DJI
# Date
# 2020-01-02	 0.018516	 0.011576
# 2020-01-03	-0.012452	-0.008103
# 2020-01-06	 0.002585	 0.002392
# 2020-01-07	-0.009118	-0.004170
# ...	...	...

df.resample("M").last().dropna()  # determine frequency

# Symbols	    MSFT    	DJI
# Date
# 2019-12-31	157.699997	28538.439453
# 2020-01-31	170.229996	28256.029297
# 2020-02-29	162.009995	25409.359375
# 2020-03-31	157.710007	21917.160156
# ...	...	...

df.resample("M").last().pct_change().dropna()  # monthly pct_change

# Symbols	    MSFT    	DJI
# Date
# 2020-01-31	 0.079455	-0.009896
# 2020-02-29	-0.048288	-0.100746
# 2020-03-31	-0.026541	-0.137438
# ...	...	...


df.resample("D").last()  # determine frequency = daily

# Symbols	    MSFT	    DJI
# Date
# 2019-12-31	157.699997	28538.439453
# 2020-01-01	NaN	NaN
# 2020-01-02	160.619995	28868.800781
# 2020-01-03	158.619995	28634.880859
# 2020-01-04	NaN	NaN
# ...	...	...
# 2020-06-26	196.330002	25015.550781
# 2020-06-27	NaN	NaN
# 2020-06-28	NaN	NaN
# 2020-06-29	198.440002	25595.800781
# 2020-06-30	203.509995	25812.880859

ret = df.resample("D").last().dropna().pct_change().dropna()  # daily returns
ret

# Symbols	     MSFT	     DJI
# Date
# 2020-01-02	 0.018516	 0.011576
# 2020-01-03	-0.012452	-0.008103
# 2020-01-06	 0.002585	 0.002392
# ...	...	...
# 2020-06-26	-0.020016	-0.028356
# 2020-06-29	 0.010747	 0.023196
# 2020-06-30	 0.025549	 0.008481

ret.corr()  # .iloc[0,1]

# Symbols	MSFT	DJI
# Symbols
# MSFT	    1.00000	0.88694
# DJI	    0.88694	1.00000


# ============================================================
# Analyzing Stocks with Python (Part 3)
# ============================================================

sns.regplot(data=ret, x=benchmark, y=symbol)  # regression plot
plt.show()

ols("MSFT~DJI", data=ret)
symbol + "~" + benchmark

model = ols(symbol + "~" + benchmark, data=ret)
results = model.fit()

print(results.summary())

#                             OLS Regression Results                            
# ==============================================================================
# Dep. Variable:                   MSFT   R-squared:                       0.787
# Model:                            OLS   Adj. R-squared:                  0.785
# Method:                 Least Squares   F-statistic:                     453.6
# Date:                Wed, 05 Oct 2022   Prob (F-statistic):           4.41e-43
# Time:                        22:08:19   Log-Likelihood:                 341.17
# No. Observations:                 125   AIC:                            -678.3
# Df Residuals:                     123   BIC:                            -672.7
# Df Model:                           1                                         
# Covariance Type:            nonrobust                                         
# ==============================================================================
#                  coef    std err          t      P>|t|      [0.025      0.975]
# ------------------------------------------------------------------------------
# Intercept      0.0029      0.001      2.060      0.041       0.000       0.006
# DJI            0.9737      0.046     21.297      0.000       0.883       1.064
# ==============================================================================
# Omnibus:                        1.334   Durbin-Watson:                   1.991
# Prob(Omnibus):                  0.513   Jarque-Bera (JB):                0.897
# Skew:                           0.169   Prob(JB):                        0.639
# Kurtosis:                       3.240   Cond. No.                         32.1
# ==============================================================================

# Notes:
# [1] Standard Errors assume that the covariance matrix of the errors is correctly specified.

results.params
# Intercept    0.002933
# DJI          0.973701

beta = results.params[1]
beta
# 0.9737012837183794

results.rsquared
# 0.7866626673541365

results.tvalues
# Intercept     2.060127
# DJI          21.296753

results.conf_int() # confidence interval
#           0       	1
# Intercept	0.000115	0.005752
# DJI	    0.883200	1.064202









