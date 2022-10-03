import pandas as pd
import numpy as np
import xlwings as xw
import matplotlib.pyplot as plt

wb = xw.Book("real_estate.xlsx")
inp = wb.sheets("Input")
calc = wb.sheets("Calc")

inp["D20"].name = "cpi"
inp["D25"].name = "ppf"
inp["D40"].name = "cost"
inp["G24:G25"].name = "performance"
inp["performance"].value
# [2.08259459669125, 0.08066461980342868]

# ============================================================
# Probability of Distribution CPI (normal)
# ============================================================

cpi_exp = 0.02
cpi_std = 0.01
sims = 10000

cpi_pd = np.random.normal(cpi_exp, cpi_std, sims)
cpi_pd
# array([0.03090248, 0.01853232, 0.03840686, ..., 0.01615315, 0.03010968,
#        0.01542572])

plt.hist(cpi_pd, bins=100)
plt.show()

# ============================================================
# Probability of Distribution Purchase Price Factor (normal)
# ============================================================

ppf_exp = 23
ppf_std = 3

ppf_pd = np.random.normal(ppf_exp, ppf_std, sims)
ppf_pd
# array([18.90575687, 19.46975441, 21.78921515, ..., 28.51028419,
#        24.01791172, 18.32934088])

plt.hist(ppf_pd, bins=100)
plt.show()

# ============================================================
# Probability Distribution Costs (normal)
# ============================================================

cost_exp = 250000
cost_std = 50000

# ============================================================
# Final Assumption: No Correlation between Inputs (can be changed)
# ============================================================

# number of simulations
sims = 10000

results = np.empty((sims, 2))  # shape: (sims, 2) > sims=10; 10rows, 2columns

for i in range(sims):
    inp["cpi"].value = np.random.normal(cpi_exp, cpi_std)
    inp["ppf"].value = np.random.normal(ppf_exp, ppf_std)
    inp["cost"].value = np.random.normal(cost_exp, cost_std)
    results[i] = inp["performance"].value

print(results)
# [[2.46648637 0.09990355]
#  [1.6307286  0.05331437]
#  [1.28016005 0.02672541]
#  ...
#  [2.14276319 0.08374907]
#  [1.66713001 0.05576042]
#  [2.25527657 0.08881433]]

plt.figure(figsize=(12, 8))
plt.hist(results[:, 0], bins=100)
plt.title("Investment Multiple", fontsize=15)
plt.show()

plt.figure(figsize=(12, 8))
plt.hist(results[:, 1], bins=100)
plt.title("IRR", fontsize=15)
plt.show()

results.mean(axis=0)
# array([1.72152754, 0.05586415])
np.median(results, axis=0)
# array([1.69801017, 0.05789403])

sum(results[:, 0] < 1) / sims
# 0.0374
