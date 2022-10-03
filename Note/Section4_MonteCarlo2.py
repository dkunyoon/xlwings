# ============================================================
# Running a more realistic/advanced Monte Carlo Simulation
# ============================================================

import xlwings as xw
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

wb = xw.Book("real_estate.xlsx")
inp = wb.sheets[0]
calc = wb.sheets[1]

cpi_exp = 0.02
cpi_std = 0.01

ppf_exp = 23
ppf_std = 3

cost_exp = 250000
cost_std = 50000

sims = 100

results = np.empty((sims, 2))

for i in range(sims):
    calc["B3"].options(transpose=True).value = np.random.normal(cpi_exp, cpi_std, 10)
    inp["D25"].value = np.random.normal(ppf_exp, ppf_std)
    calc["H3"].options(transpose=True).value = -np.random.normal(cost_exp, cost_std, 10)
    results[i] = inp["G24:G25"].value

results
# array([[1.66871047, 0.05590604],
#        [1.44750323, 0.0404696 ],
#        [1.50366201, 0.04440902],
#        ...,
#        [1.90613799, 0.07050578],
#        [1.87476549, 0.06876021],
#        [1.39950823, 0.03690093]])

plt.figure(figsize=(12,8))
plt.hist(results[:,0], bins=100)
plt.title("Investment Multiple", fontsize=15)
plt.show()

plt.figure(figsize=(12,8))
plt.hist(results[:,1], bins=100)
plt.title("IRR", fontsize=15)
plt.show()

pd.DataFrame(results, columns=["Invest_Multiple", "IRR"]).describe()
#        Invest_Multiple         IRR
# count       10000.0000  10000.0000
# mean          1.748337    0.058397
# std           0.382612    0.025408
# min           0.887576   -0.013182
# 25%           1.469314    0.041848
# 50%           1.709936    0.058565
# 75%           2.032193    0.077499
# max           2.568876    0.103961


# ============================================================
# Final Considerations
# ============================================================

# B3:B12 formula fix 안됨.
calc["B3:B12"].formula = "=Input!D20"

# fix reference to paste consistent formula
calc["B3:B12"].formula = "=Input!$D$20"
calc["H3:H12"].formula = "=-Input!$D$40"






