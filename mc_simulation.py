# Monte Carlo Simulation with Python

import xlwings as xw
import numpy as np

def monte_carlo():

    wb = xw.Book.caller()

    inp = wb.sheets("Input")
    calc = wb.sheets("Calc")
    mc = wb.sheets("Monte_Carlo")

    sims = int(mc["D8"].value)

    cpi_exp = mc["D11"].value
    cpi_std = mc["D12"].value

    ppf_exp = mc["D15"].value
    ppf_std = mc["D16"].value

    cost_exp = mc["D19"].value
    cost_std = mc["D20"].value

    results = np.empty((sims, 2))

    for i in range(sims):
        calc["B3"].options(transpose=True).value = np.random.normal(
            cpi_exp, cpi_std, 10)
        inp["D25"].value = np.random.normal(ppf_exp, ppf_std)
        calc["H3"].options(transpose=True).value = - \
            np.random.normal(cost_exp, cost_std, 10)
        results[i] = inp["G24:G25"].value

    mc["I5"].options(transpose=True).value = np.mean(results, axis=0)
    mc["J5"].options(transpose=True).value = np.median(results, axis=0)
    mc["I7"].value = sum(results[:, 0] < 1) / sims


def restore():

    wb = xw.Book.caller()
    
    calc = wb.sheets[1]

    calc["B3:B12"].formula = "=Input!$D$20"
    calc["H3:H12"].formula = "=-Input!$D$40"