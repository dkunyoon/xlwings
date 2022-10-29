# =============================================================================
# Calculate FV and PV for many Cashflows
# =============================================================================

# Future Values

# Today you have 100 USD in your savings account and you save another
# 10 USD in t1
# 20 USD in t2
# 50 USD in t3
# 30 USD in t4
# 25 USD in t5 (each cf at period's end)
# Calculate the FV of your savings account after 5 years given an interest rate of 3% p.a.

cf = [100, 10, 20, 50, 30, 25]
cf

n = list(range(6))
n
# [0, 1, 2, 3, 4, 5]

n = n[::-1]
n
# [5, 4, 3, 2, 1, 0]

f = 1.03

FV = 0
for i in range(6):
    FV += cf[i] * f ** n[i]
    print(FV)

# 115.92740743
# 127.18249553000001
# 149.03703553000003
# 202.08203553
# 232.98203553000002
# 257.98203553

FV = 0
for i in range(6):
    FV += cf[i] * f ** n[i]
print(FV)

# 257.98203553


# Present Values

# Today you agreed on a payout plan that guarantees payouts of
# 50 USD in t1
# 60 USD in t2
# 70 USD in t3
# 80 USD in t4
# 100 USD in t5 (each period's end)
# Calculate the Funding amount/ PV that needs to be paid into the plan today (t0).
# Assume an interest rate of 4% p.a.

cf = [50, 60, 70, 80, 100]
f = 1.04

PV = 0
for i in range(5):
    PV += cf[i] / f ** (i + 1)
    print(PV)

# 48.07692307692307
# 103.55029585798815
# 165.78004096495218
# 234.16437624733024
# 316.3570869232654

PV = 0
for i in range(5):
    PV += cf[i] / f ** (i + 1)
print(PV)

# 316.3570869232654


# =============================================================================
# Calculate an Investment Project's NPV
# =============================================================================

# The XYZ Company evaluates to buy an additional machine that will increase
# future profits/cashflows by
# 20 USD in t1
# 50 USD in t2
# 70 USD in t3
# 100 USD in t4
# 50 USD in t5 (each cf at period's end)

# The machine costs 200 USD (Investment in t0).
# Calculate the Project's NPV and evaluate whether XYZ should pursue the project.
# XYZ's required rate of return (Cost of Capital) is 6% p.a.

cf = [-200, 20, 50, 70, 100, 50]
f = 1.06

NPV = 0
for i in range(6):
    NPV += cf[i] / f ** (i)
    print(NPV)

# -200.0
# -181.1320754716981
# -136.63225347098611
# -77.85890365872498
# 1.3504626650770604
# 38.71337130837991

# Would your conclusion change with a purchase price of 250 USD?
cf[0] = -250
cf
# [-250, 20, 50, 70, 100, 50]
NPV = 0
for i in range(len(cf)):
    NPV += cf[i] / f ** (i)
    print(NPV)

# -250.0
# -231.1320754716981
# -186.63225347098611
# -127.85890365872498
# -48.64953733492294
# -11.286628691620088

cf = [-200, 20, 50, 70, 100, 50, 25]
cf + [25]
cf.append(25)
cf.remove(50)
cf.extend([15])
cf

r1 = 0.05
r1 = r2
r2 = r1
r1 += 0.01
r1, r2