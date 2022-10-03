import numpy as np


np.random.seed(123)
np.random.normal(10, 2.5, (10, 10))

np.random.seed(123)
np.random.normal(10, 2.5)


npempty = np.empty((2, 2))

npempty[0] = [1, 1]
npempty[1] = [2, 2]
npempty

for i in range(10):
    print(i)