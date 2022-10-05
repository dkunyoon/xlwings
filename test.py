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
    
2/6

capitals = {"India":"New Delhi", "Nigeria":"Abuja", "USA":"Washington, DC"}
capitals

capitals.update({"Pakistan":"Islam"})


capitals["Pakistan"] = "islam"
capitals


(4**2 * 3) +15

import numpy as np
np.mean([3,3,4,5,10])

my_string = "I'm ready"
my_string[4:8]

# printdang@naver.com