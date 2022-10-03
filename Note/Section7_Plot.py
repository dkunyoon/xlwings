# ============================================================
# Writing a Matplotlib plot into Excel
# ============================================================

import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns


# Study: Adult Male Height in the U.S. (normally distributed)

mean_height = 70
std_height = 4
sample_size = 1000000

sample = np.random.normal(mean_height, std_height, sample_size)
sample
# array([68.45993568, 60.58375704, 74.33996001, ..., 76.66458287,
#        67.01443366, 74.54581513])

sample.size
# 1000000

median_height = np.median(sample)
median_height

fig = plt.figure(figsize=(12, 8))
plt.hist(sample, bins=1000, alpha=0.5, label="Frequency")
plt.vlines(x=median_height, ymin=0, ymax=4000,
           linestyle="--", label="Median Height")
plt.annotate(round(median_height, 1), xy=(median_height-3, 2000), fontsize=17)
plt.title("Adult Male Height in the U.S. (Sample Size: 1 Mio.)", fontsize=20)
plt.xlabel("Height in Inches", fontsize=15)
plt.ylabel("Frequency", fontsize=15)
plt.legend(fontsize=15)
plt.show()


wb = xw.Book()
sheet = wb.sheets[0]

sheet.pictures.add(fig)
# <Picture 'Picture 4' in <Sheet [통합 문서1]Sheet1>>


# ============================================================
# Updating the Plot
# ============================================================

sheet.pictures.add(fig)
# <Picture 'Picture 6' in <Sheet [통합 문서1]Sheet1>>

sheet.pictures.add(fig, name="Height", update=True)
# this just updates the pic in excel

plt.style.use("seaborn")
# once executed, re-run the plot


# ============================================================
# Changing Size and Position
# ============================================================

fig = plt.figure(figsize=(12, 8))
plt.hist(sample, bins=1000, alpha=0.5, label="Frequency")
plt.vlines(x=median_height, ymin=0, ymax=4000,
           linestyle="--", label="Median Height")
plt.annotate(round(median_height, 1), xy=(median_height-3, 2000), fontsize=17)
plt.title("Adult Male Height in the U.S. (Sample Size: 1 Mio.)", fontsize=20)
plt.xlabel("Height in Inches", fontsize=15)
plt.ylabel("Frequency", fontsize=15)
plt.legend(fontsize=15)
plt.show()

sheet.pictures.add(fig, name="Height", update=True)

plot = sheet.pictures.add(fig, name="Height", update=True)

type(plot)
# xlwings.main.Picture

plot.left
plot.top

# reposition the plot in excel
plot.left = 100
plot.top = 50

# reposition the plot in excel
plot.left = sheet["B2"].left
plot.top = sheet["B2"].top


plot.height
# 505.4410095214844
plot.width
# 734.6414794921875

# resizing the plot
plot.height /= 2
# 252.72047424316406
plot.width /= 2
# 367.3207092285156

plot.height
plot.width *= 2


# ============================================================
# Changing Size and Position (Part 2)
# ============================================================

# pictures.add method 안에 파라미터 정의하는경우, 엑셀 업데이트 안됨
# 엑셀에서 remove 한 후 코드 다시돌려야함
sheet.pictures.add(fig, name="Height", update=True,
                   left=sheet.range("B2").left,
                   top=sheet.range("B2").top,
                   scale=0.6)
