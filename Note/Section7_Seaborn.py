# ============================================================
# Adding a Seaborn Plot
# ============================================================

from os import set_inheritable
import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

mean_height = 70
std_height = 4
sample_size = 1000

sample = np.random.normal(mean_height, std_height, sample_size)
sample

median_height = np.median(sample)
median_height

sns.distplot(sample, hist=True, kde=True, rug=True)
plt.show()

fig = plt.figure(figsize=(12,8))
sns.distplot(sample, hist = False, kde= True, rug = True, label = "KDE")
plt.vlines(x = median_height, ymin = 0, ymax = 0.1, linestyle = "--", label = "Median Height")
plt.annotate(round(median_height,1), xy = (median_height-3, 0.05), fontsize = 17)
plt.title("Adult Male Height in the U.S. (Sample Size: 1,000)", fontsize = 20)
plt.xlabel("Height in Inches", fontsize = 15)
plt.legend(fontsize = 15)
plt.show()


wb = xw.Book()
sheet = wb.sheets[0]
plot = sheet.pictures.add(fig, name="Height", update=True)

plot.left = sheet["B2"].left
plot.top = sheet["B2"].top

plot.height /= 2
plot.width /= 2

sheet.pictures.add(fig, name="Height", update=True,
                   left=sheet["B2"].left,
                   top=sheet["B2"].top,
                   scale=0.5)





