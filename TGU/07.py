from pandas import *
import matplotlib.pyplot as plt
import seaborn as sb
import numpy as np
def prt(value):
    print(value)

df = read_csv("diamonds.csv")
print(df.columns)
df["carat"].apply(prt)
df["carat"].value_counts().plot.density()
plt.show()