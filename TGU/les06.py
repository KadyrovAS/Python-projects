from pandas import *
import seaborn as sb
import matplotlib.pyplot as plt

df = read_csv('weight-height.csv')
sb.scatterplot(data=df, x = df['Height'], y = df['Weight'])
plt.show()