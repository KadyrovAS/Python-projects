from pandas import *
from scipy import stats
import seaborn as sb
import matplotlib.pyplot as plt
def set_weight(weight):
    global weight_minimum, weight_maximum
    r = (weight_maximum - weight_minimum) / 10
    n =  (weight - weight_minimum) // r
    scale = f"{weight_minimum + n * r:.2f}-{weight_minimum + (n + 1) * r:.2f}"
    return scale

def set_height(height):
    global height_minimum, height_maximum
    r = (height_maximum - height_minimum) / 10
    n =  (height - height_minimum) // r
    scale = f"{height_minimum + n * r:.2f}-{height_minimum + (n + 1) * r:.2f}"
    return scale

df = read_csv("weight-height.csv")
genders ={"Мужчин" : "Male", "Женщин" : "Female"}
parameters = {"Рост" : "Height", "Вес" : "Weight"}

n = 0
plt.suptitle("Распределение параметров Рост и Вес у мужчин и женщин")
for gender_rus, gender_eng in genders.items():
    df_gender = df[df["Gender"] == gender_eng]
    for parameter_rus, parameter_eng in parameters.items():
        n += 1
        print(f"Проверяем нормальность распределения параметра {parameter_rus} для {gender_rus}")
        stat1, p1 = stats.shapiro(df_gender[parameter_eng])
        print(f'Тест Шапиро-Уилк: Statistics = {stat1:.3f}, p-value = {p1:.3f}')
        stat2, p2 = stats.normaltest(df_gender[parameter_eng])
        print(f'Тест Пирсона: Statistics = {stat2:.3f}, p-value = {p2:.3f}')
        if p1 < 0.05 and p2 < 0.05:
            print(f"Гипотеза о нормальном распределении параметра {parameter_rus} для {gender_rus} не подтвердилась")
        else:
            print(f"Гипотеза о нормальном распределении параметра {parameter_rus} для {gender_rus} не опровергнута")
        plt.subplot(2, 2, n)
        plt.title(f"{parameter_rus} для {gender_rus}")
        plt.hist(df_gender[parameter_eng])

df_gender = df[df["Gender"] == gender_eng]
# plt.show()

height_minimum = df["Height"].min()
height_maximum = df["Height"].max()

weight_minimum = df["Weight"].min()
weight_maximum = df["Weight"].max()

df["Weight_Scale"] = df["Weight"].apply(set_weight)
df["Height_Scale"] = df["Height"].apply(set_height)

for gender_rus, gender_eng in genders.items():
    df_gender = df[df["Gender"] == gender_eng]
    df_table = df_gender.pivot_table(index="Weight_Scale", columns="Height_Scale", values="Gender", aggfunc = "count")
    sb.heatmap(df_table, cmap="crest")
    plt.show()