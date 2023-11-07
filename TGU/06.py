from pandas import *
from scipy import stats
def set_range(value: str):
    if '+' in value:
        r = value.replace('+', '')
        return int(r) // 10 - 1
    r = list(map(int, value.split("-")))
    return r[0] // 10

def set_int(value: str):
    if '+' in value:
        v = value.replace('+', '')
        return int(v) + 1
    return int(value)

df = read_csv("customer_demographics.csv")
df['family_size'] = df['family_size'].apply(set_int)
df['age_range_level'] = df['age_range'].apply(set_range)


print("Проверяем нормальность распределения данных в столбце ['family_size']")
stat1, p1 = stats.shapiro(df['family_size'])
print(f'Тест Шапиро-Уилк: Statistics = {stat1:.3f}, p-value = {p1:.3f}')
stat2, p2 = stats.normaltest(df['family_size'])
print(f'Тест Пирсона: Statistics = {stat2:.3f}, p-value = {p2:.3f}')
if p1 < 0.05 and p2 < 0.05:
    print("Гипотеза о нормальном распределении в столбце ['family_size'] не подтвердилась")
else:
    print("Гипотеза о нормальном распределении в столбце ['family_size'] не опровергнута")
ass = stats.skew(df['family_size'])
print(f'Ассиметрия данной выборки {ass:.3f}')
exc = stats.kurtosis(df['family_size'])
print(f'Эксцесс выборки {exc: .3f}')
print(f'Стандартное отклонение {df["family_size"].std(): .3f}')
print()

print("Проверяем нормальность распределения данных в столбце ['age_range']")
stat1, p1 = stats.shapiro(df['age_range_level'])
print(f'Тест Шапиро-Уилк: Statistics = {stat1:.3f}, p-value = {p1:.3f}')
stat2, p2 = stats.normaltest(df['age_range_level'])
print(f'Тест Пирсона: Statistics = {stat2:.3f}, p-value = {p2:.3f}')
if p1 < 0.05 and p2 < 0.05:
    print("Гипотеза о нормальном распределении в столбце ['age_range'] не подтвердилась")
else:
    print("Гипотеза о нормальном распределении в столбце ['age_range'] не опровергнута")
ass = stats.skew(df['age_range_level'])
print(f'Ассиметрия данной выборки {ass:.3f}')
exc = stats.kurtosis(df['age_range_level'])
print(f'Эксцесс выборки {exc: .3f}')
print(f'Стандартное отклонение {df["age_range_level"].std(): .3f}')
