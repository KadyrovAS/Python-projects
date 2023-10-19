from pandas import Series, DataFrame

df1 = DataFrame(data = {
    "shop": [427, 707, 957, 437],
    "qte": [3, 4, 2, 1]
}
)

df2 = DataFrame(data = {
    "shop": [347, 427, 707, 957, 437],
    "name": ['Киев', 'Самара', 'Минск', 'Иркутск', 'Москва']
},
 index=[1, 2, 3, 4, 5]
)

country = [u'Украина', u'РФ', u'Беларусь', u'РФ', u'РФ']
df2.insert(1, 'country', country)
print(df2)
print("-" * 30)
df = df2[df2['country'] == "Украина"]
# df.shop = 345
print(df)
print("-" * 30)


df3 = df2.merge(df1, 'left', on='shop')
print(df3)
print("-" * 30)

res = df3.pivot_table(['qte'], ['country'], aggfunc='sum', fill_value=0)
print(res)