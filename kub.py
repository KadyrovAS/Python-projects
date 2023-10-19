#Задача с кубами
import numpy as np

n = int(input())
count = 0
while n > 0:
    count += 1
    n -= int(np.cbrt(n)) ** 3

print(count)