from Master0 import Master0
from Master import Master
from Table import Table

count_legs = 40
table = Table()
master0 = Master0(0.5, table, count_legs) #Подмастерье
master1 = Master(3, table, 1)
master2 = Master(2, table, 2)

step = 0.5
time = 0

while master0.countBootGood != count_legs:
    for _ in range(2):
        master0.do(time)
        master2.do(time)
        master1.do(time)

    time += step

print(f"Максимальное количество ног без сапог равно {master0.maxWithoutBoot}")