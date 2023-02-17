def findChains(item, chains = []):
    global matrix, n, max, chain
    chains.append(item)

    for i in range(item + 1, n):
        if matrix[item][i]:
            if len(chains) > max:
                max = len(chains)
            findChains(i)
            chains.remove(i)
    if len(chains) > max:
        max = len(chains)
        chain = chains.copy()

from Recept import Recept
from Tree import Tree

# n, k = map(int, input().split(" "))

max = 0
chain = []
tree = Tree()

# --------------------------------------------------
ar = [[3, 1, 2, 3],[3, 1, 2, 4],[4, 1, 2, 3, 4],[5, 1, 2, 3, 4, 5],[2, 1, 2]]
n = len(ar)
# --------------------------------------------------


for i in range(n):
    # recept = list(map(int, input().split(" ")))
    recept = ar[i]
    tree.put(Recept(recept))

recepts = []
tree.head.process(recepts)

matrix = [[False for _ in range(n)] for _ in range(n)]

# определяем какие множества в какие входят
for i in range(n - 1):
    for j in range(i + 1, n):
        if recepts[i].recept <= recepts[j].recept:
            matrix[i][j] = True

for row in recepts:
    print(row)
    
num = -1
for i in range(n - 1):
    for k in range(i + 1, n):
        if matrix[i][k]:
            num = i
            break
    if num > -1:
        break

findChains(num)
print(max, chain)