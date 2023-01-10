from BinaryTree import Tree
from Recept import Recept

# n, k = map(int, input().split(" "))

tree = Tree()

n = 4
ar_in = [[3, 1, 2, 4], [4, 1, 2, 3, 4], [3, 1, 2, 3], [5, 1, 2, 3, 4, 5]]

for i in range(n):
    # recept = list(map(int, input().split(" ")))
    recept = ar_in[i]
    tree.put(Recept(ar_in[i]))

tree.getRecepts()