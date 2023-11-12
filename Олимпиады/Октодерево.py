def dToB(value):
    res = f"{int(value)}."
    if value % 1 == 0:
        res = f"{res}0"
        return res
    while value % 1 != 0:
        value *= 2
        b = int(value)
        res = f"{res}{0 if b == 0 else 1}"
        value = value % 1
    return res
class Tree:
    def __init__(self, node):
        self.head = node

    def tree_sum(self, simbol, node=None):
        sum = 0
        if node is None:
            return self.tree_sum(simbol, self.head)

        if node.color == simbol:
            sum += node.level

        for node_branch in node.branch:
            sum += self.tree_sum(simbol, node_branch)


        return sum
    def set_level(self):
        node = self.head
        node.set_level(1)
class Node:
    def __init__(self, color):
        self.branch = []
        self.parent = None
        self.color = color

    def set_parent(self, node):
        self.parent = node
        if not self in node.branch:
            node.branch.append(self)

    def set_level(self, level):
        self.level = level
        for branch in self.branch:
            branch.set_level(level / 8)

def create_tree(tree, parent=None):
    global line, cursor
    if cursor == len(line):
        return
    color = line[cursor]
    node = Node(color)
    if cursor == 0:
        tree.head = node

    if color == "Q":
        if parent is not None:
            node.set_parent(parent)
        for _ in range(8):
            cursor += 1
            create_tree(tree, node)
    else:
        if parent is not None:
            node.set_parent(parent)


line = input()
color_cube = input()

cursor = 0
tree = Tree(Node("Q"))
create_tree(tree)
tree.set_level()

sum = tree.tree_sum(color_cube)
sum_2 = dToB(sum)
print(sum_2)