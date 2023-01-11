class Recept:
    def __init__(self, recept: list):
        self.recept = set(recept[1:])
        self.length = recept[0]
        self.children = []
        self.parent = None
        self.union_parent = None
        self.left = None
        self.right = None
        self.mark = False
        self.level = 0

    def __str__(self):
        line = f"{self.recept} level = {self.level} "
        if self.parent is not None:
            line += f"parent: {self.parent.recept} "

        if len(self.children) > 0:
            line += "\n Дети \n"
            for value in self.children:
                line += f"{value.recept} \n"
        line += "\n"

        return line

    def set_parent(self, node):
        self.union_parent = node

    def add_child(self, node):
        self.children.append(node)
        node.set_level(self.level + 1)

    def set_level(self, level):
        self.level = level
        for node in self.children:
            node.set_level(level + 1)

    def process(self, ar):
        if self.left is not None:
            self.left.process(ar)
        ar.append(self)
        if self.right is not None:
            self.right.process(ar)

    def set_main(self):
        self.mark = True
        if self.union_parent is not None:
            for node in self.union_parent.children:
                if node == self:
                    continue
                node.set_level(0)
                self.union_parent.children.remove(node)
            self.union_parent.set_main()


class Tree:
    def __init__(self):
        self.head = None

    def put(self, recept: Recept):
        if self.head is None:
            self.head = recept
            return

        branch: Recept = self.head
        parent: Recept = None
        conerLeft = True
        while True:
            if branch is None:
                branch = recept
                branch.parent = parent
                if conerLeft:
                    branch.parent.left = branch
                else:
                    branch.parent.right = branch
                return
            parent = branch
            if branch.length > recept.length:
                branch = branch.left
                conerLeft = True
            else:
                branch = branch.right
                conerLeft = False


n, k = map(int, input().split(" "))

tree = Tree()

for i in range(n):
    recept = list(map(int, input().split(" ")))
    tree.put(Recept(recept))

recepts = []
tree.head.process(recepts)

# определяем какие множества в какие входят
for i in range(n - 1):
    for j in range(i + 1, n):
        if recepts[i].recept <= recepts[j].recept:
            recepts[i].set_parent(recepts[j])
            recepts[j].add_child(recepts[i])
            break

# помечаем гланые ветки
was_found = True
while was_found:
    was_found = False
    max = 0
    node = None
    for i in range(n):
        if not recepts[i].mark and recepts[i].level > max:
            max = recepts[i].level
            node = recepts[i]
            was_found = True
    if was_found:
        node.set_main()

count = 0
for node in recepts:
    if node.level == 0:
        count += 1

print(count)
