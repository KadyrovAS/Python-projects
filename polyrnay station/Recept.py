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