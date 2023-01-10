class Recept:
    def __init__(self, recept: list):
        self.recept = set(recept[1:])
        self.length = recept[0]
        self.children = []
        self.parent = None
        self.left = None
        self.right = None

        self.level = 0


    def __str__(self):
        line = f"{self.recept} "
        if self.parent is not None:
            line += f"parent: {self.parent.recept} "
        if self.left is not None:
            line += f"left: {self.left.recept} "
        if self.right is not None:
            line += f"right: {self.right.recept} "
        line += "\n"
        return line