from Recept import Recept
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

