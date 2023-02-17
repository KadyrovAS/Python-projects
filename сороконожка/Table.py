class Table:
    def __init__(self):
        self.badBootCount = 0
        self.goodBootCount = 0
    def putBadBoot(self):
        #Положить на стол плохой сапог
        self.badBootCount += 1
    def getBadBoot(self):
        #Взять плохой сапог
        self.badBootCount -= 1
    def putGoodBoot(self):
        #Положить на стол хороший сапог
        self.goodBootCount += 1

    def getGoodBoot(self):
        #Взять со стола плохой сапог
        self.goodBootCount -= 1

    def isGoodBoot(self):
        #На столе лежит отремонтированный сапог
        if self.goodBootCount > 0:
            return True
        else:
            return False

    def isBadBoot(self):
        #На столе лежит снятый старый сапог
        if self.badBootCount > 0:
            return True
        else:
            return False
