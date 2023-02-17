class Master0:
    def __init__(self, time, table, countBoots):
        self.table = table
        self.countBootBad = countBoots
        self.countBootGood = 0
        self.withoutBoot = 0
        self.maxWithoutBoot = 0
        self.operationTime = time #Время на операцию
        self.timeStart = 0
        self.timeFinis = 0
        self.operation = ""
        self.isBusy = False

    def do(self, time):

        if time == self.timeFinis and self.operation == "get":
            self.table.putBadBoot()
            self.operation = ""
            self.isBusy = False
            print(f"Время: {time} Master0 положил на стол старый сапог")
            return

        if time == self.timeFinis and self.operation == "put":
            self.operation = ""
            self.isBusy = False
            print(f"Время: {time} Master0 одел на сороконожку отремонтированный сапог")
            return

        if self.isBusy:
            return

        self.timeStart = time  # Время начала операции
        self.timeFinis = time + self.operationTime #Время конца операции
        if self.table.isGoodBoot():
            print(f"Время: {time} Master0 взял со стола отремонтированный сапог и начал одевать на сороконожку")
            self.operation = "put"
            self.table.getGoodBoot()
            self.putBoot()
        else:
            print(f"Время: {time} Master0 снял старый сапог с сороконожки")
            self.operation = "get"
            self.getBoot()

    def putBoot(self):
        self.isBusy = True
        #Мастер одевает один отремонтированный сапог
        self.countBootGood += 1
        self.withoutBoot -= 1

    def getBoot(self):
        #Мастер снимает один плохой сапог
        if self.countBootBad == 0:
            return
        self.isBusy = True
        self.countBootBad -= 1
        self.withoutBoot += 1
        if self.withoutBoot > self.maxWithoutBoot:
            self.maxWithoutBoot = self.withoutBoot