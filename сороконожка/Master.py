
class Master:
    def __init__(self, time, table, id):
        self.table = table
        self.operationTime = time #Время на операцию
        self.timeStart = 0
        self.timeFinis = 0
        self.operation = ""
        self.id = id
        self.isBusy = False

    def do(self, time):
        #Мастер ремонтирует один сапог

        if time == self.timeFinis and self.operation == "repair":
            print(f"Время: {time} Master{self.id} положил на стол отремонтированный сапог")
            self.table.putGoodBoot()
            self.operation = ""
            self.isBusy = False
            return

        if self.isBusy: #Мастер занят
            return

        if self.table.isBadBoot(): #На столе находится плохой сапог
            print(f"Time: {time} Master{self.id} взял со стола старый сапог на ремонт")
            self.isBusy = True
            self.table.getBadBoot() #Взять со стола плохой сапог на ремонт
            self.operation = "repair"
            self.timeStart = time #Время начала операции
            self.timeFinis = time + self.operationTime