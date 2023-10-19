import openpyxl as op
import excel_format

fl = open("tabel.ini", "r", encoding="utf-8")
lines = fl.readlines()
dir = None
fn = None

for item in lines:
    com = item.split("=")
    if com[0].strip() == "папка":
        dir = com[1].strip()
    elif com[0].strip() == "актуальный файл":
        fn = com[1].strip()

if dir is None:
    print("Отсутствует значение 'папка' в ini файле")
if fn is None:
    print("Отсутствует значение 'актуальный файл' в ini файле")

file_path = dir + "\\" + "Шаблон.xlsx"

excel_doc = op.open(filename = file_path, data_only = False)
sheetnames = excel_doc.sheetnames
sheet = excel_doc[sheetnames[0]]

spisok = []
nColumn = 1
while sheet.cell(row = 1, column = nColumn).value is not None:
    nColumn += 1

nColumn -= 1
row = 2
while sheet.cell(row = row, column = 1).value is not None:
    person = [sheet.cell(row = row, column=col).value for col in range(1, nColumn + 1)]
    spisok.append(person)
    row += 1

for i in range(len(spisok) - 1):
    for k in range(i + 1, len(spisok)):
        for j in range(nColumn):
            if spisok[i][j] is None or spisok[k][j] is None:
                continue
            arg1 = spisok[i][j]
            arg2 = spisok[k][j]
            if spisok[i][j] == spisok[k][j]:
                continue
            elif spisok[i][j] < spisok[k][j]:
                break
            else:
                spisok[i], spisok[k] = spisok[k], spisok[i]
                break
row = 2
for person in spisok:
    col = 1
    for val in person:
        sheet.cell(row = row, column = col).value = val
        col += 1
    row += 1

excel_format.border(sheet, True, 23, 25, 1, 8)

excel_doc.save(file_path)
excel_doc.close

# excel_doc.create_sheet(title = "Проверочный лист", index = 0)
# excel_doc.save(file_path)