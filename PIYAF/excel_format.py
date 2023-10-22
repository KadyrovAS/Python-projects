import datetime
import openpyxl as op
import openpyxl.styles as st
import os.path
import createTabel as create
def border(sheet: op.worksheet, arg = False, nRowStart = 0, nRowFinis = 0, nColStart = 0, nColFinis = 0):
    # arg = True - бордюр рисуется
    #arg = False - бордюр убирается
    #Если nRowStart, nRowFinis, nColStart, nColFinis не указаны, то применяется ко всему листу
    #Если nRowStart, nRowFinis, nColStart, nColFinis указаны, то применяется только к указанному диапазону
    #Диапазон листа рассчитывается следующим образом:
    # - количество столбцов считается в 1-й строке до первой пустой клетки
    # - количество строк считается по 1-му столбцу до первой пустой клетки
    thin = st.Side(border_style="thin", color="000000") #параметры тонкой линии
    if nRowStart == 0:
        #Определяется диапазон листа
        nRowStart = 1
        nRowFinis = 1
        nColStart = 1
        nColFinis = 1
        #считаем количество строк по 1-му столбцу
        while sheet.cell(row = nRowFinis, column = 1).value is not None:
            nRowFinis += 1
        #считаем количество столбцов по 1-й строке
        while sheet.cell(row = 1, column = nColFinis).value is not None:
            nColFinis += 1
        nRowFinis -= 1
        nColFinis -= 1

    #применяем правило к диапазону
    for row in range(nRowStart, nRowFinis + 1):
        for col in range(nColStart, nColFinis + 1):
            currentCell = sheet.cell(row = row, column = col)
            if arg:
                currentCell.border = st.Border(top=thin, left=thin, right=thin, bottom=thin)
            else:
                currentCell.border=st.Border(top=None, left=None, right=None, bottom=None)

def getMonth(month = 0, arg = 'Name', year = 0):
    if year == 0:
        year = datetime.datetime.now().year

    if month == 0:
        month = datetime.datetime.now().month

    months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
              'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
    monthsR = ['Января', 'Февраля', 'Марта', 'Апреля', 'Мая', 'Июня',
               'Июля', 'Августа', 'Сентября', 'Октября', 'Ноября', 'Декабря']
    days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    if year % 4 == 0:
        days[1] = 29

    if arg == "Name":
        return months[month - 1]
    elif arg == "NameR":
        return monthsR[month - 1]
    elif arg == "count":
        return days[month - 1]


def tabelCreate(spisok, path: str, year = 0, month = 0):
    if year == 0:
        year = datetime.datetime.now().year
        month = datetime.datetime.now().month

    filename = f"{path}\\Табель {year}.xlsx"
    if os.path.exists(filename):
        #Открываем существующий файл
        wbFileTabel = op.open(filename=filename, data_only=False)
    else:
        #Создаем файл
        wbFileTabel = op.Workbook()

    sheets = wbFileTabel.sheetnames
    currentSheet = None
    sheetName = f"{year} {getMonth(month)}"

    for sheet in sheets:
        if sheet == 'Sheet':
            currentSheet = sheet
            wbFileTabel[currentSheet].title = sheetName
            create.createNewTable(spisok, wbFileTabel[sheetName], path)
        elif sheet == sheetName:
            currentSheet = sheet
            break

    if currentSheet is None:
        wbFileTabel.create_sheet(sheetName)
        create.createNewTable(spisok, wbFileTabel[sheetName], path)

    wbFileTabel.save(filename=filename)

def sortSpisok(spisok : list, brench = False):
    if brench:
        start = 0
    else:
        start = 1
    #Сортировка списка по алфавиту
    for i in range(len(spisok) - 1):
        for k in range(i + 1, len(spisok)):
            for j in range(start, 4):
                if spisok[i][j] > spisok[k][j]:
                    spisok[i], spisok[k] = spisok[k], spisok[i]
                    break
                elif spisok[i][j] == spisok[k][j]:
                    continue
                else:
                    break
