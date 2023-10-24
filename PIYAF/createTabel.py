import openpyxl as op
import openpyxl.styles as st
from openpyxl.utils import get_column_letter

import excel_format


def createHead(workSheet, rowStart, n):
    # Объединение ячеек
    workSheet.merge_cells(f'B{rowStart}:C{rowStart}')

    letter = get_column_letter(n + 6)
    workSheet.merge_cells(f'E{rowStart}:{letter}{rowStart}')

    workSheet.merge_cells(f'A{rowStart}:A{rowStart + 1}')
    workSheet.merge_cells(f'D{rowStart}:D{rowStart + 1}')

    # Заполняем шапку таблицы
    workSheet.cell(row=rowStart, column=1).value = "Фамилия, имя, отчество"
    workSheet.cell(row=rowStart, column=1).alignment = st.Alignment(wrap_text=True,
                                                                    horizontal='center', vertical='center')
    workSheet.cell(row=rowStart, column=2).value = "Учетный номер"
    workSheet.cell(row=rowStart + 1, column=2).value = "Табельный номер"
    workSheet.cell(row=rowStart + 1, column=3).value = "Премия"
    workSheet.cell(row=rowStart, column=4).value = "Должность (профессия)"
    workSheet.cell(row=rowStart, column=4).alignment = st.Alignment(wrapText=True,
                                                                    horizontal='center', vertical='center')
    workSheet.cell(row=rowStart, column=5).value = "Числа месяца"
    col = 4
    for i in range(n):
        col += 1
        if i == 15:
            workSheet.cell(row=rowStart + 1, column=col).value = "Итого дней (часов) явок (неявок) с 1 по 15"
            workSheet.cell(row=rowStart + 1, column=col).alignment = st.Alignment(wrap_text=True,
                                                                                  horizontal='center',
                                                                                  vertical='center')
            col += 1
        workSheet.cell(row=rowStart + 1, column=col).value = i + 1

    col += 1
    workSheet.cell(row=rowStart + 1, column=col).value = "Всего дней (часов) явок (неявок) за месяц"
    workSheet.cell(row=rowStart + 1, column=col).alignment = st.Alignment(wrap_text=True,
                                                                          horizontal='center', vertical='center')

    workSheet[f'B{rowStart + 1}'].alignment = st.Alignment(textRotation=90, horizontal='center', vertical='center')
    workSheet[f'C{rowStart + 1}'].alignment = st.Alignment(textRotation=90, horizontal='center', vertical='center')
    alignment = st.Alignment(horizontal='center', vertical='center')

    for i in range(1, col + 1):
        workSheet.cell(row=rowStart + 2, column=i).value = i
        workSheet.cell(row=rowStart + 2, column=i).alignment = alignment


def createNewTable(spisok, workSheet, path):
    font = st.Font(
        size=8,
        bold=False,
        italic=False,
        color="000000",
        name="Arial"
    )
    alignment = st.Alignment(wrap_text=True, horizontal='center', vertical='center')

    # ориентация страницы
    workSheet.page_setup.orientation = workSheet.ORIENTATION_LANDSCAPE
    workSheet.page_setup.paperSize = workSheet.PAPERSIZE_A4
    # зумирование по ширине
    workSheet.sheet_properties.pageSetUpPr.fitToPage = True
    workSheet.page_setup.fitToWidth = False

    for row in range(1, 1000):
        for col in range(1, 40):
            workSheet.cell(row=row, column=col).font = font
            workSheet.cell(row=row, column=col).alignment = alignment

    pathShablon = f"{path}\\Шаблон.xlsx"
    wbShablon = op.open(filename=pathShablon, data_only=False)
    ws = wbShablon['Настройки']
    row = 2
    spr = dict()
    while ws.cell(row=row, column=1).value is not None:
        cl = ws.cell(row=row, column=1).value
        clValue = ws.cell(row=row, column=2).value
        clValue = float(clValue)
        if cl == "ФИО":
            spr.update({"FIO": clValue})
        elif cl == "Табельный номер":
            spr.update({'Tabel': clValue})
        elif cl == "Премия":
            spr.update({'Prem': clValue})
        elif cl == "Должность":
            spr.update({'Dol': clValue})
        elif cl == "Дата":
            spr.update({'Data': clValue})
        elif cl == "Итого":
            spr.update({'Itog': clValue})
        elif cl == "Количество":
            spr.update({'Count': clValue})
        row += 1

    workSheet.column_dimensions['A'].width = spr['FIO']  # фамилия имя отчество
    workSheet.column_dimensions['B'].width = spr['Tabel']  # Табельный номер
    workSheet.column_dimensions['C'].width = spr['Prem']  # Премия
    workSheet.column_dimensions['D'].width = spr['Dol']  # Должность
    col = 4
    for i in range(15):
        col += 1
        letter = get_column_letter(col)
        workSheet.column_dimensions[letter].width = spr['Data']  # Даты

    col += 1
    letter = get_column_letter(col)
    workSheet.column_dimensions[letter].width = spr['Itog']

    n = excel_format.getMonth(arg='count')
    for i in range(16, n + 1):
        col += 1
        letter = get_column_letter(col)
        workSheet.column_dimensions[letter].width = spr['Data']  # Даты

    col += 1
    letter = get_column_letter(col)
    workSheet.column_dimensions[letter].width = spr['Itog']

    createHead(workSheet, 11, n)

    # Заполняем таблицу списком

    row = 14
    for i in range(len(spisok)):
        worker = spisok[i]

        workSheet.merge_cells(f'A{row}:A{row + 1}')
        workSheet.cell(row=row, column=1).alignment = alignment
        workerName = f'{worker[1]} {worker[2][0]}.' + ('' if worker[3] == '' else (worker[3][0] + '.'))
        workSheet.cell(row=row, column=1).value = workerName

        workSheet.merge_cells(f'B{row}:B{row + 1}')
        workSheet.cell(row=row, column=2).alignment = alignment
        workSheet.cell(row=row, column=2).value = worker[5]

        workSheet.merge_cells(f'D{row}:D{row + 1}')
        workSheet.cell(row=row, column=4).alignment = alignment
        workSheet.cell(row=row, column=4).value = worker[4]

        row += 2
        if (i + 1) % spr['Count'] == 0:
            # next_page_horizon, next_page_vertical = workSheet.page_breaks
            # next_page_horizon.append(Break(row))
            createHead(workSheet, row, n)
            row += 3

    excel_format.border(workSheet, True, 11, row - 1, 1, n + 6)
    createMainHead(workSheet)

def createMainHead(workSheet):
    n = excel_format.getMonth(arg = 'count')
    month = excel_format.getMonth(arg = 'month')
    column = get_column_letter(n + 3)
    alignment = st.Alignment(horizontal='center', vertical='center')
    font1 = st.Font(
        size=14,
        bold=True,
        italic=False,
        color="000000",
        name="Arial"
    )

    workSheet.merge_cells(f'A1:{column}1')
    workSheet.cell(row=1, column=1).value = f'Табель №{month}'
    workSheet.cell(row=1, column=1).alignment = alignment
    workSheet.cell(row=1, column=1).font = font1

    font2 = st.Font(
        size=11,
        bold=True,
        italic=False,
        color="000000",
        name="Arial"
    )
    workSheet.merge_cells(f'A2:{column}2')
    workSheet.cell(row=2, column=1).value = 'учета использования рабочего времени'
    workSheet.cell(row=2, column=1).alignment = alignment
    workSheet.cell(row=2, column=1).font = font2

    column1_letter = get_column_letter(n + 4)
    column2_letter = get_column_letter(n + 6)

    font3 = st.Font(
        size=8,
        bold=False,
        italic=False,
        color="000000",
        name="Arial"
    )

    for row in range(2, 10):
        workSheet.merge_cells(f'{column1_letter}{row}:{column2_letter}{row}')
        workSheet.cell(row = row, column = n + 4).alignment = alignment
        workSheet.cell(row = row, column = n + 4).font = font3

    workSheet.cell(row = 2, column = n + 4).value = "Коды"
    workSheet.cell(row = 3, column = n + 4).value = "0504421"
    workSheet.cell(row = 4, column = n + 4).value = excel_format.getMonth(arg = 'date')
    workSheet.cell(row = 5, column = n + 4).value = "02698654"
    workSheet.cell(row = 8, column = n + 4).value = "0"
    workSheet.cell(row = 9, column = n + 4).value = excel_format.getMonth(arg = 'date')

    excel_format.border(workSheet, True, 2, 9, n + 4, n + 6)