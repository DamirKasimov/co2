# здесь шаг №2, лист №1

import pandas
import openpyxl
import PySimpleGUI as sg
import numpy_financial

MULTIP = 0.1  # задаем дисконтирование


# окно загрузки большой таблицы с листа 1, из файла "CO2_очищенная"
def final_table():
    layout = [[
            sg.Input(key='-Input-'),
            sg.FileBrowse(button_text="Выбрать", key='-IN-')],
            [sg.Button("Загрузить *.xlsx файл")]]

    window = sg.Window("Загрузить исходные данные, шаг 2",
                       layout=layout, size=(700, 150))

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "Exit":
            break
        # по нажатию кнопки "Загрузить *.xlsx файл" грузим большую таблицу
        # с листа 1, из файла "CO2_очищенная"
        elif event == "Загрузить *.xlsx файл":
            filename = values['-IN-']
            f2 = pandas.read_excel(filename, sheet_name='1',
                                   header=None,
                                   engine='openpyxl',
                                   usecols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
                                   nrows=36)
            f2 = f2.fillna(0)
            window.close()
    return(f2)


# пересчет ячеек, зависящих от суммы налога,
# в загруженной большой таблице с листа '1'
def make_table(num_rows, num_cols):
    global table_data
    wb = openpyxl.load_workbook('./temp.xlsx')  # загрузка файла temp.xlsx
    sheet = wb.get_sheet_by_name('temp')
    # загрузка скорректированного значения из файла temp.xlsx
    c = -abs(float((sheet['A1'].value)))
    total_c = c*8
    # инстанцирование пустого шаблона таблицы с типом данных 'списки в списке'
    # для последующего заполнения и графического отображения
    table_data = [[j for j in range(num_cols)] for i in range(num_rows)]
    for i in range(0, (num_rows)):
        table_data[i] = [*f2.loc[i]]  # заполняем шаблон данными 'f2' построчно из загруженной большой таблицы, как есть
        # перезаписываем строку 32 скорректированным значением из файла temp.xlsx
        table_data[31] = ['Расчетное влияние налога на выбросы углерода (начиная с момента реализации капиталовложений) («+» экономия / «-» затраты)', c,c,c,c,c,c,c,c, total_c]
        # перезаписываем строку 33
        total_33 = 0
        if i == 32:
            discount = (1+MULTIP)**(-0.98)
            table_data[32][1] = (float(table_data[31][1]))*(discount)
            total_33 = total_33 + table_data[32][1]
            for column in range(3, num_cols):
                year = -(column-2)
                discount = (1+MULTIP)**year
                offset_col = column-1
                table_data[32][column-1] = (float(table_data[31][offset_col]))*(discount)
                total_33 = total_33 + table_data[32][column-1]
                table_data[32][num_cols-1] = total_33
        # перезаписываем строку 34
        accrued_total = 0
        for column in range(1, num_cols):
            accrued_total = table_data[32][column] + accrued_total
            table_data[33][column] = accrued_total
        table_data[33][9] = table_data[33][8]
        # перезаписываем строку 35
        for column in range(1, num_cols):
            table_data[34][column] = table_data[32][column] + table_data[28][column]
        # перезаписываем строку 36
        for column in range(1, num_cols):
            table_data[35][column] = table_data[33][column] + table_data[29][column]
    return(table_data)


# рендер большой таблицы с перерасчитанными данными
def building_final_table():
    global f2
    f2 = final_table()
    num_rows = f2.shape[0]
    num_cols = f2.shape[1]
    values = make_table(num_rows, num_cols)  #здесь забираем из функции make_table готовую таблицу с перерасчитанными данными
    headings = ('1','2','3','4','5','6','7','8','9','10')
    layout = [
        [sg.Table(values=values,
                  headings=headings,
                  expand_x=True,
                  expand_y=True,
                  key='-TABLE-')],
        [sg.Button('Основные показатели')], ]

    # рендер окна "Расчет окупаемости" с основными показателями
    window = sg.Window('Расчет окупаемости', layout=layout, resizable=True, auto_size_text=True)
    while True:
        event, values = window.read()
        irr = numpy_financial.irr(table_data[27][1:9])
        if event == "Основные показатели":
            sg.popup(f'Срок окупаемости в годах: {round((float(table_data[0][6]) + abs(float(table_data[29][6]))/float(table_data[28][7])), 4)}',
                    f'Срок окупаемости в годах. вкл. углеродный сбор: {round((float(table_data[0][6])) + abs(float(table_data[35][6]))/float(table_data[34][7]),4)}',
                    f'Чистая приведенная стоимость (7 лет): {round(float(table_data[29][8]),2)} тыс. руб.',
                    f'Внутренняя норма доходности: {round(irr,4)*100}%',
                    title='Основные показатели')
