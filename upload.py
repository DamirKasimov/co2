# здесь шаг №1, лист №2
import pandas
import PySimpleGUI as sg
import openpyxl


# рендер окна загрузки маленькой таблицы, шаг 1
def first_table():
    global f
    layout = [[
            sg.Input(key='-Input-'),
            sg.FileBrowse(button_text="Выбрать", key='-IN-')],
          [sg.Button("Загрузить *.xlsx файл")]]
    window = sg.Window("Загрузить исходные данные, шаг 1",
                       layout=layout, size=(700, 250))
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "Exit":
            break
        elif event == "Загрузить *.xlsx файл":
            filename = values['-IN-']
            # загруженную таблицу передаем в переменную 'f'
            f = pandas.read_excel(filename, sheet_name='2',
                                  engine='openpyxl', nrows=5)
        window.close()
    return(f)
