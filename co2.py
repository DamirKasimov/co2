import PySimpleGUI as sg
from openpyxl import Workbook
import upload
import final_tables_totals as ftt
import openpyxl


# получаем данные сырой таблицы 'маленькой' из модуля 'upload'
def number(i, num_cols):
    return (upload.f.iloc[i, 1:(num_cols)])


# формируем данные для отображения загруженной таблицы (маленькой), шаг 1
def make_table(num_rows, num_cols):
    data = [[j for j in range(num_cols)] for i in range(num_rows)]
    for i in range(0, num_rows):
        data[i] = [rows[i], *number(i, num_cols)]
    return data


# функция-редактор на экране загруженной таблицы "Расчет углеродного сбора"
def edit_cell(window, key, row, col, justify='left'):
    global textvariable, edit, text

    def callback(event, row, col, text, key):
        global edit
        widget = event.widget
        if key == 'Return':
            text = widget.get()
            wb = Workbook()  # выгружаем новое значение ячейки в файл temp.xlsx
            ws1 = wb.active
            ws1 = wb.create_sheet("temp", 0)
            ws1.append([text])
            wb.save(filename="temp.xlsx")
        widget.destroy()
        widget.master.destroy()
        values = list(table.item(row, 'values'))
        values[col] = text
        table.item(row, values=values)
        edit = False

    if edit or row <= 0:
        return

    edit = True
    root = window.TKroot
    table = window[key].Widget

    text = table.item(row, "values")[col]
    x, y, width, height = table.bbox(row, col)

    frame = sg.tk.Frame(root)
    frame.place(x=x, y=y, anchor="nw", width=width, height=height)
    textvariable = sg.tk.StringVar()
    textvariable.set(text)
    entry = sg.tk.Entry(frame, textvariable=textvariable, justify=justify)
    entry.pack()
    entry.select_range(0, sg.tk.END)
    entry.icursor(sg.tk.END)
    entry.focus_force()
    entry.bind("<Return>", lambda e, r=row, c=col, t=text, k='Return': callback(e, r, c, t, k))
    entry.bind("<Escape>", lambda e, r=row, c=col, t=text, k='Escape': callback(e, r, c, t, k))


def main_example():
    global edit
    global rows
    f = upload.first_table()
    columns = []
    for column in range(len(f.columns)):
        columns.append(f.columns[column])
    rows = []
    for row in range(5):
        rows.append(f['РАСЧЕТ УГЛЕРОДНОГО СБОРА'][row])
    edit = False

    # рендер таблицы "Расчет углеродного сбора"
    data = make_table(num_rows=len(rows), num_cols=len(columns))
    headings = columns
    sg.set_options(dpi_awareness=True)
    sg.set_options(font=("Courier New", 14))
    layout = [[sg.Table(values=data, headings=headings, max_col_width=40,
                        auto_size_columns=True,
                        justification='right',
                        num_rows=len(rows)-1,
                        alternating_row_color=sg.theme_button_color()[1],
                        key='-TABLE-',
                        expand_x=True,
                        expand_y=True,
                        row_height=24,
                        enable_click_events=True,
                        )],
              [sg.Button('Данные')],
              [sg.Text('Cell clicked:'), sg.T(k='-CLICKED-')]]

    window = sg.Window('Table Element - Example 1', layout, resizable=True, finalize=True, auto_size_text=True)
    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        elif event == "Данные":
            window.close()
            ftt.building_final_table()
        elif isinstance(event, tuple):
            cell = row, col = event[2]
            window['-CLICKED-'].update(cell)
            edit_cell(window, '-TABLE-', row+1, col, justify='right')
            row_ch = (event[2][0])
            col_ch = (event[2][1])
            data[row_ch][col_ch] = text
            [data[row_ch][col_ch]]


main_example()
