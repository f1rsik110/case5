from openpyxl import load_workbook
from datetime import datetime
import PySimpleGUI as sg

layout = [[sg.Text('ФИО'), sg.Push(), sg.Input(key='FIO')],
          [sg.Text('Дата рождения'), sg.Push(), sg.Input(key='data')],
          [sg.Text('Номер телефона'), sg.Push(), sg.Input(key='telephone')],
          [sg.Text('Серия паспорта'), sg.Push(), sg.Input(key='seria')],
          [sg.Text('Номер паспорта'), sg.Push(), sg.Input(key='nomer')],
          [sg.Text('Адрес проживания'), sg.Push(), sg.Input(key='adres')],
          [sg.Button('Добавить'), sg.Button('Закрыть')]]
window = sg.Window('Регистрация стоматология', layout, element_justification='center')

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT or event == "Закрыть":
        break
    if event == 'Добавить':
        try:
            wb = load_workbook('БД.xlsx')
            sheet = wb['Лист1']
            ID = len(sheet['ID'])
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            data = [ID, values['FIO'], values['data'], values['telephone'], values['seria'], values['nomer'], values['adres'], time_stamp]
            sheet.append(data)
            wb.save('БД.xlsx')
            window['FIO'].update(value='')
            window['data'].update(value='')
            window['telephone'].update(value='')
            window['seria'].update(value='')
            window['nomer'].update(value='')
            window['adres'].update(value='')
            window['FIO'].set_focus()
            sg.popup('Данные сохранены')
        except PermissionError:
            sg.popup('File in use', 'File is being used by another User.\nPlease try again later.')
window.close()
