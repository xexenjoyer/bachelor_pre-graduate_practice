# -*- coding: utf-8 -*-
import PySimpleGUI as sg
import pandas as pd
from process_excel import process_excel, get_available_columns
import subprocess

class my_window(object):
    def __init__(self):
        self.NAME_OPEN = "Открыть"
        self.EVENT_OPEN = "file_path"
        self.EVENT_EXECUTE = "Выполнить"
        self.EVENT_COLUMNS = "columns"
        # Обозначение всех кнопок на интерфейсе
        self.layout = [[sg.FileBrowse(button_text=self.NAME_OPEN, k=self.EVENT_OPEN, enable_events=True,
                                 file_types=(("EXCEL Files", "*.xlsx"),))],
                  [sg.Text('Файл:', size=(16, 1)), sg.Text(k='selected')],
                  [sg.Text('Доступные столбцы:', size=(16, 1)), sg.Text(size=(105,3),k='available_columns',)],
                  #название столбцов можно было не только ввести в текстовое поле но и выбрать из списка
                  [sg.Text('Столбцы, которые нужно сохранить:'), sg.InputText(size=(45,2), key=self.EVENT_COLUMNS, do_not_clear=False), sg.Listbox([''], select_mode='extended',size=(45,3),k='OptionMenu')],
                  [sg.Button("Сохранить")],
                  [sg.Text('Столбцы, которые нужно свести в массив в случае похожих строк:'), sg.InputText(size=(45,2),key='columns_merged_to_list', do_not_clear=False), sg.Listbox([''], select_mode='extended', size=(45,3), k='OptionMenu2')],
                  [sg.Button("Сохранить!")],
                  #Фильтры
                  [sg.Text('Фильтры (Пример: A=ayz)'), sg.InputText(key='restrictions', do_not_clear=False)],
                  [sg.Button(self.EVENT_EXECUTE)],
                  [sg.Text('', size=(16, 1)), sg.Text(k='out')],

                  ]

    def run(self):
        helpstring1 = ""
        helpstring2 = ""
        window = sg.Window('Название', self.layout, grab_anywhere=True,
                       keep_on_top=True,
                       use_default_focus=False,
                       font='any 15',
                       resizable=False,
                           size=(1500,600)
                       )



        # Цикл событий
        while True:
            # Получение названия события и его значений
            event, values = window.read()

            # Закрытие окна
            if event == sg.WIN_CLOSED:
                break

            # Выбор пути
            if event == self.EVENT_OPEN:
                window['selected'].update(values[self.EVENT_OPEN])
                window['available_columns'].update(get_available_columns(values[self.EVENT_OPEN]))
                kostyl=(get_available_columns(values[self.EVENT_OPEN])).split(",")
                for i in range(len(kostyl)):
                    kostyl[i]=kostyl[i].lstrip()
                window['OptionMenu'].update(values=kostyl)
                window['OptionMenu2'].update(values=kostyl)

            if event == "Сохранить":
                helpstring1 = values["OptionMenu"][0]
                for i in range(len(values["OptionMenu"])-1):
                    helpstring1 += ", " + values["OptionMenu"][i+1]
                window['columns'].update(helpstring1)
                window['columns_merged_to_list'].update(helpstring2)

            if event == "Сохранить!":
                helpstring2 = values["OptionMenu2"][0]
                for i in range(len(values["OptionMenu2"]) - 1):
                    helpstring2 += ", " + values["OptionMenu2"][i + 1]
                window['columns'].update(helpstring1)
                window['columns_merged_to_list'].update(helpstring2)
            # Выполнение
            if event == self.EVENT_EXECUTE:
                if not values[self.EVENT_OPEN]:
                    window['out'].update("Не успешно")
                    helpstring1 = ""
                    helpstring2 = ""
                    continue

                process_excel(values[self.EVENT_OPEN], values[self.EVENT_COLUMNS], values['columns_merged_to_list'],values['restrictions'])





                helpstring1 = ""
                helpstring2 = ""

                window['out'].update("Успешно")



        window.close()
