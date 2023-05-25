import PySimpleGUI as sg
from process_excel import process_excel, get_available_columns

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
                  [sg.Text('Доступные столбцы:', size=(16, 1)), sg.Text(k='available_columns')],
                  [sg.Text('Столбцы, которые нужно сохранить:'), sg.InputText(key=self.EVENT_COLUMNS, do_not_clear=False)],
                  [sg.Text('Столбцы, которые нужно свести в массив в случае похожих строк:'), sg.InputText(key='columns_merged_to_list', do_not_clear=False)],
                  [sg.Button(self.EVENT_EXECUTE)],
                  [sg.Text('', size=(16, 1)), sg.Text(k='out')],
                  ]

    def run(self):
        window = sg.Window('Название', self.layout)
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


            # Выполнение
            if event == self.EVENT_EXECUTE:
                if not values[self.EVENT_OPEN]:
                    window['out'].update("Не успешно")
                    continue

                process_excel(values[self.EVENT_OPEN], values[self.EVENT_COLUMNS], values['columns_merged_to_list'])

                window['out'].update("Успешно")

        window.close()