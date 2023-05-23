import PySimpleGUI as sg
from process_excel import process_excel

class my_window(object):
    def __init__(self):
        self.NAME_OPEN = "Открыть"
        self.EVENT_OPEN = "file_path"
        self.EVENT_EXECUTE = "Выполнить"
        self.EVENT_COLUMNS = "columns"
        # Обозначение всех кнопок на интерфейсе
        self.layout = [[sg.FileBrowse(button_text=self.NAME_OPEN, k=self.EVENT_OPEN, enable_events=True,
                                 file_types=(("EXCEL Files", "*.xlsx"),))],
                  [sg.Text('', size=(16, 1)), sg.Text(k='selected')],
                  [sg.Button(self.EVENT_EXECUTE)],
                  [sg.Text('Введите имена столбцов через запятую:'), sg.InputText(key=self.EVENT_COLUMNS, do_not_clear=False)],
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

            # Выполнение
            if event == self.EVENT_EXECUTE:
                if not values[self.EVENT_OPEN]:
                    window['out'].update("Не успешно")
                    continue

                try:
                    process_excel(values[self.EVENT_OPEN], values[self.EVENT_COLUMNS])
                except:
                    continue

                window['out'].update("Успешно")

        window.close()