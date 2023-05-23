import pandas as pd
import json

def process_excel(file_path, selected_columns):
    # Чтение данных из Excel-файла
    df = pd.read_excel(file_path)

    # Запрос у пользователя выбранных столбцов
    selected_columns = selected_columns.split(',')
    for i in range(len(selected_columns)-1, -1, -1):
        selected_columns[i] = selected_columns[i].strip()
        if not selected_columns[i]:
            selected_columns.pop(i)

    # Фильтрация данных по выбранным столбцам
    df_selected = df[selected_columns]

    # Преобразование каждой строки в JSON-объект
    json_objects = []
    for _, row in df_selected.iterrows():
        json_objects.append(row.to_dict())

    # Запись JSON-объектов в текстовый файл
    with open('output.json', 'w') as file:
        json.dump(json_objects, file)