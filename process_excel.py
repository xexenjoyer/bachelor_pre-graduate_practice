import pandas as pd
import json

def get_available_columns(file_path):
    df = pd.read_excel(file_path)
    columns = ', '.join(df.columns)
    return columns

def process_excel(file_path, selected_columns, merged_columns):
    # Чтение данных из Excel-файла
    df = pd.read_excel(file_path)

    # Запрос у пользователя выбранных столбцов
    selected_columns = selected_columns.split(',')
    for i in range(len(selected_columns)-1, -1, -1):
        selected_columns[i] = selected_columns[i].strip()
        if not selected_columns[i]:
            selected_columns.pop(i)

    merged_columns = merged_columns.split(',')
    for i in range(len(merged_columns)-1, -1, -1):
        merged_columns[i] = merged_columns[i].strip()
        if not merged_columns[i]:
            merged_columns.pop(i)

    not_merged_columns = list(set(selected_columns) - set(merged_columns))

    # Фильтрация данных по выбранным столбцам
    df = df[selected_columns]

    # Преобразование каждой строки в JSON-объект
    converted_rows = [False for i in range(len(df.index))]
    json_objects = []
    for i in range(len(df.index)):
        if not converted_rows[i]:
            converted_rows[i] = True
            for j in range(i, len(df.index)):
                if not converted_rows[j] and df.loc[i][not_merged_columns].equals(df.loc[j][not_merged_columns]):
                    converted_rows[j] = True
                    for col in merged_columns:
                        col_i = df.loc[i][col]
                        col_j = df.loc[j][col]
                        if type(col_i) is not list and col_i != col_j:
                            col_i = [col_i, col_j]
                        if type(col_i) is list and col_j not in col_i:
                            col_i.append(col_j)
                        df.loc[i][col] = col_i
            json_objects.append(df.loc[i].to_dict())

    # Запись JSON-объектов в текстовый файл
    with open('output.json', 'w') as file:
        json.dump(json_objects, file)