# -*- coding: utf-8 -*-
import pandas as pd
import json
import webbrowser
def get_available_columns(file_path):
    df = pd.read_excel(file_path)
    columns = ', '.join(df.columns)
    return columns

def process_excel(file_path, selected_columns, merged_columns,restrictions):
    # Чтение данных из Excel-файла
    df = pd.read_excel(file_path)
    print(df)

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

    #Фильтры
    restrictions = restrictions.split(',')
    for i in range(len(restrictions)-1,-1,-1):
        restrictions[i]=restrictions[i].strip().split("=")
        if not restrictions[i]:
            restrictions.pop(i)

    not_merged_columns = list(set(selected_columns) - set(merged_columns))

    # Фильтрация данных по выбранным столбцам
    print(list(df))
    df = df[selected_columns]
    print(df)



    # Фильтрация данных по содержанию строк
    if len(restrictions[0][0])!=0:
        for i in range(len(restrictions)):
            print(restrictions[i][1])
            col = restrictions[i][0]
            df[col] = df[col].astype(str)
            df = df[df[col] == restrictions[i][1]]





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
                        flag = 0
                        col_i = df.loc[i][col]

                        col_j = df.loc[j][col]
                        if type(col_i) is not list and col_i != col_j:
                            col_i = [col_i, col_j]
                        if type(col_i) is list and col_j not in col_i:
                            col_i.append(col_j)
                        df.loc[i][col] = col_i
            json_objects.append(df.loc[i].to_dict())

    # Запись JSON-объектов в текстовый файл
    with open('C:/Users/KolFi/Desktop/bachelor_pre-graduate_practice-master/output/output.json', 'w') as file:
        # исправить работу с кириллицей
        json.dump(json_objects, file,ensure_ascii=False)

        df = pd.read_json(r'output.json', encoding='cp1251')
        df.to_csv(r'C:\Users\KolFi\Desktop\bachelor_pre-graduate_practice-master\rdy_2_read1.txt', index=False)

        # открыть папку с json'чиком
        webbrowser.open('file:///C:/Users/KolFi/Desktop/bachelor_pre-graduate_practice-master/output')
