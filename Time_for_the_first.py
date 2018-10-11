import pandas as pd
import numpy as np
# Название файла
name = 3
name_files = str(name) + ".xlsx"
# name_files = "1_1.xlsx"
# Загрузка файлов
data = pd.DataFrame(pd.read_excel(name_files, header=None))
ansi_table_before25 = pd.DataFrame(pd.read_excel("./tables/ansi_ss_before25.xlsx", header=None))
ansi_table_after25 = pd.DataFrame(pd.read_excel("./tables/ansi_ss_after25.xlsx", header=None))
ansi_table_ansi_cs = pd.DataFrame(pd.read_excel("./tables/ansi_cs.xlsx", header=None))
name_of_columns = pd.DataFrame(pd.read_excel("./tables/vars_for_rows.xlsx", header=None))
vars_material = pd.DataFrame(pd.read_excel("./tables/english_vars_material.xlsx", header=None))
def Clear_Spec(): #Отчистка от всего чего не должно быть
    global data
    data = data.dropna(thresh=5) #Удаление строк.
    data = data.dropna(axis = 1, thresh = (data.shape[1] // 5)) #Удаление столбцов.
    data.reset_index(inplace = True, drop = True) #Сбрасываем индекс.
    n = 10
    flag = False
    for Y in range(data.shape[0]): 
        count = 1
        for X in range(n):
            if flag:
                break
            if count >= 5:
                data = data.drop(list(data.index.values)[Y])
                flag = True
                break
            if data.iloc[Y, X] == count:
                count += 1
    data = data.T
    data.reset_index(inplace = True, drop = True) #Сбрасываем индекс.
    data = data.T
    
def Column_search(): # Поиск нужных характеристик материала
    global data
    global result
    flag_0 = False
    flag_1 = False
    flag_2 = False
    flag_3 = False
    pointer = 0
    data_for_parsing = data
    data_for_parsing = data_for_parsing.drop(index= 0)
    data_for_parsing.reset_index(inplace = True, drop = True) #Сбрасываем индекс.
    data.reset_index(inplace = True, drop = True) #Сбрасываем индекс.
    result = pd.DataFrame(np.zeros((data_for_parsing.shape[0]-1,  ), dtype=[
    ('Наименование','a4'),
    ('Диаметр','f4'),
    ('Давление','f4'),
    ('Температура','f4'),
    ('Корпус','a4'),
    ('Присоединение','a4')
    ]))
    headers = ['Наименование','Диаметр','Давление','Температура','Корпус','Присоединение']
    # Пробег по клиентской таблице
    for Y in list(data.columns.values):
        if flag_0 != True:
            #Пробег по таблице наименований
            for B in list(name_of_columns.columns.values):
                if flag_2 != True:
                    for A in list(name_of_columns.index.values):
                        if flag_3 != True:
                            if (str(data.iloc[0,Y]).lower() == str(name_of_columns.iloc[A,B]).lower()):
                                data.to_excel("data.xlsx")
                                data_for_parsing.to_excel("data_for_parsing.xlsx")
                                for G in (data_for_parsing.index.values):
                                    data_column = data_for_parsing[Y]
                                    result[headers[B]] = data_column.iloc[0:data_column.shape[0]]
                                pointer += 1
                                if pointer == 6:
                                    flag_3 = True
                                    break
                        else:
                            flag_2 = True
                            break
                else:
                    flag_1 = True
                    break
        else:
            flag_0 = True
            break
    return(result)

def ANSI_column(result): # Создание колонки ANSI
    global vars_material
    vars_material = list(vars_material[0])
    header_ANSI = ['ANSI']
    ANSI = [[]]
    for main_x in range (0, result.shape[0]):
        stop_point = 1
        #Выбор талаблицы расчета ANSI
        ss = False
        for vars_material_key in range (0, len(vars_material)):
            if (result.iloc[(main_x, 4)]) == vars_material[vars_material_key]:
                ss = True
                if result.iloc[(main_x, 1)]  <= 25:
                    ansi_table = ansi_table_before25
                else:
                    ansi_table = ansi_table_after25
        
        if  ss == False:
            ansi_table = ansi_table_ansi_cs
        #Выбор коэффициента ANSI
        for k in range (1, ansi_table.shape[0]):
            if result.iloc[(main_x, 3)] <= ansi_table.iloc[(k,0)] and stop_point == 1:
                for j in range (1,ansi_table.shape[1]):
                    if result.iloc[(main_x, 2)] <= ansi_table.iloc[(k,j)] and stop_point == 1:
                        ANSI[0].append(ansi_table.iloc[(0,j)])
                        stop_point = 2
                        break
    error = "[[]]'"
    ANSI_list = ANSI
    for char in error:
        ANSI_list = str(ANSI_list).replace(char,"")
    ANSI_list = ANSI_list.split(',', maxsplit=-1)
    result['ANSI'] = ANSI_list
    return(result)

#вызов функций
Clear_Spec()
Column_search() 
ANSI_column(result)
name_files_w = name_files.split('.', maxsplit=1)
result.to_excel("result/" + str(name_files_w[0]) + "writer.xlsx")