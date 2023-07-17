import sys

import pandas as pd
from openpyxl import *
import shutil
import datetime
import os

''' наименования тканей переносим в буфер как есть, трансформируем (BLACK-OUT в б/о) при переносе в шаблон
цель буфера:

в нём хранятся данные, которые трудно извлекать из прайса, где они обозначены картинками,
поэтому в буфер добавляем данные в первый раз копированием, при этом создаём в исходном прайсе столбцы, в которых 
словами обозначаем смысл картинок.

При получении от поставщика нового прайса этот скрипт:
1. сравнивает позиции в прайсе и в буфере. Обновляет данные, но только строковые и числовые. Удаляет позиции, которые 
в новом прайсе больше не присутствуют. Добавляет новые, но заполняет только строковые и числовые данные, а те, что 
заполнены в прайсе картинками - нужно заполнять руками, иначе скрипт, переносящий потом данные в шаблон выдаст ошибку,
что не все данные заполнены'''

class FabricsPriceUpdate():
    def __init__(self):
        self.price_filename = 'price.xlsx'
        self.fabrics_sheet_name_in_price = 3
        self.buffer_filename = 'буфер.xlsx'
        self.fabrics_sheet_name_in_buffer = 'fabrics'

    def make_backup(self): # сделаем бэкап файла БУФЕР перед его изменением

        # Получаем текущую дату и время
        current_datetime = datetime.datetime.now()

        # Генерируем строку с текущей датой и временем в формате ГГГГММДД_ЧЧММСС
        timestamp = current_datetime.strftime("%Y%m%d_%H%M%S")

        # Получаем расширение файла
        file_extension = os.path.splitext(self.buffer_filename)[1]

        # Создаем новое имя файла с добавлением текущей даты и времени
        new_filename = f"{os.path.splitext(self.buffer_filename)[0]}_{timestamp}{file_extension}"

        # Копируем файл с новым именем
        shutil.copy(self.buffer_filename, new_filename)
        print('Бэкап файла БУФЕР создан')

    def actualize_fabrics_data_in_buffer(self):
        # определим номер начальной строки списка тканей в прайсе
        last_row_in_price = 142#int(input('Введите номер последней строки списка тканей: '))
        df_fabrics_in_price = pd.read_excel(self.price_filename, sheet_name=3, header=None, nrows=last_row_in_price + 1)
        cell_title_in_price = df_fabrics_in_price[df_fabrics_in_price == 'НАИМЕНОВАНИЕ'].stack().index[0]  # ищем индекс ячейки с "наименованием"
        start_row_in_price = cell_title_in_price[0] + 1
        #print(start_row_in_price, last_row_in_price)
        print(df_fabrics_in_price.iloc[start_row_in_price: last_row_in_price, [cell_title_in_price[1], 8, 9, 10, 11]])

        # определим номер начальной строки списка тканей в буфере

        df_fabrics_in_buffer = pd.read_excel(self.buffer_filename, self.fabrics_sheet_name_in_buffer, header=None)
        cell_title_in_buffer = df_fabrics_in_buffer[df_fabrics_in_buffer == 'НАИМЕНОВАНИЕ'].stack().index[0]
        start_row_in_buffer = cell_title_in_buffer[0] + 1
        #print('start_row_in_buffer', start_row_in_buffer)
        #print(df_fabrics_in_buffer.iloc[start_row_in_buffer:, [cell_title_in_buffer[1], 8, 9, 10, 11]])

        # Проверка, все ли наименования столбцов БУФЕРА присутствуют в ПРАЙСЕ
        print('Идёт проверка соответствия названий столбцов в прайсе и в буфере')
        buffer_titles = set(df_fabrics_in_buffer.loc[start_row_in_buffer - 1])
        price_titles = set(df_fabrics_in_price.loc[start_row_in_price - 1])
        print(buffer_titles)
        print(price_titles)
        result = df_fabrics_in_buffer.loc[start_row_in_buffer - 1].isin(df_fabrics_in_price.loc[start_row_in_price - 1]).all()
        if result:
            print('\n')
            print('Все заголовки колонок буфера содержатся в прайсе, можно обновлять список тканей')
            #input('Для продолжения нажмите Enter')
        else:
            print('\n', "Не все заголовки столбцов буфера содержатся в прайсе, приведите в соотвестсвие.",
                  "Несоответствия в БУФЕРЕ:",
                  sep='\n', end='\n')

            diff = set(buffer_titles).difference(price_titles)
            diff = list(diff)
            for i in range(len(diff)):
                print(diff[i])
            sys.exit()

        # найдём какие ткани в буфере нужно убрать и добавить
        fabrics_in_price = set(df_fabrics_in_price.iloc[start_row_in_price:last_row_in_price, cell_title_in_price[1]])
        fabrics_in_buffer = set(df_fabrics_in_buffer.iloc[start_row_in_buffer:, cell_title_in_buffer[1]])
        to_delete = fabrics_in_buffer - fabrics_in_price
        to_add = fabrics_in_price - fabrics_in_buffer
        print('удалить', to_delete)
        print('добавить', to_add)

fabrics_price_update = FabricsPriceUpdate()
fabrics_price_update.make_backup()
fabrics_price_update.actualize_fabrics_data_in_buffer()




# Читаем данные из файла Excel и выбираем нужный диапазон ячеек
#df = pd.read_excel(file_path, sheet_name=3, header=None, usecols='Q', nrows=142)

# Отображаем DataFrame
#print(df.iloc[start_row:end_row])


    # df = pd.read_excel(self.price_filename, sheet_name=3, header=None, nrows=142)
    # cell_title = df[df == 'НАИМЕНОВАНИЕ'].stack().index[0] # ищем индекс ячейки с "наименованием"
    # cell_vip_price = df[df == 'ВИП    15%'].stack().index[0]
    #
    # print(cell_vip_price)
    # print(df.loc[cell_title[0] + 1: 9, [cell_title[1], cell_vip_price[1]]])
