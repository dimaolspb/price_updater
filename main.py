import shutil
import datetime
import os

def copy_xlsx_file_with_timestamp(filename):
    # Получаем текущую дату и время
    current_datetime = datetime.datetime.now()

    # Генерируем строку с текущей датой и временем в формате ГГГГММДД_ЧЧММСС
    timestamp = current_datetime.strftime("%Y%m%d_%H%M%S")

    # Получаем расширение файла
    file_extension = os.path.splitext('test.xlsx')[1]

    # Создаем новое имя файла с добавлением текущей даты и времени
    new_filename = f"{os.path.splitext('test.xlsx')[0]}_{timestamp}{file_extension}"

    # Копируем файл с новым именем
    shutil.copy(filename, new_filename)

# Пример использования функции
copy_xlsx_file_with_timestamp("test.xlsx")