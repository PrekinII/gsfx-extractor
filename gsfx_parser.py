#import os
import zipfile
#from collections import defaultdict

#import xml.etree.ElementTree as ET
#import pandas as pd
#import openpyxl
#from openpyxl.styles import Font, PatternFill, Alignment
#from openpyxl.utils import get_column_letter


def extract_total_from_gsfx(gsfx_path):
    """
    Открывает формат ГрандСметы как zip, читает файл Properties.txt и извлекает Суммму по смете Total=
    """
    try:
        with zipfile.ZipFile(gsfx_path, 'r') as archive:
            if 'Properties.txt' not in archive.namelist():
                return None
            with archive.open('Properties.txt') as prop_file:
                for line in prop_file:
                    line = line.decode('windows-1251').strip()
                    if line.startswith('Total='):
                        return float(line.split("=", 1)[1])
    except Exception as e:
        print(f'Ошибка в файле {gsfx_path}: {e}')
    return None

# def process_folder_structure(base_path):
#     """
#     Проходит по папкам и подпапкам, исходя из структуры забирает наименование объектов, подрядчиков,
#     обрабатывает только gsfx файлы, игнорирует папку _Архив, содержащую неактуальные файлы.
#     """
#     smeta_data = [] #Список всех смет
#     contractor_totals = defaultdict(float) #defaultdict автоматически создает значение по умолчанию, если его не было.
#     object_totals = defaultdict(float)

#     #Проходим по объектам
#     for object_name in os.listdir(base_path):
#         object_path = os.path.join(base_path, object_name)
#         if not os.path.isdir(object_path):
#             continue
        
#         #Проходим по подрядчикам
#         for contractor_name in os.listdir(object_path):
#             contractor_path = os.path.join(object_path, contractor_name)
#             if not os.path.isdir(contractor_path) or contractor_name.lower() == '_архив':
#                 continue

#             #Обарбатываем gsfx файлы Гранда
#             for file in os.listdir(contractor_path):
#                 if file.lower().endswith('.gsfx'):
#                     file_path = os.path.join(contractor_path, file)
#                     total = extract_total_from_gsfx(file_path)

#                     if total is not None:
#                         smeta_name = os.path.splitext(file)[0]
#                         smeta_data.append({
#                             'Объект': object_name,
#                             'Подрядчик': contractor_name,
#                             'Смета': smeta_name,
#                             'Стоимость по смете': total
#                         })

#                         #Накопление итогов
#                         contractor_totals[(object_name, contractor_name)] += total
#                         object_totals[object_name] += total

#     return smeta_data, contractor_totals, object_totals


# #Сохраняем в excel
# def save_to_excel_form(smeta_data, contractor_totals, object_totals, output_path):
#     """
#     Создает файл excel  с термя листами Сметы, Итоги по подрядчикам, Итоги по объктам
#     """
#     df_smety = pd.DataFrame(smeta_data)
#     df_contractors = pd.DataFrame([
#         {'Объект': obj, 'Подрядчик': contr, 'Итого по подрядчику': total}
#         for (obj,contr), total in contractor_totals.items()
#     ])
#     df_objects = pd.DataFrame([
#         {'Объект': obj, 'Итого по объекту': total}
#         for obj, total in object_totals.items()
#     ])

#     #Создаем Excel-файл с несколькими листами
#     with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
#         df_smety.to_excel(writer, sheet_name='Сметы', index=False)
#         df_contractors.to_excel(writer, sheet_name='Итоги по подрядчикам', index=False)
#         df_objects.to_excel(writer, sheet_name='Итоги по объектам', index=False)

#     #Загружаем workbook чтобы навести красоту
#     wb = openpyxl.load_workbook(output_path)

#     #Цвета и стили
#     header_fill = PatternFill(start_color='B0C4DE', end_color='B0C4DE', fill_type='solid')
#     header_font = Font(bold=True)
#     money_align = Alignment(horizontal='right')

#     for sheet_name in wb.sheetnames:
#         ws = wb[sheet_name]

#         #Подгон ширины столбцов и формление заголовков
#         for col_idx, col in enumerate(ws.columns, 1):
#             max_length = 0
#             for cell in col:
#                 cell_value = str(cell.value) if cell.value else ""
#                 max_length = max(max_length, len(cell_value))
#                 #Оформляем заголовки
#                 if cell.row == 1:
#                     cell.fill = header_fill
#                     cell.font = header_font
#                     cell.alignment = Alignment(horizontal='center')
#                 #Выравниваем числа справа
#                 if isinstance(cell.value, (int, float)):
#                     cell.alignment = money_align
#                     cell.number_format = '#,##0.00'
#             adjusted_width = max_length + 4
#             column_letter = get_column_letter(col_idx)
#             ws.column_dimensions[column_letter].width = adjusted_width

#     wb.save(output_path)
#     print(f'\n✅ Excel-файл успешно создан: {output_path}')

# if __name__ == "__main__":
#     #Путь к папке с объектами
#     base_folder = r"C:\Users\Пользователь\Desktop\Тест\Объекты"
#     #Выходной файл
#     output_excel = r"C:\Users\Пользователь\Desktop\Тест\Сводная\Сводная по объектам.xlsx"
#     #Обработка и сбор данных
#     smeta_data, contractor_totals, object_totals = process_folder_structure(base_folder)
#     #Оформляем и сохраняем
#     save_to_excel_form(smeta_data, contractor_totals, object_totals, output_excel)