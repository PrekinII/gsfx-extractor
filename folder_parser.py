import os
from collections import defaultdict
from gsfx_parser import extract_total_from_gsfx


def process_folder_structure(base_path):
    """
    Проходит по папкам и подпапкам, исходя из структуры забирает наименование объектов, подрядчиков,
    обрабатывает только gsfx файлы, игнорирует папку _Архив, содержащую неактуальные файлы.
    """
    smeta_data = [] #Список всех смет
    contractor_totals = defaultdict(float) #defaultdict автоматически создает значение по умолчанию, если его не было.
    object_totals = defaultdict(float)

    #Проходим по объектам
    for object_name in os.listdir(base_path):
        object_path = os.path.join(base_path, object_name)
        if not os.path.isdir(object_path):
            continue
        
        #Проходим по подрядчикам
        for contractor_name in os.listdir(object_path):
            contractor_path = os.path.join(object_path, contractor_name)
            if not os.path.isdir(contractor_path) or contractor_name.lower() == '_архив':
                continue

            #Обарбатываем gsfx файлы Гранда
            for file in os.listdir(contractor_path):
                if file.lower().endswith('.gsfx'):
                    file_path = os.path.join(contractor_path, file)
                    total = extract_total_from_gsfx(file_path)

                    if total is not None:
                        smeta_name = os.path.splitext(file)[0]
                        smeta_data.append({
                            'Объект': object_name,
                            'Подрядчик': contractor_name,
                            'Смета': smeta_name,
                            'Стоимость по смете': total
                        })

                        #Накопление итогов
                        contractor_totals[(object_name, contractor_name)] += total
                        object_totals[object_name] += total

    return smeta_data, contractor_totals, object_totals