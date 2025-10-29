from folder_parser import process_folder_structure
from excel_writer import save_to_excel_form
from settings import INPUT_PATH, OUTPUT_PATH


if __name__ == "__main__":
    #Обработка и сбор данных
    smeta_data, contractor_totals, object_totals = process_folder_structure(INPUT_PATH)
    #Оформляем и сохраняем
    save_to_excel_form(smeta_data, contractor_totals, object_totals, OUTPUT_PATH)
    print(f'\n✅ Excel-файл успешно создан: {OUTPUT_PATH}')