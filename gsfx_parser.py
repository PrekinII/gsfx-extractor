import zipfile


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

