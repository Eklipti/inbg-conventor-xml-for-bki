import openpyxl
import logging
from pathlib import Path
import datetime

def parse_active_sheet(file_path: Path, logger: logging.Logger) -> dict:
    logger.debug(f"Начинаем парсинг данных из файла: {file_path}")
    
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb["Активные"]
    
    parsed_data = {0: "null"}
    
    # Собираем заголовки
    headers = {}
    for cell in sheet[2]:
        if cell.value:
            headers[str(cell.value).strip()] = cell.column - 1 
            
    col_v33 = headers.get("# в33")
    
    if col_v33 is None:
        logger.error("Критическая ошибка: на второй строке не найден столбец '# в33'.")
        return parsed_data

    # Список стандартных колонок формата
    standard_columns = [
        "ИД договора", "Фамилия", "Имя", "Отчество", "Дата рождения", 
        "Место рождения", "Дата выдачи", "Кем выдан", "Адрес регистрации", 
        "Адрес фактический", "Код страны", "Код региона", "Почтовый индекс", 
        "Населённый пункт", "Улица", "Номер дома", "Номер квартиры", 
        "Уникальный идентификатор договора (сделки) БАНКА", 
        "Дата согласия на обработку ПДН (дата договора)", 
        "Статус долга (Операция)", "Группа долга", 
        "Дата создания (дата передачи цессии)", "Дата передачи (обновления)", 
        "Общая сумма долга", "Сумма возвратов", "Сумма выплаченная", 
        "Остаток долга", "Сумма последнего возрата", "Дата последнего возврата", 
        "Цессия"
    ]

    def get_str_value(row, col_name):
        col_idx = headers.get(col_name)
        if col_idx is None or col_idx >= len(row) or row[col_idx] is None:
            return ""
        
        val = row[col_idx]
        # Если openpyxl распознал дату
        if isinstance(val, datetime.datetime):
            return val.strftime("%d.%m.%Y")
            
        return str(val).strip()

    logger.info(f"Заголовки найдены.")
    
    for row_idx, row in enumerate(sheet.iter_rows(min_row=3, values_only=True), start=3):
        v33_val = row[col_v33]
        
        if v33_val is None:
            continue
            
        try:
            key = int(v33_val)
            
            row_data = {}
            for col_name in standard_columns:
                row_data[col_name] = get_str_value(row, col_name)
            
            seria, nomer = "0000", "000000"
            sn_raw = get_str_value(row, "Серия-Номер")
            if sn_raw:
                if "-" in sn_raw:
                    parts = sn_raw.split("-", 1)
                    seria = parts[0].strip()
                    nomer = parts[1].strip()
                else:
                    logger.warning(f"Строка {row_idx}: Значение 'Серия-Номер' ({sn_raw}) не содержит дефис. Поля оставлены пустыми.")
            else:
                logger.debug(f"Строка {row_idx}: Значение 'Серия-Номер' пустое. Заполнено нулями.")
            
            row_data["Серия"] = seria
            row_data["Номер"] = nomer

            parsed_data[key] = row_data
            
        except ValueError:
            logger.warning(f"Строка {row_idx}: Не удалось преобразовать '# в33' в число: {v33_val}. Пропущена.")
            
    wb.close()
    logger.debug(f"Парсинг листа 'Активные' завершен. Собрано записей: {len(parsed_data) - 1}")
    
    return parsed_data