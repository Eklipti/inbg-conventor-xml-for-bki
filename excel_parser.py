import openpyxl
import logging
from pathlib import Path

def parse_active_sheet(file_path: Path, logger: logging.Logger) -> dict:
    logger.debug(f"Начинаем парсинг данных из файла: {file_path}")
    
    wb = openpyxl.load_workbook(file_path, data_only=True) # читает значения
    sheet = wb["Активные"]
    
    parsed_data = {0: "null"} # нулевой элемент для синхронизации
    
    headers = {}
    for cell in sheet[2]:
        if cell.value: # strip() убирает случайные пробелы по краям
            headers[str(cell.value).strip()] = cell.column - 1 
    col_v33 = headers.get("# в33")
    col_id = headers.get("ИД договора")
    
    if col_id is None:
        logger.critical("Критическая ошибка: на второй строке не найдены столбец 'ИД договора'.")
        return parsed_data
        
    logger.info(f"Заголовки найдены. Начинаем чтение данных со строки 3...")
    
    for row in sheet.iter_rows(min_row=3, values_only=True):
        v33_val = row[col_v33]
        id_val = row[col_id]
        
        # Если порядковый номер пуст, считаем, что таблица закончилась
        if v33_val is None:
            continue
            
        try:
            key = int(v33_val)
            value = str(id_val) if id_val is not None else ""
            
            parsed_data[key] = value
        except ValueError:
            logger.warning(f"Не удалось преобразовать '# в33' в число: {v33_val}. Строка пропущена.")
            
    wb.close()
    logger.debug(f"Парсинг листа 'Активные' завершен. Собрано записей: {len(parsed_data) - 1}")
    
    return parsed_data