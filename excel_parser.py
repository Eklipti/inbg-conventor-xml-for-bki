import openpyxl
import logging
from pathlib import Path

def parse_active_sheet(file_path: Path, logger: logging.Logger) -> dict:
    logger.debug(f"Начинаем парсинг данных из файла: {file_path}")
    
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb["Активные"]
    
    parsed_data = {0: "null"}
    
    # Собираем все заголовки и их индексы колонок
    headers = {}
    for cell in sheet[2]:
        if cell.value:
            headers[str(cell.value).strip()] = cell.column - 1 
            
    col_v33 = headers.get("# в33")
    col_id = headers.get("ИД договора")
    
    if col_id is None:
        logger.critical("Критическая ошибка: на второй строке не найдены столбец 'ИД договора'.")
        return parsed_data
        
    logger.debug(f"Заголовки найдены. Начинаем чтение данных со строки 3...")
    
    for row in sheet.iter_rows(min_row=3, values_only=True):
        v33_val = row[col_v33]
        
        if v33_val is None:
            continue
            
        try:
            key = int(v33_val)
            
            # --- ИЗМЕНЕНИЯ ЗДЕСЬ ---
            # Теперь значением выступает вложенный словарь
            parsed_data[key] = {
                "ИД договора": str(row[col_id]) if row[col_id] is not None else "",
                # Место для будущих полей:
                # "Новое поле": row[headers.get("Название колонки")],
            }
            
        except ValueError:
            logger.warning(f"Не удалось преобразовать '# в33' в число: {v33_val}. Строка пропущена.")
            
    wb.close()
    logger.debug(f"Парсинг листа 'Активные' завершен. Собрано записей: {len(parsed_data) - 1}")
    
    return parsed_data