import openpyxl
import logging
import warnings
from datetime import datetime
from pathlib import Path
from utils import convert_xls_to_xlsx

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def validate_excel(file_path: Path, logger: logging.Logger) -> bool:
    """Проверка существования файла и наличия нужных листов."""
    if not file_path.exists():
        logger.error(f"Файл не найден: {file_path}")
        return False

    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        wb.close()
    except Exception as e:
        logger.critical(f"Не удалось прочитать файл. Убедитесь, что это корректный формат Excel. Ошибка: {e}")
        sys.exit(1)

    required_sheets = {"Активные"}
    current_sheets = set(sheet_names)
    
    if not required_sheets.issubset(current_sheets):
        logger.critical(f"В файле отсутствуют обязательные листы. Ожидалось: {required_sheets}, найдено: {current_sheets}")
        sys.exit(1)
        
    logger.info("Валидация файла прошла успешно.")
    return True

def parse_active_sheet(file_path: Path, logger: logging.Logger) -> dict:
    actual_file_path = file_path
    is_temp_file = False
    parsed_data = {0: "null"}

    if actual_file_path.suffix.lower() == '.xls':
        logger.warning("Обнаружен старый формат .xls. Конвертирую во временный xlsx.")
        temp_xlsx = convert_xls_to_xlsx(actual_file_path, logger)
        if not temp_xlsx:
            logger.critical("Отмена парсинга из-за ошибки смены формата.")
            sys.exit(1)
        actual_file_path = temp_xlsx
        is_temp_file = True

    try:
        if not validate_excel(actual_file_path, logger):
            return parsed_data

        logger.debug(f"Начинаем парсинг данных из файла: {str(actual_file_path)[-10:]}")
        wb = openpyxl.load_workbook(actual_file_path, data_only=True)
        sheet = wb["Активные"]
        
        headers = {}
        for cell in sheet[2]:
            if cell.value:
                headers[str(cell.value).strip()] = cell.column - 1 
        
        col_v33 = headers.get("# в33")
        
        if col_v33 is None:
            logger.critical("На второй строке не найден столбец '# в33'.")
            wb.close()
            sys.exit(1)

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
            if isinstance(val, datetime):
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
                    logger.warning(f"Строка {row_idx}: Значение 'Серия-Номер' пустое. Заполнено нулями.")
                
                row_data["Серия"] = seria
                row_data["Номер"] = nomer

                parsed_data[key] = row_data
            except ValueError:
                logger.warning(f"Строка {row_idx}: Не удалось преобразовать '# в33' в число: {v33_val}. Пропущена.")
                
        wb.close()
        logger.debug(f"Парсинг листа 'Активные' завершен. Собрано записей: {len(parsed_data) - 1}")
        
        return parsed_data
        
    finally:
        if is_temp_file and actual_file_path.exists():
            try:
                actual_file_path.unlink()
                logger.debug(f"Временный файл {actual_file_path.name} успешно удален.")
            except Exception as e:
                logger.warning(f"Не удалось удалить временный файл {actual_file_path.name}: {e}")