import logging
import sys
import warnings
from datetime import datetime
from pathlib import Path
from typing import Any

import openpyxl

from utils import convert_xls_to_xlsx

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def validate_excel(file_path: Path, logger: logging.Logger) -> bool:
    """Проверяет существование Excel-файла и наличие в нем обязательных листов.

    Функция убеждается, что файл доступен по указанному пути, может быть
    успешно прочитан и содержит лист с именем "Активные". При критических
    ошибках чтения (например, битый файл) выполнение программы прерывается.

    Args:
        file_path (Path): Путь к проверяемому файлу Excel.
        logger (logging.Logger): Настроенный логгер для вывода статуса и ошибок.

    Returns:
        bool: True, если файл валиден и содержит необходимые листы, иначе False.
    """
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
        logger.critical(
            f"В файле отсутствуют обязательные листы. Ожидалось: {required_sheets}, найдено: {current_sheets}"
        )
        sys.exit(1)

    logger.info("Валидация файла прошла успешно.")
    return True


def parse_active_sheet(file_path: Path, logger: logging.Logger) -> dict:
    """Парсит данные из листа "Активные" переданного Excel-файла.

    Считывает заголовки со второй строки и собирает значения по стандартизированному
    списку колонок. Ключом для каждой записи выступает целочисленное значение
    из колонки '# в33'. Колонка 'Серия-Номер' автоматически разделяется на два
    отдельных поля. Если на вход поступает файл устаревшего формата (.xls),
    автоматически создается и используется временный файл .xlsx.

    Args:
        file_path (Path): Путь к целевому Excel-файлу (.xls или .xlsx).
        logger (logging.Logger): Настроенный логгер для записи процесса парсинга.

    Returns:
        dict: Словарь с извлеченными данными. Ключи — ID из колонки '# в33' (int),
        значения — словари с данными конкретной строки по стандартным колонкам.
        По умолчанию всегда содержит стартовую запись {0: "null"}.
    """
    actual_file_path = file_path
    is_temp_file = False
    parsed_data: dict[int, Any] = {0: "null"}

    if actual_file_path.suffix.lower() == ".xls":
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
            "Фамилия",
            "Имя",
            "Отчество",
            "Дата рождения",
            "Место рождения",
            "Дата выдачи",
            "Кем выдан",
            "Уникальный идентификатор договора (сделки) БАНКА",
            "Дата согласия на обработку ПДН (дата договора)",
            "Статус долга (Операция)",
            "Дата создания (дата передачи цессии)",
            "Общая сумма долга",
            "Остаток долга",
            "Сумма последнего возрата",
            "Дата последнего возврата",
        ]

        def get_str_value(row, col_name):
            col_idx = headers.get(col_name)
            if col_idx is None or col_idx >= len(row) or row[col_idx] is None:
                return ""
            val = row[col_idx]
            if isinstance(val, datetime):
                return val.strftime("%d.%m.%Y")
            return str(val).strip()

        logger.info("Заголовки найдены.")

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
                        logger.warning(
                            f"Строка {row_idx}: Значение 'Серия-Номер' ({sn_raw}) не содержит дефис. "
                            f"Поля оставлены пустыми."
                        )
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
