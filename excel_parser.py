import sys
import warnings
from datetime import datetime
from pathlib import Path
from typing import Any

import openpyxl
from loguru import logger

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def parse_active_sheet(file_path: Path) -> dict[int, Any]:
    """Парсит данные из листа "Активные" переданного Excel-файла.

    Считывает заголовки со второй строки и собирает значения по стандартизированному
    списку колонок. Ключом для каждой записи выступает целочисленное значение
    из колонки '# в33'. Колонка 'Серия-Номер' автоматически разделяется на два
    отдельных поля. Ожидается, что файл уже прошел валидацию и нормализацию.

    Args:
        file_path (Path): Путь к нормализованному целевому Excel-файлу (.xlsx).

    Returns:
        dict[int, Any]: Словарь с извлеченными данными. Ключи — ID из колонки '# в33' (int),
        значения — словари с данными конкретной строки по стандартным колонкам.
        По умолчанию всегда содержит стартовую запись {0: "null"}.
    """
    parsed_data: dict[int, Any] = {0: "null"}

    try:
        logger.trace(f"Парсинг данных из файла: {file_path!s}")
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb["Активные"]

        headers = {}
        for cell in sheet[2]:
            if cell.value:
                headers[str(cell.value).strip()] = cell.column - 1

        col_v33 = headers.get("# в33")

        if col_v33 is None:
            logger.critical('На второй строке не найден столбец "# в33".')
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
            "Код страны",
            "Код подразделения",
            "Уникальный идентификатор договора (сделки) БАНКА",
            "Дата согласия на обработку ПДН (дата договора)",
            "Статус долга (Операция)",
            "Дата создания (дата передачи цессии)",
            "Общая сумма долга",
            "Остаток долга",
            "Сумма последнего возрата",
            "Дата последнего возврата",
        ]

        def get_str_value(row: tuple[Any, ...], col_name: str) -> str:
            col_idx = headers.get(col_name)
            if col_idx is None or col_idx >= len(row) or row[col_idx] is None:
                return ""
            val = row[col_idx]
            if isinstance(val, datetime):
                return val.strftime("%d.%m.%Y")
            return str(val).strip()

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
                            f"Заполнено нулями."
                        )
                else:
                    logger.warning(f"Строка {row_idx}: Значение 'Серия-Номер' пустое. Заполнено нулями.")

                row_data["Серия"] = seria
                row_data["Номер"] = nomer

                parsed_data[key] = row_data
            except ValueError:
                logger.warning(f"Строка {row_idx}: Не удалось преобразовать '# в33' в число: {v33_val}. Пропущена.")

        wb.close()
        logger.success(f"Парсинг листа завершен.")
        logger.debug(f"Собрано записей: {len(parsed_data) - 1}")

        return parsed_data

    finally:
        try:
            file_path.unlink()
            logger.success(f"Временный файл - {file_path.name} - удален.")
        except Exception as e:
            logger.exception(f"Не удалось удалить временный файл {file_path.name}: {e}")
