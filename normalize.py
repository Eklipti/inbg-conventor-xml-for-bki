from pathlib import Path

import openpyxl
from loguru import logger
from utils import convert_xls_to_xlsx, validate_excel

def process_excel_returns(file_path: str | Path) -> Path | None:
    """Обрабатывает Excel-файл, суммируя возвраты по ключевому полю и дате.

    Ищет заголовки, собирает данные, группирует суммы по ключу
    ("Ключевое поле", "Дата последнего возврата"). Оставляет первую найденную
    строку для каждой уникальной группы, суммирует в ней возвраты, а значения
    в ячейках дублирующих строк очищает (оставляет пустые строки).

    Args:
        file_path (str | Path): Путь к исходному файлу Excel.

    Returns:
        Path | None: Путь к обработанному файлу при успехе, иначе None.
    """
    path_obj = Path(file_path)

    try:
        if path_obj.suffix.lower() == ".xls":
            logger.info("Обнаружен формат .xls, инициирована конвертация.")
            # TODO: normalize.py уже делает проверку, должна идти работа уже с готовым временным/исходным файлом, который
            # TODO: будет предоставлен. excel_parser.py (28-45), conventor.py (624). excel_parser.py вызывать после в conventor.py.
            # TODO: Поэтому всё же нужно возвращать путь к изменённому файлу при успехи (временный/исходный), иначе None.
            path_obj = convert_xls_to_xlsx(path_obj)
        else:
            logger.debug(f"Файл имеет формат {path_obj.suffix}, конвертация не требуется.")

        logger.info("Запуск валидации Excel-файла.")
        if not validate_excel(path_obj):
            return None

        workbook = openpyxl.load_workbook(path_obj)
        sheet = workbook["Активные"]

        target_headers = ["Ключевое поле", "Сумма последнего возрата", "Дата последнего возврата"]
        header_indices: dict[str, int] = {}

        for row in sheet.iter_rows(min_row=2, max_row=2):
            for cell in row:
                if cell.value in target_headers:
                    header_indices[cell.value] = cell.column - 1

        if len(header_indices) != len(target_headers):
            logger.error(f"Не удалось найти все необходимые заголовки. Найдено: {list(header_indices.keys())}")
            return None

        idx_key = header_indices["Ключевое поле"]
        idx_sum = header_indices["Сумма последнего возрата"]
        idx_date = header_indices["Дата последнего возврата"]

        aggregated_sums: dict[tuple[str, str], float] = {}
        first_seen_row_idx: dict[tuple[str, str], int] = {}

        for row_idx, row in enumerate(sheet.iter_rows(min_row=3), start=3):
            cell_key = row[idx_key].value
            cell_sum = row[idx_sum].value
            cell_date = row[idx_date].value

            key_val = str(cell_key).strip() if cell_key is not None else ""
            date_val = str(cell_date).strip() if cell_date is not None else ""
            sum_str = str(cell_sum).strip() if cell_sum is not None else "0"

            if not key_val:
                continue

            try:
                clean_sum = sum_str.replace(" ", "").replace(",", ".")
                sum_float = float(clean_sum)
            except ValueError:
                logger.error(f"Невозможно преобразовать сумму '{sum_str}' в число.")
                return None

            dict_key = (key_val, date_val)
            if dict_key in aggregated_sums:
                aggregated_sums[dict_key] += sum_float
                for cell in row:
                    cell.value = None
            else:
                aggregated_sums[dict_key] = sum_float
                first_seen_row_idx[dict_key] = row_idx

        logger.info("Обновление сумм в первых найденных строках.")
        for (key, date), total_sum in aggregated_sums.items():
            row_idx = first_seen_row_idx[(key, date)]
            formatted_sum = f"{total_sum:.2f}".replace(".", ",")
            sheet.cell(row=row_idx, column=idx_sum + 1, value=formatted_sum)

        workbook.save(path_obj)
        logger.success(f"Обработка завершена. Результат сохранен в {path_obj}")

        return path_obj

    except Exception as e:
        logger.critical(f"Необработанное исключение в процессе выполнения: {e}")
        return None