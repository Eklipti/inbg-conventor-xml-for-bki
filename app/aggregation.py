import json
import sys
from pathlib import Path

import openpyxl
from loguru import logger
from openpyxl.worksheet.worksheet import Worksheet

from app.utils import validate_excel


def process_excel_returns(returns_file_path: Path, main_file_path: Path) -> tuple[Path | None, int, int]:
    """Извлекает данные из файла возвратов, агрегирует их и обновляет основной файл.

    Читает файл возвратов (лист "взносы", любой регистр), ищет заголовки.
    Агрегирует суммы платежей по ключу (Номер ДО, Дата платежа).
    Затем открывает основной файл (лист "Активные"), ищет соответствующие
    записи по "Ключевое поле" (совпадает с "Номер ДО"). В найденную строку
    записывается дата, общая сумма возврата для группы и пересчитывается
    остаток долга (включая установку статуса "close" и обработку перевозврата).

    Args:
        returns_file_path (Path): Путь к файлу возвратов Excel (файл-донор).
        main_file_path (Path): Путь к основному файлу Excel (файл-реципиент).

    Returns:
        tuple[Path | None, int, int]: Кортеж, содержащий:
            - Путь к обработанному основному файлу (или None при ошибке).
            - Количество успешно обновленных записей (int).
            - Количество ненайденных записей (int).

    Raises:
        SystemExit: При возникновении критических ошибок в процессе обработки.
    """
    logger.debug(f'Слияние данных из "{returns_file_path.name}" в "{main_file_path.name}"')

    ret_path_obj = Path(returns_file_path)
    main_path_obj = Path(main_file_path)

    try:
        ret_wb = openpyxl.load_workbook(ret_path_obj, data_only=True)

        ret_sheet = None
        for sheet_name in ret_wb.sheetnames:
            if sheet_name.lower() == "взносы":
                ret_sheet = ret_wb[sheet_name]
                break

        if ret_sheet is None:
            logger.error('Лист "взносы" не найден в файле возвратов.')
            return None, 0, 0

        ret_target_headers: list[str] = [
            "Долговое обязательство.Номер ДО",
            "Поступление платежа.Сумма платежа",
            "Поступление платежа.Дата платежа",
        ]
        ret_headers_idx: dict[str, int] = {}

        for r_idx in [1]:
            for row in ret_sheet.iter_rows(min_row=r_idx, max_row=r_idx):
                for cell in row:
                    if cell.value and str(cell.value).strip() in ret_target_headers:
                        ret_headers_idx[str(cell.value).strip()] = cell.column - 1
            if len(ret_headers_idx) == len(ret_target_headers):
                break

        if len(ret_headers_idx) != len(ret_target_headers):
            missing = set(ret_target_headers) - set(ret_headers_idx.keys())
            logger.error(f"В файле возвратов не найдены заголовки: {missing}.")
            return None, 0, 0

        idx_ret_key: int = ret_headers_idx["Долговое обязательство.Номер ДО"]
        idx_ret_sum: int = ret_headers_idx["Поступление платежа.Сумма платежа"]
        idx_ret_date: int = ret_headers_idx["Поступление платежа.Дата платежа"]

        aggregated_sums: dict[tuple[str, str], float] = {}

        logger.info("Начало слияния возвратов.")
        for row_idx, row in enumerate(ret_sheet.iter_rows(min_row=3), start=3):
            cell_key = row[idx_ret_key].value
            cell_sum = row[idx_ret_sum].value
            cell_date = row[idx_ret_date].value

            key_val = str(cell_key).strip() if cell_key is not None else ""
            date_val = str(cell_date).strip() if cell_date is not None else ""
            sum_str = str(cell_sum).strip() if cell_sum is not None else "0"

            if not key_val:
                continue

            try:
                clean_sum = sum_str.replace(" ", "").replace(",", ".")
                sum_float = float(clean_sum)
            except ValueError:
                logger.error(f"Файл возвратов, строка {row_idx}: невозможно преобразовать '{sum_str}' в число.")
                return None, 0, 0

            dict_key = (key_val, date_val)
            if dict_key in aggregated_sums:
                aggregated_sums[dict_key] += sum_float
            else:
                aggregated_sums[dict_key] = sum_float

        logger.debug(f"Агрегация возвратов завершена. Уникальных записей: {len(aggregated_sums)}.")

        if not validate_excel(main_path_obj):
            logger.warning(f"Основной файл {main_path_obj.name} не прошел валидацию.")
            return None, 0, 0

        main_wb = openpyxl.load_workbook(main_path_obj)
        main_sheet = main_wb["Активные"]

        main_target_headers: list[str] = [
            "Ключевое поле",
            "Сумма последнего возрата",
            "Дата последнего возврата",
            "Остаток долга",
            "Статус долга (Операция)",
            "Перевозврат",
        ]
        main_headers_idx: dict[str, int] = {}

        for row in main_sheet.iter_rows(min_row=2, max_row=2):
            for cell in row:
                if cell.value and str(cell.value).strip() in main_target_headers:
                    main_headers_idx[str(cell.value).strip()] = cell.column - 1

        if len(main_headers_idx) != len(main_target_headers):
            missing = set(main_target_headers) - set(main_headers_idx.keys())
            logger.error(f"В основном файле не найдены заголовки: {missing}.")
            return None, 0, 0

        idx_main_key: int = main_headers_idx["Ключевое поле"]
        idx_main_sum: int = main_headers_idx["Сумма последнего возрата"]
        idx_main_date: int = main_headers_idx["Дата последнего возврата"]
        idx_main_debt: int = main_headers_idx["Остаток долга"]
        idx_main_status: int = main_headers_idx["Статус долга (Операция)"]
        idx_main_overpayment: int = main_headers_idx["Перевозврат"]

        error_sheet_name = "Не найдено в реестре"
        if error_sheet_name in main_wb.sheetnames:
            main_wb.remove(main_wb[error_sheet_name])
            logger.trace(f"Лист '{error_sheet_name}' очищен для новой итерации.")

        err_sheet = main_wb.create_sheet(error_sheet_name)
        err_sheet.append(["Источник", "Ключ", "Дата", "Сумма / Группа"])

        main_key_row_map: dict[str, int] = {}
        for row_idx, row in enumerate(main_sheet.iter_rows(min_row=3), start=3):
            cell_key = row[idx_main_key].value
            key_val = str(cell_key).strip() if cell_key is not None else ""
            if key_val and key_val not in main_key_row_map:
                main_key_row_map[key_val] = row_idx

        success_count = 0
        not_found_count = 0

        logger.info("Начало записи агрегированных сумм в основной файл.")
        for (key, date), total_sum in aggregated_sums.items():
            if key not in main_key_row_map:
                logger.warning(f"{key}: не найден.")
                err_sheet.append(["Возвраты", key, date, total_sum])
                not_found_count += 1
                continue

            success_count += 1
            row_idx = main_key_row_map[key]

            formatted_sum: str = f"{total_sum:.2f}".replace(".", ",")
            main_sheet.cell(row=row_idx, column=idx_main_sum + 1, value=formatted_sum)
            main_sheet.cell(row=row_idx, column=idx_main_date + 1, value=date)

            raw_debt = main_sheet.cell(row=row_idx, column=idx_main_debt + 1).value
            debt_str: str = str(raw_debt).strip() if raw_debt is not None else "0"

            try:
                clean_debt: str = debt_str.replace(" ", "").replace(",", ".")
                current_debt: float = float(clean_debt)
            except ValueError:
                logger.error(
                    f"Ошибка в строке {row_idx} основного файла: невозможно преобразовать остаток '{debt_str}'."
                )
                return None, 0, 0

            new_debt: float = current_debt - total_sum

            if new_debt == 0:
                logger.trace(f"{key}: Долг полностью погашен (в ноль).")
                main_sheet.cell(row=row_idx, column=idx_main_status + 1, value="close")
                formatted_new_debt = f"{new_debt:.2f}".replace(".", ",")
                main_sheet.cell(row=row_idx, column=idx_main_debt + 1, value=formatted_new_debt)

            elif new_debt < 0:
                logger.trace(f"{key}: Перевозврат - {new_debt:.2f}.")
                main_sheet.cell(row=row_idx, column=idx_main_status + 1, value="close")
                formatted_overpayment: str = f"{new_debt:.2f}".replace(".", ",")
                main_sheet.cell(row=row_idx, column=idx_main_overpayment + 1, value=formatted_overpayment)
                main_sheet.cell(row=row_idx, column=idx_main_debt + 1, value="0,00")

            else:
                logger.trace(f"{key}: Обновлен остаток долга - {current_debt:.2f} (было) / {new_debt:.2f} (стало).")
                formatted_new_debt = f"{new_debt:.2f}".replace(".", ",")
                main_sheet.cell(row=row_idx, column=idx_main_debt + 1, value=formatted_new_debt)

        logger.debug(f"Сохранение изменений в основной файл: {main_path_obj.name}")
        main_wb.save(main_path_obj)
        logger.success("Слияние возвратов успешно завершено.")

        return main_path_obj, success_count, not_found_count

    except Exception as e:
        logger.critical(f"Критический сбой при слиянии файлов: {e}")
        logger.exception("Стек вызовов:")
        sys.exit(1)


def process_other_closures(returns_file_path: Path, main_file_path: Path | None) -> tuple[Path | None, int, int]:
    """Обрабатывает иные закрытия на основе данных из специального листа.

    Ищет лист "закрытие иное" в файле возвратов. На первой строке находит
    столбцы "Объект" и "Группа ДО", собирает из них данные. Затем в основном
    файле (лист "Активные") обнуляет остаток долга для найденных объектов
    и устанавливает статус из dictionary.json согласно их группе.

    Args:
        returns_file_path (Path): Путь к файлу возвратов (источник).
        main_file_path (Path): Путь к основному файлу реестра (цель).

    Returns:
        tuple[Path | None, int, int]: Кортеж, содержащий:
            - Путь к измененному основному файлу (или None при ошибке).
            - Количество успешно обработанных закрытий (int).
            - Количество ненайденных закрытий (int).
    """
    logger.info('Запуск обработки "закрытие иное".')
    logger.debug(f"Для {returns_file_path}")

    if not main_file_path:
        logger.critical("Путь к основному файлу не передан (получен None).")
        return None, 0, 0

    try:
        project_root = Path(__file__).resolve().parent.parent
        dict_path = project_root / "dictionary.json"

        if not dict_path.exists():
            logger.critical(f"Файл словаря не найден в корне проекта: {dict_path}")
            return None, 0, 0

        with dict_path.open(encoding="utf-8") as f:
            status_mapping = json.load(f)

        ret_wb = openpyxl.load_workbook(returns_file_path, data_only=True)
        ret_sheet: Worksheet | None = None

        for sheet_name in ret_wb.sheetnames:
            if sheet_name.lower() == "закрытие иное":
                ret_sheet = ret_wb[sheet_name]
                break

        if not ret_sheet:
            logger.debug("Лист 'закрытие иное' отсутствует.")
            return None, 0, 0

        idx_ret_obj = None
        idx_ret_group = None

        first_row = next(ret_sheet.iter_rows(min_row=1, max_row=1))
        for cell in first_row:
            val = str(cell.value).strip() if cell.value else ""
            if val == "Объект":
                idx_ret_obj = cell.column
            elif val == "Группа ДО":
                idx_ret_group = cell.column

        if not idx_ret_obj or not idx_ret_group:
            logger.error("Заголовки 'Объект' или 'Группа ДО' не найдены на первой строке.")
            return None, 0, 0

        closures_to_process = {}
        for row_idx in range(2, ret_sheet.max_row + 1):
            obj_val = str(ret_sheet.cell(row_idx, idx_ret_obj).value or "").strip()
            group_val = str(ret_sheet.cell(row_idx, idx_ret_group).value or "").strip()

            if obj_val and obj_val != "None":
                closures_to_process[obj_val] = group_val

        if not closures_to_process:
            logger.warning("Данные для обработки не найдены под заголовками.")
            return None, 0, 0

        main_wb = openpyxl.load_workbook(main_file_path)

        error_sheet_name = "Не найдено в реестре"
        if error_sheet_name in main_wb.sheetnames:
            err_sheet = main_wb[error_sheet_name]
            logger.trace(f"Лист '{error_sheet_name}' найден, данные будут добавлены.")
        else:
            err_sheet = main_wb.create_sheet(error_sheet_name)
            err_sheet.append(["Источник", "Ключ", "Дата", "Сумма / Группа"])
            logger.debug(f"Лист '{error_sheet_name}' не найден, создан новый.")

        main_sheet = main_wb["Активные"]

        headers = {str(cell.value).strip(): cell.column for cell in main_sheet[2] if cell.value}
        idx_key = headers.get("Ключевое поле")
        idx_debt = headers.get("Остаток долга")
        idx_status = headers.get("Статус долга (Операция)")

        if not all([idx_key, idx_debt, idx_status]):
            logger.error("В основном файле отсутствуют необходимые столбцы.")
            return None, 0, 0

        processed_count = 0
        found_keys = set()
        for row_idx in range(3, main_sheet.max_row + 1):
            cell_val = str(main_sheet.cell(row_idx, idx_key).value or "").strip()

            if cell_val in closures_to_process:
                group_do = closures_to_process[cell_val]
                new_status = status_mapping.get(group_do, f"Неизвестный статус: {group_do}")

                main_sheet.cell(row_idx, idx_debt).value = "0,00"
                main_sheet.cell(row_idx, idx_status).value = new_status
                processed_count += 1
                found_keys.add(cell_val)
                logger.trace(f"{cell_val}: закрыт ({new_status})")

        for obj_val, group_val in closures_to_process.items():
            if obj_val not in found_keys:
                logger.warning(f"{obj_val}: не найден.")
                err_sheet.append(["Иные закрытия", obj_val, "-", group_val])

        if processed_count > 0 or len(found_keys) < len(closures_to_process):
            main_wb.save(main_file_path)
            not_found_count = len(closures_to_process) - processed_count
            logger.success("Обработка завершена.")
            return main_file_path, processed_count, not_found_count

        logger.warning("Совпадений по 'Ключевое поле' не найдено в реестре.")
        return None, 0, 0

    except Exception as e:
        logger.error(f"Критическая ошибка при обработке иных закрытий: {e}")
        logger.exception("Стек вызовов:")
        sys.exit(2)
