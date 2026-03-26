import json
import logging
import sys
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

logger = logging.getLogger(__name__)


def convert_xls_to_xlsx(xls_path: Path, logger: logging.Logger) -> Path:
    """Конвертирует устаревший формат .xls во временный файл .xlsx.

    Использует библиотеки pandas и xlrd для чтения данных и движок openpyxl
    для сохранения в новый формат. При критической ошибке прерывает выполнение программы.

    Args:
        xls_path (Path): Путь к исходному файлу .xls.
        logger (logging.Logger): Логгер для записи процесса и ошибок.

    Returns:
        Path: Путь к созданному временному файлу .xlsx.
    """
    xlsx_path = xls_path.with_name(f"{xls_path.stem}_temp.xlsx")
    logger.info(f"Конвертация старого формата {xls_path.name} в {xlsx_path.name}...")

    try:
        df_dict = pd.read_excel(xls_path, sheet_name=None, header=None, engine="xlrd")

        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            for sheet_name, df in df_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

        logger.info("Конвертация успешно завершена.")
        return xlsx_path
    except Exception as e:
        logger.critical(f"Ошибка при конвертации .xls в .xlsx: {e}")
        sys.exit(1)


def validate_config(config_data: dict, logger: logging.Logger) -> bool:
    """Проверяет корректность структуры конфигурационного словаря.

    Убеждается, что присутствуют обязательный блок "organization", все
    требуемые реквизиты (ИНН, ОГРН и т.д.), а также целочисленный
    счетчик запусков "run_counter".

    Args:
        config_data (dict): Словарь с конфигурационными данными.
        logger (logging.Logger): Логгер для вывода сообщений о недостающих полях.

    Returns:
        bool: True, если конфигурация валидна, иначе False.
    """
    if "organization" not in config_data:
        logger.critical("В конфигурации отсутствует обязательный блок 'organization'.")
        return False

    org = config_data["organization"]
    required_org_keys = ["inn", "ogrn", "fullName", "shortName", "otherName", "sourceDateStart"]
    for key in required_org_keys:
        if key not in org:
            logger.critical(f"В блоке 'organization' отсутствует обязательное поле: '{key}'.")
            return False

    if "run_counter" not in config_data or not isinstance(config_data["run_counter"], int):
        logger.critical("Параметр 'run_counter' отсутствует или не является целым числом (int).")
        return False

    return True


def load_config(config_path: Path, logger: logging.Logger) -> dict:
    """Загружает конфигурацию из JSON-файла и проводит её валидацию.

    В случае отсутствия файла, ошибок чтения или невалидной структуры
    выводит критическую ошибку в лог и прерывает выполнение программы.

    Args:
        config_path (Path): Путь к JSON-файлу конфигурации.
        logger (logging.Logger): Логгер приложения.

    Returns:
        dict: Словарь с загруженными конфигурационными данными.
    """
    if not config_path.exists():
        logger.critical(f"Конфигурационный файл не найден: {config_path}")
        sys.exit(1)
    try:
        with open(config_path, encoding="utf-8") as f:
            data: dict[str, Any] = json.load(f)
    except Exception as e:
        logger.critical(f"Ошибка при чтении {config_path}: {e}")
        sys.exit(1)

    if not validate_config(data, logger):
        sys.exit(1)

    return data


def format_date(date_str: str) -> str:
    """Преобразует строку с датой из формата ДД.ММ.ГГГГ в ГГГГ-ММ-ДД.

    Если на вход поступает уже отформатированная дата или пустая строка,
    возвращает её без изменений.

    Args:
        date_str (str): Исходная строка с датой.

    Returns:
        str: Отформатированная строка с датой (YYYY-MM-DD) или пустая строка.
    """
    if not date_str:
        return ""
    date_str = date_str.strip()
    if "." in date_str:
        parts = date_str.split(".")
        if len(parts) == 3:
            return f"{parts[2]}-{parts[1]}-{parts[0]}"
    return date_str


def clean_fio(text: str) -> str:
    """Очищает строку с ФИО от лишних символов и типичных заглушек.

    Убирает нижние подчеркивания, неразрывные пробелы и заменяет
    маркеры отсутствия данных ('NaN', 'None', 'нет', '-') на пустую строку.

    Args:
        text (str): Исходная строка с ФИО или её частью.

    Returns:
        str: Очищенная строка или пустая строка, если валидных данных нет.
    """
    if not text:
        return ""

    cleaned = str(text).replace("_", " ").replace("\xa0", " ").strip()

    if cleaned in ("-", ".", "None", "nan", "NaN", "нет"):
        return ""

    return cleaned


def clean_issuer(text: str) -> str:
    """Очищает текстовые поля (место рождения, кем выдан) и приводит их к стандарту.

    Переводит текст в верхний регистр, удаляет спецсимволы (*, <, >, кавычки),
    убирает лишние пробелы и обрезает результат до 200 символов.

    Args:
        text (str): Исходная строка текста.

    Returns:
        str: Очищенная и нормализованная строка.
    """
    if not text:
        return ""
    text = str(text).upper()  # Требуется верхний регистр для таких полей
    for char in ["*", "<", ">", "«", "»", '"']:
        text = text.replace(char, " ")
    text = " ".join(text.split())
    return text[:200]


def save_xml(root_element: ET.Element, filename: str, logger: logging.Logger):
    """Сохраняет XML-дерево в файл с форматированием (отступами) и декларацией.

    Args:
        root_element (ET.Element): Корневой элемент готового XML-дерева.
        filename (str): Имя (или путь) целевого файла для сохранения.
        logger (logging.Logger): Логгер для записи статуса операции.
    """
    tree = ET.ElementTree(root_element)
    ET.indent(tree, space="  ", level=0)

    try:
        tree.write(filename, encoding="UTF-8", xml_declaration=True)
        logger.debug(f"Файл успешно сформирован: {filename}")
    except Exception as e:
        logger.error(f"Ошибка при сохранении {filename}: {e}")


def calculate_days_difference(start_date_str: str, end_date_str: str) -> str:
    """Вычисляет количество дней между двумя датами.

    Поддерживает начальную дату в форматах ДД.ММ.ГГГГ или ГГГГ-ММ-ДД.
    Конечная дата ожидается в формате ГГГГ-ММ-ДД. Если разница отрицательная,
    возвращает '0'. Если дата некорректна — возвращает пустую строку.

    Args:
        start_date_str (str): Начальная дата (дата начала просрочки).
        end_date_str (str): Конечная дата (дата формирования документа).

    Returns:
        str: Строка с количеством дней (целое число >= 0) или пустая строка при ошибке.
    """
    if not start_date_str or not end_date_str:
        logger.warning(f"Дата отсутствует: {start_date_str}; {end_date_str}")
        return ""

    start_date_str = start_date_str.strip()
    end_date_str = end_date_str.strip()

    try:
        if "." in start_date_str:
            start_dt = datetime.strptime(start_date_str, "%d.%m.%Y")
        elif "-" in start_date_str:
            start_dt = datetime.strptime(start_date_str, "%Y-%m-%d")
        else:
            logger.critical(f"Неизвестный формат даты просрочки: '{start_date_str}'")
            return ""
        end_dt = datetime.strptime(end_date_str, "%Y-%m-%d")
        delta = (end_dt - start_dt).days
        return str(max(0, delta))

    except ValueError as e:
        logger.critical(f"Невозможно вычислить дни просрочки. Кривая дата: '{start_date_str}'. Ошибка: {e}")
        return ""


def format_sum(value) -> str:
    """Приводит денежную сумму к строковому формату с двумя знаками после запятой.

    Заменяет запятые на точки при парсинге. Если значение пустое или
    не может быть преобразовано в число, возвращает '0.00'.

    Args:
        value (Any): Исходное значение суммы (строка, число, None).

    Returns:
        str: Отформатированная строка суммы (например, '123.45', '0.00').
    """
    if value is None or str(value).strip() == "":
        return "0.00"
    try:
        clean_val = str(value).strip().replace(",", ".")
        float_val = float(clean_val)
        return f"{float_val:.2f}"
    except ValueError:
        return "0.00"


def save_config(config_data: dict, config_path: Path, logger: logging.Logger):
    """Сохраняет обновленный словарь конфигурации обратно в JSON-файл.

    Используется для сохранения инкрементированного счетчика запусков (run_counter)
    после успешной генерации всех XML-документов.

    Args:
        config_data (dict): Словарь с актуальными данными конфигурации.
        config_path (Path): Путь к целевому JSON-файлу.
        logger (logging.Logger): Логгер приложения.
    """
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config_data, f, ensure_ascii=False, indent=4)
        logger.debug("Конфигурационный файл успешно обновлен (счетчик увеличен).")
    except Exception as e:
        logger.error(f"Ошибка при перезаписи {config_path}: {e}")
