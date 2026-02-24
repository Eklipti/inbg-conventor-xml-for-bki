import sys
import json
import logging
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime

logger = logging.getLogger(__name__)

def load_config(config_path: Path, logger: logging.Logger) -> dict:
    """Загрузка конфигурации из JSON файла."""
    if not config_path.exists():
        logger.critical(f"Конфигурационный файл не найден: {config_path}")
        sys.exit(1)
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.critical(f"Ошибка при чтении {config_path}: {e}")
        sys.exit(1)

def format_date(date_str: str) -> str:
    """Переводит дату из ДД.ММ.ГГГГ в ГГГГ-ММ-ДД."""
    if not date_str:
        return ""
    date_str = date_str.strip()
    if "." in date_str:
        parts = date_str.split(".")
        if len(parts) == 3:
            return f"{parts[2]}-{parts[1]}-{parts[0]}"
    return date_str

def clean_fio(text: str) -> str:
    """Очищает ФИО: заменяет '_' и неразрывные пробелы на обычные, убирает пробелы по краям."""
    if not text:
        return ""
    # \xa0 - неразрывный пробел (NBSP)
    return str(text).replace("_", " ").replace("\xa0", " ").strip()

def clean_issuer(text: str) -> str:
    """Очищает поле 'Кем выдан' от запрещенных символов и обрезает до 200 символов."""
    if not text:
        return ""
    text = str(text)
    for char in ["*", "<", ">", "«", "»"]:
        text = text.replace(char, " ")
    # Заменяем случайно образовавшиеся двойные пробелы на одинарные и обрезаем
    text = " ".join(text.split())
    return text[:200]

def save_xml(root_element: ET.Element, filename: str, logger: logging.Logger):
    """Сохраняет XML дерево в файл с нужным заголовком."""
    tree = ET.ElementTree(root_element)
    ET.indent(tree, space="  ", level=0) 
    
    try:
        tree.write(filename, encoding="UTF-8", xml_declaration=True)
        logger.debug(f"Файл успешно сформирован: {filename}")
    except Exception as e:
        logger.error(f"Ошибка при сохранении {filename}: {e}")

def calculate_days_difference(start_date_str: str, end_date_str: str) -> str:
    """Вычисляет количество дней между двумя датами. Ожидает формат ДД.ММ.ГГГГ или ГГГГ-ММ-ДД."""
    if not start_date_str or not end_date_str:
        logger.warning(f"Дата отсутствует: {start_date_str}; {end_date_str}")
        return ""
    
    start_date_str = start_date_str.strip()
    end_date_str = end_date_str.strip()
        
    # Возвращаем пустую строку, чтобы сломать валидацию
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
    """Округляет сумму до 2 знаков после запятой и возвращает строку (например, '0.00', '123.45')."""
    if value is None or str(value).strip() == "":
        return "0.00"
    try:
        clean_val = str(value).strip().replace(',', '.')
        float_val = float(clean_val)
        return f"{float_val:.2f}"
    except ValueError:
        return "0.00"
