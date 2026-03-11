import sys
import json
import logging
import xml.etree.ElementTree as ET
import pandas as pd
from pathlib import Path
from datetime import datetime

logger = logging.getLogger(__name__)

def convert_xls_to_xlsx(xls_path: Path, logger: logging.Logger) -> Path:
    """Конвертирует формат .xls во временный .xlsx с помощью pandas и xlrd."""
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
    """Проверяет структуру конфигурационного файла на соответствие требованиям."""
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
    """Загрузка конфигурации из JSON файла с валидацией."""
    if not config_path.exists():
        logger.critical(f"Конфигурационный файл не найден: {config_path}")
        sys.exit(1)
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        logger.critical(f"Ошибка при чтении {config_path}: {e}")
        sys.exit(1)
        
    # Валидируем структуру. Если кривая - завершаем работу
    if not validate_config(data, logger):
        sys.exit(1)
        
    return data

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
    """Очищает ФИО: убирает мусор и типичные заглушки."""
    if not text:
        return ""
        
    cleaned = str(text).replace("_", " ").replace("\xa0", " ").strip()
    
    if cleaned in ("-", ".", "None", "nan", "NaN", "нет"):
        return ""
        
    return cleaned

def clean_issuer(text: str) -> str:
    """Очищает текстовые поля (Кем выдан, Место рождения) от мусора."""
    if not text:
        return ""
    text = str(text).upper() # Требуется верхний регистр для таких полей
    for char in ["*", "<", ">", "«", "»", '"']:
        text = text.replace(char, " ")
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

def save_config(config_data: dict, config_path: Path, logger: logging.Logger):
    """Сохраняет обновленную конфигурацию обратно в JSON файл."""
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, ensure_ascii=False, indent=4)
        logger.debug("Конфигурационный файл успешно обновлен (счетчик увеличен).")
    except Exception as e:
        logger.error(f"Ошибка при перезаписи {config_path}: {e}")