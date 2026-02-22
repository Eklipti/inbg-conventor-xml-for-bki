import logging
import json
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime

import excel_parser

def load_config(config_path: Path, logger: logging.Logger) -> dict:
    """Загрузка конфигурации из JSON файла."""
    if not config_path.exists():
        logger.crtitical(f"Конфигурационный файл не найден: {config_path}")
        sys.exit(1)
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.crtitical(f"Ошибка при чтении {config_path}: {e}")
        sys.exit(1)

def save_xml(root_element: ET.Element, filename: str, logger: logging.Logger):
    """Сохраняет XML дерево в файл с нужным заголовком."""
    tree = ET.ElementTree(root_element)
    ET.indent(tree, space="  ", level=0) 
    
    try:
        tree.write(filename, encoding="UTF-8", xml_declaration=True)
        logger.debug(f"Файл успешно сформирован: {filename}")
    except Exception as e:
        logger.error(f"Ошибка при сохранении {filename}: {e}")


def generate_xml_okb(data_dict: dict, config: dict, logger: logging.Logger):
    logger.debug("Генерация XML для ОКБ.")
    org = config.get("organization", {})
    okb_conf = config.get("bureaus", {}).get("okb", {})
    
    now = datetime.now()
    source_id = okb_conf.get("sourceID", "02173")
    reg_num = f"CHP_{source_id}_EFK_04-10_{now.strftime('%Y%m%d%H%M%S')}000"
    subjects_count = str(len(data_dict) - 1)
    
    attr = {
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
        "xsi:noNamespaceSchemaLocation": "Main.xsd",
        "schemaVersion": "4.1",
        "inn": org.get("inn", ""),
        "ogrn": org.get("ogrn", ""),
        "sourceID": source_id,
        "regNumberDoc": reg_num,
        "dateDoc": now.strftime("%Y-%m-%d"),
        "subjectsCount": subjects_count,
        "groupBlocksCount": subjects_count,
        "regNumberDocInaccept": reg_num
    }
    
    root = ET.Element("Document", attr)
    # Здесь в будущем будет цикл по data_dict для добавления дочерних элементов
    
    save_xml(root, "okb_output.xml", logger)

def generate_xml_scoring(data_dict: dict, config: dict, logger: logging.Logger):
    logger.debug("Генерация XML для Скоринга.")
    org = config.get("organization", {})
    now = datetime.now()
    subjects_count = str(len(data_dict) - 1)
    
    attr = {
        "schemaVersion": "4.1",
        "inn": org.get("inn", ""),
        "ogrn": org.get("ogrn", ""),
        "sourceID": "DMH",
        "regNumberDoc": "2637",
        "dateDoc": now.strftime("%Y-%m-%d"),
        "subjectsCount": subjects_count,
        "groupBlocksCount": subjects_count,
        "regNumberDocInaccept": "2637"
    }
    
    root = ET.Element("Document", attr)
    save_xml(root, "scoring_output.xml", logger)

def generate_xml_kbrs(data_dict: dict, config: dict, logger: logging.Logger):
    logger.debug("Генерация XML для КБРС.")
    org = config.get("organization", {})
    now = datetime.now()
    subjects_count = str(len(data_dict) - 1)
    reg_num = f"KBRS_1136_{now.strftime('%Y%m%d')}_2637"
    
    attr = {
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
        "xsi:noNamespaceSchemaLocation": "Main.xsd",
        "schemaVersion": "4.1",
        "inn": org.get("inn", ""),
        "ogrn": org.get("ogrn", ""),
        "sourceID": "1136",
        "regNumberDoc": reg_num,
        "dateDoc": now.strftime("%Y-%m-%d"),
        "subjectsCount": subjects_count,
        "groupBlocksCount": subjects_count,
        "regNumberDocInaccept": reg_num
    }
    
    root = ET.Element("Document", attr)
    save_xml(root, "kbrs_output.xml", logger)

def generate_xml_nbki(data_dict: dict, config: dict, logger: logging.Logger):
    logger.debug("Генерация XML для НБКИ.")
    org = config.get("organization", {})
    now = datetime.now()
    subjects_count = str(len(data_dict) - 1)
    reg_num = f"SJ01SS000001_{now.strftime('%Y%m%d')}_{now.strftime('%H%M%S')}"
    
    attr = {
        "schemaVersion": "4.1",
        "dateDoc": now.strftime("%Y-%m-%d"),
        "inn": org.get("inn", ""),
        "ogrn": org.get("ogrn", ""),
        "sourceID": "SJ01SS000001",
        "regNumberDoc": reg_num,
        "subjectsCount": subjects_count,
        "groupBlocksCount": subjects_count,
        "regNumberDocInaccept": reg_num
    }
    
    root = ET.Element("Document", attr)
    save_xml(root, "nbki_output.xml", logger)

def run_conversion(file_path: Path, config_path: Path, logger: logging.Logger):
    logger.info("Запуск основного процесса конвертации...")
    
    config_data = load_config(config_path, logger)
    data_dict = excel_parser.parse_active_sheet(file_path, logger)
    
    if len(data_dict) <= 1:
        logger.warning("Нет данных для конвертации (файл пуст или содержит только заголовки).")
        return

    generate_xml_okb(data_dict, config_data, logger)
    generate_xml_scoring(data_dict, config_data, logger)
    generate_xml_kbrs(data_dict, config_data, logger)
    generate_xml_nbki(data_dict, config_data, logger)
    
    logger.info("Конвертация завершена!")