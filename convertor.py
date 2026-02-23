import logging
import json
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime

import excel_parser

now = datetime.now()
date_doc_str = now.strftime("%Y-%m-%d")

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

def calculate_days_difference(start_date_str: str, end_date_str: str) -> str:
    """Вычисляет количество дней между двумя датами. Ожидает формат ДД.ММ.ГГГГ или ГГГГ-ММ-ДД."""
    if not start_date_str or not end_date_str:
        return "0"
        
    try:
        if "." in start_date_str:
            start_dt = datetime.strptime(start_date_str, "%d.%m.%Y")
        elif "-" in start_date_str:
            start_dt = datetime.strptime(start_date_str, "%Y-%m-%d")
        else:
            logger.error(f"Неизвестные формат даты: {start_date_str}")
            return "0"

        end_dt = datetime.strptime(end_date_str, "%Y-%m-%d")
        
        delta = (end_dt - start_dt).days
        return str(max(0, delta))
    except Exception:
        logger.error(f"Неизвестная дата.")
        return "0"


def build_source_block(parent_element: ET.Element, config: dict, date_str: str):
    """Универсальная функция для создания блока Source во всех форматах."""
    org = config.get("organization", {})
    
    source_elem = ET.SubElement(parent_element, "Source")
    org_source = ET.SubElement(source_elem, "FL_46_UL_36_OrgSource")
    
    ET.SubElement(org_source, "sourceCode").text = "1"
    ET.SubElement(org_source, "sourceRegistrationFact_1")
    ET.SubElement(org_source, "fullName").text = org.get("fullName", "")
    ET.SubElement(org_source, "shortName").text = org.get("shortName", "")
    ET.SubElement(org_source, "otherName").text = org.get("otherName", "")
    ET.SubElement(org_source, "sourceDateStart").text = org.get("sourceDateStart", "")
    ET.SubElement(org_source, "regNum").text = org.get("ogrn", "")
    
    tax_group = ET.SubElement(org_source, "TaxNum_group_FL_46_UL_36_OrgSource")
    ET.SubElement(tax_group, "taxCode").text = org.get("taxCode", "")
    ET.SubElement(tax_group, "taxNum").text = org.get("inn", "")
    
    # передаем строкой, так как форматы дат могут отличаться в других местах, но тут нужен YYYY-MM-DD
    ET.SubElement(org_source, "sourceCreditInfoDate").text = date_str

def build_title_block(subject_fl: ET.Element, row: dict):
    """Универсальная функция для создания блока Title (ФИО, паспорт, рождение)."""
    title = ET.SubElement(subject_fl, "Title")
    
    # === Блок ФЛ 1 и ФЛ 4 (Имя и Документ) ===
    fl_1_4_group = ET.SubElement(title, "FL_1_4_Group")
    
    fl_1_name = ET.SubElement(fl_1_4_group, "FL_1_Name")
    ET.SubElement(fl_1_name, "lastName").text = row.get("Фамилия", "")
    ET.SubElement(fl_1_name, "firstName").text = row.get("Имя", "")
    # Если отчества нет, .text будет пустой строкой - пустой тег
    ET.SubElement(fl_1_name, "middleName").text = row.get("Отчество", "")
    
    fl_4_doc = ET.SubElement(fl_1_4_group, "FL_4_Doc")
    ET.SubElement(fl_4_doc, "countryCode").text = "643"  # Россия
    ET.SubElement(fl_4_doc, "docCode").text = "21"       # Паспорт РФ
    ET.SubElement(fl_4_doc, "docSeries").text = row.get("Серия", "")
    ET.SubElement(fl_4_doc, "docNum").text = row.get("Номер", "")
    ET.SubElement(fl_4_doc, "issueDate").text = row.get("Дата выдачи", "")
    ET.SubElement(fl_4_doc, "docIssuer").text = row.get("Кем выдан", "")
    ET.SubElement(fl_4_doc, "deptCode").text = "000-000" # hardcore
    ET.SubElement(fl_4_doc, "foreignerCode").text = "0"
    
    # === Блок ФЛ 2 и ФЛ 5 (Предыдущие данные в реестре отсутствуют) ===
    fl_2_5_group = ET.SubElement(title, "FL_2_5_Group")
    
    fl_2_prev_name = ET.SubElement(fl_2_5_group, "FL_2_PrevName")
    ET.SubElement(fl_2_prev_name, "prevNameFlag_0")
    
    fl_5_prev_doc = ET.SubElement(fl_2_5_group, "FL_5_PrevDoc")
    ET.SubElement(fl_5_prev_doc, "prevDocFact_0")
    
    # === Блок ФЛ 3 (Рождение) ===
    fl_3_birth = ET.SubElement(title, "FL_3_Birth")
    ET.SubElement(fl_3_birth, "birthDate").text = row.get("Дата рождения", "")
    ET.SubElement(fl_3_birth, "countryCode").text = "643"
    ET.SubElement(fl_3_birth, "birthPlace").text = row.get("Место рождения", "")


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
    
    # === БЛОК SOURCE (Источник) ===
    build_source_block(root, config, date_doc_str)

    # === БЛОК DATA (Данные) ===
    data_elem = ET.SubElement(root, "Data")
    
    for key, row in data_dict.items():
        if key == 0:
            continue
            
        subject_fl = ET.SubElement(data_elem, "Subject_FL")

        build_title_block(subject_fl, row)
        
        events = ET.SubElement(subject_fl, "Events")
        
        # TODO: Здесь будем добавлять теги событий 2.3 или 2.5
        # В зависимости от логики статусов

    save_xml(root, f"{reg_num}.xml", logger)

def generate_xml_scoring(data_dict: dict, config: dict, logger: logging.Logger):
    logger.debug("Генерация XML для Скоринга.")
    org = config.get("organization", {})
    now = datetime.now()
    subjects_count = str(len(data_dict) - 1)
    
    scoring_conf = config.get("bureaus", {}).get("scoring", {})
    reg_num = scoring_conf.get("reg_num", "2637")

    attr = {
        "schemaVersion": "4.1",
        "inn": org.get("inn", ""),
        "ogrn": org.get("ogrn", ""),
        "sourceID": "DMH",
        "regNumberDoc": reg_num,
        "dateDoc": now.strftime("%Y-%m-%d"),
        "subjectsCount": subjects_count,
        "groupBlocksCount": subjects_count,
        "regNumberDocInaccept": reg_num
    }
    
    root = ET.Element("Document", attr)

    # === БЛОК SOURCE (Источник) ===
    build_source_block(root, config, date_doc_str)

    # === БЛОК DATA (Данные) ===
    data_elem = ET.SubElement(root, "Data")
    
    for key, row in data_dict.items():
        if key == 0:
            continue
            
        subject_fl = ET.SubElement(data_elem, "Subject_FL")

        build_title_block(subject_fl, row)
        
        events = ET.SubElement(subject_fl, "Events")
        
        # TODO: Здесь будем добавлять теги событий 2.3 или 2.5
        # В зависимости от логики статусов

    save_xml(root, f"DMH_FCH_{date_doc_str}_{reg_num}.xml", logger)

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

    # === БЛОК SOURCE (Источник) ===
    build_source_block(root, config, date_doc_str)

    # === БЛОК DATA (Данные) ===
    data_elem = ET.SubElement(root, "Data")
    
    for key, row in data_dict.items():
        if key == 0:
            continue
            
        subject_fl = ET.SubElement(data_elem, "Subject_FL")

        build_title_block(subject_fl, row)
        
        events = ET.SubElement(subject_fl, "Events")
        
        # TODO: Здесь будем добавлять теги событий 2.3 или 2.5
        # В зависимости от логики статусов

    save_xml(root, f"{reg_num}.xml", logger)

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

    # === БЛОК SOURCE (Источник) ===
    build_source_block(root, config, date_doc_str)

    # === БЛОК DATA (Данные) ===
    data_elem = ET.SubElement(root, "Data")
    
    for key, row in data_dict.items():
        if key == 0:
            continue
            
        subject_fl = ET.SubElement(data_elem, "Subject_FL")

        build_title_block(subject_fl, row)
        
        events = ET.SubElement(subject_fl, "Events")
        
        # TODO: Здесь будем добавлять теги событий 2.3 или 2.5
        # В зависимости от логики статусов
        
    save_xml(root, f"{reg_num}.xml", logger)

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