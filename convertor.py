import logging
import json
import xml.etree.ElementTree as ET
from pathlib import Path
import excel_parser

def load_config(config_path: Path, logger: logging.Logger) -> dict:
    """Загрузка конфигурации из JSON файла."""
    if not config_path.exists():
        logger.crtitical(f"Конфигурационный файл не найден: {config_path}.")
        sys.exit(1)
        
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
            logger.debug("Конфигурация успешно загружена.")
            return config_data
    except Exception as e:
        logger.crtitical(f"Ошибка при чтении {config_path}: {e}")
        sys.exit(1)


def generate_xml_okb(data_dict: dict, config: dict, logger: logging.Logger):
    logger.info("Генерация XML для ОКБ")
    pass

def generate_xml_scoring(data_dict: dict, config: dict, logger: logging.Logger):
    logger.info("Генерация XML для Скоринг")
    pass

def generate_xml_kbrs(data_dict: dict, config: dict, logger: logging.Logger):
    logger.info("Генерация XML для КБРС")
    pass

def generate_xml_nbki(data_dict: dict, config: dict, logger: logging.Logger):
    logger.info("Генерация XML для НБКИ")
    pass


def run_conversion(file_path: Path, config_path: Path, logger: logging.Logger):
    logger.debug("Запуск основного процесса конвертации.")
    
    config_data = load_config(config_path, logger)
    
    data_dict = excel_parser.parse_active_sheet(file_path, logger)
    
    sample_data = dict(list(data_dict.items())[:2])
    logger.debug(f"Тестовый вывод собранных данных: {sample_data}")
    logger.debug(f"Тестовый вывод конфига: {config_data}")
    
    generate_xml_okb(data_dict, config_data, logger)
    generate_xml_scoring(data_dict, config_data, logger)
    generate_xml_kbrs(data_dict, config_data, logger)
    generate_xml_nbki(data_dict, config_data, logger)
    
    logger.info("Конвертация завершена")