import json
import logging
import xml.etree.ElementTree as ET
from pathlib import Path

from utils import (
    format_date,
    clean_fio,
    clean_issuer,
    format_sum,
    calculate_days_difference,
    save_config,
    load_config,
    save_xml
)

dummy_logger = logging.getLogger("TestLogger")
dummy_logger.addHandler(logging.NullHandler())

# === ТЕСТЫ ДЛЯ РАБОТЫ С ДАТАМИ ===

def test_format_date():
    assert format_date("25.02.2026") == "2026-02-25"
    assert format_date(" 15.01.2020 ") == "2020-01-15"
    assert format_date("2026-02-25") == "2026-02-25"
    assert format_date("") == ""
    assert format_date(None) == ""

def test_calculate_days_difference():

    # Конечное значения всегда генерируется в правильном формате
    assert calculate_days_difference("01.01.2026", "2026-01-10") == "9"
    assert calculate_days_difference("2026-01-01", "2026-01-10") == "9"
    assert calculate_days_difference("10.01.2026", "2026-01-01") == "0"
    
    assert calculate_days_difference("", "2026-01-10") == ""
    assert calculate_days_difference("2026-01-10", "") == ""
    assert calculate_days_difference("непонятная_дата", "2026-01-10") == ""


# === ТЕСТЫ ДЛЯ ОЧИСТКИ ТЕКСТА ===

def test_clean_fio():
    assert clean_fio("Иванов_Иван\xa0Иванович") == "Иванов Иван Иванович"
    assert clean_fio("  Петров  ") == "Петров"
    assert clean_fio("") == ""
    assert clean_fio(None) == ""

def test_clean_issuer():
    assert clean_issuer("ОТДЕЛ * < > «» МВД") == "ОТДЕЛ МВД"
    assert clean_issuer("ГУ   МВД    РОССИИ") == "ГУ МВД РОССИИ"
    assert clean_issuer("") == ""
    
    long_text = "А" * 250
    assert len(clean_issuer(long_text)) == 200


# === ТЕСТЫ ДЛЯ СУММ ===

def test_format_sum():
    assert format_sum("1000") == "1000.00"
    assert format_sum("1000.5") == "1000.50"
    assert format_sum("1000,5") == "1000.50"  # Замена запятой
    assert format_sum("  1234.56  ") == "1234.56"
    assert format_sum("") == "0.00"
    assert format_sum(None) == "0.00"
    assert format_sum("какая-то_строка") == "0.00"


# === ТЕСТЫ ДЛЯ РАБОТЫ С ФАЙЛАМИ ===

def test_config_operations(tmp_path: Path):
    """Проверяем загрузку и сохранение конфига."""
    config_file = tmp_path / "test_config.json"
    test_data = {"run_counter": 2640, "bureau": "okb"}

    save_config(test_data, config_file, dummy_logger)
    assert config_file.exists()

    loaded_data = load_config(config_file, dummy_logger)
    assert loaded_data["run_counter"] == 2640
    assert loaded_data["bureau"] == "okb"

def test_save_xml(tmp_path: Path):
    """Проверяем, что XML корректно сохраняется на диск."""
    xml_file = tmp_path / "test_doc.xml"
    
    root = ET.Element("TestRoot")
    child = ET.SubElement(root, "Child")
    child.text = "Hello"

    save_xml(root, str(xml_file), dummy_logger)
    
    assert xml_file.exists()
    
    content = xml_file.read_text(encoding="UTF-8")
    assert "<?xml version='1.0' encoding='UTF-8'?>" in content
    assert "<TestRoot>" in content
    assert "<Child>Hello</Child>" in content