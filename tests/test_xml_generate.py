import pytest
import os
import json
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime
import logging

from convertor import generate_xml_scoring

dummy_logger = logging.getLogger("TestLogger")
dummy_logger.addHandler(logging.NullHandler())

EXAMPLE_DIR = Path(__file__).parent / "example"
CONFIG_PATH = EXAMPLE_DIR / "example.json"

@pytest.fixture
def test_environment(tmp_path: Path):
    """
    Фикстура: загружает ваш example.json, генерирует свежий XML 
    через ваш реальный код и отдает пути к ним для проверок.
    """
    assert CONFIG_PATH.exists(), f"Не найден конфиг: {CONFIG_PATH}"
    
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        config = json.load(f)

    data_dict = {
        0: "null",
        1: {
            "Фамилия": "Иванов", "Имя": "Иван", "Отчество": "Иванович",
            "Код страны": "643", "Серия": "1234", "Номер": "567890",
            "Дата выдачи": "01.01.2010", "Кем выдан": "МВД", "Дата рождения": "01.01.1990",
            "Уникальный идентификатор договора (сделки) БАНКА": "UID-111",
            "Дата согласия на обработку ПДН (дата договора)": "01.01.2020",
            "Статус долга (Операция)": "close 1",
            "Дата создания (дата передачи цессии)": "01.01.2025",
            "Общая сумма долга": "10000", "Остаток долга": "0", 
            "Сумма последнего возрата": "10000", "Дата последнего возврата": "10.01.2026"
        },
        2: {
            "Фамилия": "Петров", "Имя": "Петр", "Отчество": "Петрович",
            "Код страны": "643", "Серия": "4321", "Номер": "098765",
            "Дата выдачи": "01.01.2015", "Кем выдан": "УВД", "Дата рождения": "01.01.1995",
            "Уникальный идентификатор договора (сделки) БАНКА": "UID-222",
            "Дата согласия на обработку ПДН (дата договора)": "01.01.2021",
            "Статус долга (Операция)": "edit",
            "Дата создания (дата передачи цессии)": "01.01.2025",
            "Общая сумма долга": "50000", "Остаток долга": "20000",
            "Сумма последнего возрата": "5000", "Дата последнего возврата": "10.01.2026"
        }
    }

    now = datetime(2026, 2, 25, 12, 0, 0)
    date_doc_str = "2026-02-25"
    run_counter = config.get("run_counter", 9999)

    original_cwd = os.getcwd()
    os.chdir(tmp_path)
    try:
        generate_xml_scoring(data_dict, config, dummy_logger, now, date_doc_str, run_counter)
        generated_xml = tmp_path / f"DMH_FCH_20260225_{run_counter}.xml"
        assert generated_xml.exists(), "Генератор не создал файл!"
        
        return {"xml_path": generated_xml, "config": config}
    finally:
        os.chdir(original_cwd)


def test_xml_header_and_root(test_environment):
    xml_path = test_environment["xml_path"]

    with open(xml_path, 'r', encoding='utf-8') as f:
        first_line = f.readline().strip()
        assert first_line == "<?xml version='1.0' encoding='UTF-8'?>"

    tree = ET.parse(xml_path)
    root = tree.getroot()
    
    assert root.tag == "Document"
    assert root.attrib["schemaVersion"] == "4.1"
    assert root.attrib["sourceID"] == "DMH"
    assert "dateDoc" in root.attrib


def test_xml_source_matches_config(test_environment):
    xml_path = test_environment["xml_path"]
    org_config = test_environment["config"]["organization"]

    tree = ET.parse(xml_path)
    root = tree.getroot()
    
    org_source = root.find("Source/FL_46_UL_36_OrgSource")
    assert org_source is not None, "Блок Source не найден"
    
    assert org_source.find("regNum").text == org_config["ogrn"]
    assert org_source.find("fullName").text == org_config["fullName"]
    assert org_source.find("shortName").text == org_config["shortName"]
    
    tax_group = org_source.find("TaxNum_group_FL_46_UL_36_OrgSource")
    assert tax_group is not None, "Группа TaxNum_group не найдена"
    assert tax_group.find("taxNum").text == org_config["inn"]
    assert tax_group.find("taxCode").text == org_config["taxCode"]

def test_xml_events_structure(test_environment):
    xml_path = test_environment["xml_path"]
    tree = ET.parse(xml_path)
    root = tree.getroot()
    
    events = root.findall(".//Events/*")
    assert len(events) == 2, "Должно быть ровно 2 события"

    event_2_5 = root.find(".//FL_Event_2_5")
    assert event_2_5 is not None
    assert event_2_5.find("FL_17_DealUid/uid").text == "UID-111"
    
    group_25_28_ivanov = event_2_5.find("FL_25_26_27_28_Group")

    assert group_25_28_ivanov.find("exist_0") is not None
    
    assert group_25_28_ivanov.find("FL_25_Debt") is None
    assert group_25_28_ivanov.find("FL_26_DebtDue") is None
    
    fl_27_ivanov = group_25_28_ivanov.find("FL_27_DebtOverdue")
    assert fl_27_ivanov is not None
    assert fl_27_ivanov.find("missDuration").text == "0"
    
    fl_28_ivanov = group_25_28_ivanov.find("FL_28_Payment")
    assert fl_28_ivanov is not None
    assert fl_28_ivanov.find("paymentSum").text == "10000.00"

    event_2_3 = root.find(".//FL_Event_2_3")
    assert event_2_3 is not None
    assert event_2_3.find("FL_17_DealUid/uid").text == "UID-222"
    
    group_25_28_petrov = event_2_3.find("FL_25_26_27_28_Group")

    assert group_25_28_petrov.find("exist_1") is not None
    
    assert group_25_28_petrov.find("FL_25_Debt") is not None
    assert group_25_28_petrov.find("FL_26_DebtDue") is not None
    
    assert group_25_28_petrov.find("FL_27_DebtOverdue") is not None
    assert group_25_28_petrov.find("FL_28_Payment") is not None