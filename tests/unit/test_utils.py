import json
import xml.etree.ElementTree as ET
from pathlib import Path
from unittest.mock import patch

import pytest
from loguru import logger

from utils import (
    calculate_days_difference,
    clean_fio,
    clean_issuer,
    convert_xls_to_xlsx,
    format_date,
    format_sum,
    load_config,
    save_config,
    save_xml,
    validate_config,
)


@pytest.fixture
def valid_config_data():
    return {
        "organization": {
            "inn": "1234567890",
            "ogrn": "1234567890123",
            "fullName": "Полное Название ООО",
            "shortName": "ООО ПН",
            "otherName": "-",
            "sourceDateStart": "2023-01-01",
        },
        "run_counter": 5,
    }


@patch("utils.pd.read_excel")
def test_convert_xls_to_xlsx_critical_error(mock_read_excel, tmp_path):
    mock_read_excel.side_effect = Exception("Excel read error")
    xls_path = tmp_path / "test.xls"

    with pytest.raises(SystemExit) as exc_info:
        convert_xls_to_xlsx(xls_path)

    assert exc_info.value.code == 1


# ==========================================
# Тесты для validate_config
# ==========================================
def test_validate_config_success(valid_config_data):
    assert validate_config(valid_config_data) is True


def test_validate_config_missing_organization(valid_config_data):
    del valid_config_data["organization"]
    assert validate_config(valid_config_data) is False


@pytest.mark.parametrize("missing_key", ["inn", "ogrn", "fullName", "shortName", "otherName", "sourceDateStart"])
def test_validate_config_missing_org_keys(valid_config_data, missing_key):
    del valid_config_data["organization"][missing_key]
    assert validate_config(valid_config_data) is False


def test_validate_config_invalid_run_counter(valid_config_data):
    valid_config_data["run_counter"] = "5"  # Строка вместо int
    assert validate_config(valid_config_data) is False

    del valid_config_data["run_counter"]
    assert validate_config(valid_config_data) is False


# ==========================================
# Тесты для load_config
# ==========================================
def test_load_config_success(tmp_path, valid_config_data):
    config_file = tmp_path / "config.json"
    with open(config_file, "w", encoding="utf-8") as f:
        json.dump(valid_config_data, f)

    loaded_data = load_config(config_file)
    assert loaded_data == valid_config_data


def test_load_config_file_not_found(tmp_path):
    missing_file = tmp_path / "missing.json"
    with pytest.raises(SystemExit) as exc_info:
        load_config(missing_file)
    assert exc_info.value.code == 1


def test_load_config_invalid_json(tmp_path):
    config_file = tmp_path / "config.json"
    with open(config_file, "w", encoding="utf-8") as f:
        f.write("{invalid json format,")

    with pytest.raises(SystemExit) as exc_info:
        load_config(config_file)
    assert exc_info.value.code == 1


def test_load_config_invalid_structure(tmp_path, valid_config_data):
    del valid_config_data["run_counter"]  # Ломаем структуру
    config_file = tmp_path / "config.json"
    with open(config_file, "w", encoding="utf-8") as f:
        json.dump(valid_config_data, f)

    with pytest.raises(SystemExit) as exc_info:
        load_config(config_file)
    assert exc_info.value.code == 1


# ==========================================
# Тесты для format_date
# ==========================================
@pytest.mark.parametrize(
    "input_date, expected",
    [
        ("31.12.2023", "2023-12-31"),
        (" 15.05.2022 ", "2022-05-15"),
        ("2023-10-10", "2023-10-10"),
        ("", ""),
        (None, ""),
        ("не дата", "не дата"),
    ],
)
def test_format_date(input_date, expected):
    assert format_date(input_date) == expected


# ==========================================
# Тесты для clean_fio
# ==========================================
@pytest.mark.parametrize(
    "input_fio, expected",
    [
        ("Иванов_Иван\xa0Иванович", "Иванов Иван Иванович"),
        ("   Петров Петр   ", "Петров Петр"),
        ("NaN", ""),
        ("None", ""),
        ("нет", ""),
        ("-", ""),
        (".", ""),
        ("", ""),
        (None, ""),
    ],
)
def test_clean_fio(input_fio, expected):
    assert clean_fio(input_fio) == expected


# ==========================================
# Тесты для clean_issuer
# ==========================================
@pytest.mark.parametrize(
    "input_text, expected",
    [
        ('мвд *по* <г.> москва "отдел«»"', "МВД ПО Г. МОСКВА ОТДЕЛ"),
        ("   очень   много    пробелов   ", "ОЧЕНЬ МНОГО ПРОБЕЛОВ"),
        ("", ""),
        (None, ""),
    ],
)
def test_clean_issuer(input_text, expected):
    assert clean_issuer(input_text) == expected


def test_clean_issuer_truncation():
    long_string = "А" * 250
    result = clean_issuer(long_string)
    assert len(result) == 200
    assert result == "А" * 200


# ==========================================
# Тесты для save_xml
# ==========================================
@patch("utils.ET.ElementTree.write")
def test_save_xml_success(mock_write):
    root = ET.Element("root")
    save_xml(root, "test.xml")
    mock_write.assert_called_once_with("test.xml", encoding="UTF-8", xml_declaration=True)


@patch("utils.ET.ElementTree.write")
def test_save_xml_exception(mock_write):
    mock_write.side_effect = Exception("Write error")
    root = ET.Element("root")

    messages = []
    handler_id = logger.add(lambda msg: messages.append(msg))

    try:
        save_xml(root, "test.xml")

        assert any("Ошибка при сохранении test.xml" in msg for msg in messages)
    finally:
        logger.remove(handler_id)

# ==========================================
# Тесты для calculate_days_difference
# ==========================================
@pytest.mark.parametrize(
    "start, end, expected",
    [
        ("01.01.2023", "2023-01-10", "9"),
        ("2023-01-01", "2023-01-10", "9"),
        (" 01.01.2023 ", " 2023-01-10 ", "9"),
        ("15.01.2023", "2023-01-10", "0"),
        ("invalid_date", "2023-01-10", ""),
        ("", "2023-01-10", ""),
        ("01.01.2023", "", ""),
    ],
)
def test_calculate_days_difference(start, end, expected):
    assert calculate_days_difference(start, end) == expected


# ==========================================
# Тесты для format_sum
# ==========================================
@pytest.mark.parametrize(
    "input_val, expected",
    [
        (123.456, "123.46"),
        ("123,45", "123.45"),
        ("  45.6  ", "45.60"),
        (100, "100.00"),
        (0, "0.00"),
        ("", "0.00"),
        (None, "0.00"),
        ("не число", "0.00"),
    ],
)
def test_format_sum(input_val, expected):
    assert format_sum(input_val) == expected


# ==========================================
# Тесты для save_config
# ==========================================
def test_save_config_success(tmp_path, valid_config_data):
    config_file = tmp_path / "config.json"

    save_config(valid_config_data, config_file)

    assert config_file.exists()
    with open(config_file, encoding="utf-8") as f:
        saved_data = json.load(f)
    assert saved_data == valid_config_data


@patch("utils.open")
def test_save_config_exception(mock_open, valid_config_data):
    mock_open.side_effect = Exception("File system error")

    messages = []
    handler_id = logger.add(lambda msg: messages.append(msg))

    try:
        save_config(valid_config_data, Path("dummy.json"))
        assert any("Ошибка при перезаписи dummy.json" in msg for msg in messages)
    finally:
        logger.remove(handler_id)
