import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

from loguru import logger

import excel_parser
from utils import (
    calculate_days_difference,
    clean_fio,
    clean_issuer,
    format_date,
    format_sum,
    load_config,
    save_config,
    save_xml,
)


def build_source_block(parent_element: ET.Element, config: dict, date_str: str):
    """Формирует блок <Source> (данные об организации) в XML-дереве.

    Args:
        parent_element (ET.Element): Родительский XML-элемент (<Document>).
        config (dict): Словарь с конфигурацией (данные об организации).
        date_str (str): Дата формирования документа в формате YYYY-MM-DD.
    """
    logger.trace("Формирование блока <Source>.")

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

    ET.SubElement(org_source, "sourceCreditInfoDate").text = date_str


def build_title_block(subject_fl: ET.Element, row: dict):
    """Формирует блок <Title> (титульные данные субъекта) в XML-дереве.

    Заполняет информацию о ФИО, документе удостоверяющем личность и дате/месте
    рождения на основе переданной строки данных. Автоматически обрабатывает
    ситуации с отсутствующими паспортными данными.

    Args:
        subject_fl (ET.Element): Родительский элемент <Subject_FL>.
        row (dict): Словарь с данными одной записи (строки) из Excel.
    """
    logger.trace("Формирование блока <Title>.")

    title = ET.SubElement(subject_fl, "Title")

    fl_1_4_group = ET.SubElement(title, "FL_1_4_Group")

    fl_1_name = ET.SubElement(fl_1_4_group, "FL_1_Name")
    ET.SubElement(fl_1_name, "lastName").text = clean_fio(row.get("Фамилия", ""))
    ET.SubElement(fl_1_name, "firstName").text = clean_fio(row.get("Имя", ""))
    middle_name = clean_fio(row.get("Отчество", ""))
    if middle_name:
        ET.SubElement(fl_1_name, "middleName").text = middle_name

    country_code = str(row.get("Код страны", "")).strip()
    doc_series = str(row.get("Серия", "")).strip()
    doc_num = str(row.get("Номер", "")).strip()
    doc_code = "21"

    if not country_code:
        logger.debug('Не найден код страны, замена на "9", серия и номер паспорта замены нулями.')
        doc_code = "9"
        doc_series = "00"
        doc_num = "000000"

    fl_4_doc = ET.SubElement(fl_1_4_group, "FL_4_Doc")
    ET.SubElement(fl_4_doc, "countryCode").text = "643"  # Россия
    ET.SubElement(fl_4_doc, "docCode").text = doc_code
    ET.SubElement(fl_4_doc, "docSeries").text = doc_series
    ET.SubElement(fl_4_doc, "docNum").text = doc_num
    ET.SubElement(fl_4_doc, "issueDate").text = format_date(row.get("Дата выдачи", ""))
    ET.SubElement(fl_4_doc, "docIssuer").text = clean_issuer(row.get("Кем выдан", ""))
    ET.SubElement(fl_4_doc, "deptCode").text = "000-000"
    ET.SubElement(fl_4_doc, "foreignerCode").text = "3"

    fl_2_5_group = ET.SubElement(title, "FL_2_5_Group")
    ET.SubElement(ET.SubElement(fl_2_5_group, "FL_2_PrevName"), "prevNameFlag_0")
    ET.SubElement(ET.SubElement(fl_2_5_group, "FL_5_PrevDoc"), "prevDocFact_0")

    fl_3_birth = ET.SubElement(title, "FL_3_Birth")
    ET.SubElement(fl_3_birth, "birthDate").text = format_date(row.get("Дата рождения", ""))
    ET.SubElement(fl_3_birth, "countryCode").text = "643"
    birth_place = clean_issuer(row.get("Место рождения", ""))
    if birth_place:
        ET.SubElement(fl_3_birth, "birthPlace").text = birth_place


def build_event(
    events_elem: ET.Element,
    row: dict,
    date_doc_str: str,
    bureau: str,
    order_num: int,
    event_type: str,
    extra_param: str | None = None,
) -> None:
    """Формирует основной XML-блок кредитного события.

    Создает базовые теги события (сделка, суммы, условия) и, в зависимости от
    переданного типа (2.3, 2.5, 2.11.2), вызывает соответствующие функции-помощники
    для добавления специфичных подблоков.

    Args:
        events_elem (ET.Element): Родительский элемент <Events>.
        row (dict): Словарь с данными одной записи из Excel.
        date_doc_str (str): Дата формирования документа.
        bureau (str): Идентификатор целевого БКИ.
        order_num (int): Порядковый номер события.
        event_type (str): Тип события (например, '2_3', '2_5', '2_11_2').
        extra_param (str | None, optional): Дополнительный параметр (например, код закрытия).
    """
    logger.trace("Формирование событий <FL_Event_*>.")

    status_raw = row.get("Статус долга (Операция)", "").strip().lower()

    op_code = "A" if status_raw == "add" else "B"

    event_tag = f"FL_Event_{event_type}"
    event_attrs = {"orderNum": str(order_num), "eventDate": date_doc_str, "operationCode": op_code}

    event_elem = ET.SubElement(events_elem, event_tag, event_attrs)

    if row.get("Статус долга (Операция)", "").strip().lower() == "add":
        build_suffix_2_11_2(event_elem)

    fl_17 = ET.SubElement(event_elem, "FL_17_DealUid")
    ET.SubElement(fl_17, "uid").text = row.get("Уникальный идентификатор договора (сделки) БАНКА", "")
    ET.SubElement(fl_17, "openDate").text = format_date(row.get("Дата согласия на обработку ПДН (дата договора)", ""))

    fl_18 = ET.SubElement(event_elem, "FL_18_Deal")
    ET.SubElement(fl_18, "role").text = "1"
    ET.SubElement(fl_18, "code").text = "1"
    ET.SubElement(fl_18, "kindCode").text = "99"
    ET.SubElement(fl_18, "purposeCode").text = "99"
    ET.SubElement(fl_18, "consumerExist_0")
    ET.SubElement(fl_18, "cardExist_0")
    ET.SubElement(fl_18, "novationExist_0")
    ET.SubElement(fl_18, "monetarySourceExist_1")
    ET.SubElement(fl_18, "monetarySubjectExist_1")
    ET.SubElement(fl_18, "endDate").text = "9999-12-31"
    ET.SubElement(fl_18, "creditorCode").text = "1"
    ET.SubElement(fl_18, "partialExist_1")
    ET.SubElement(fl_18, "transferUid").text = row.get("Уникальный идентификатор договора (сделки) БАНКА", "")
    ET.SubElement(fl_18, "creditLineExist_0")
    ET.SubElement(fl_18, "floatRateExist_0")
    ET.SubElement(fl_18, "partialTransferExist_0")
    ET.SubElement(fl_18, "startDate").text = format_date(row.get("Дата создания (дата передачи цессии)", ""))
    ET.SubElement(fl_18, "repaymentFact_0")
    ET.SubElement(fl_18, "transferFact_0")
    ET.SubElement(fl_18, "partnerFinancingFact_0")

    fl_19 = ET.SubElement(event_elem, "FL_19_Amount")
    ET.SubElement(fl_19, "sum").text = format_sum(row.get("Общая сумма долга"))
    ET.SubElement(fl_19, "currency").text = "RUB"
    ET.SubElement(fl_19, "calcDate").text = format_date(row.get("Дата создания (дата передачи цессии)", ""))

    fl_19_1 = ET.SubElement(event_elem, "FL_19_1_AmountInfo")
    ET.SubElement(fl_19_1, "securityFact_0")

    fl_21 = ET.SubElement(event_elem, "FL_21_PaymentTerms")
    ET.SubElement(fl_21, "mainPaySum").text = "0.00"
    ET.SubElement(fl_21, "percentPaySum").text = "0.00"

    if event_type in ("2_3", "2_11_2"):
        build_suffix_2_3(event_elem, row, date_doc_str, bureau)
    elif event_type == "2_5":
        build_suffix_2_5(event_elem, row, date_doc_str, bureau, extra_param)


def build_fl_27_28(group_25_28: ET.Element, row: dict, date_doc_str: str, event_type: str):
    """Формирует блоки FL_27 (Просроченная задолженность) и FL_28 (Внесенные платежи).

    Args:
        group_25_28 (ET.Element): Родительский элемент группы 25-28.
        row (dict): Словарь с данными одной записи из Excel.
        date_doc_str (str): Дата формирования документа.
        event_type (str): Тип события для корректировки логики расчета дней просрочки.
    """
    logger.trace("Формирование блоков <FL_27> и <FL_28>")

    debt_remains = format_sum(row.get("Остаток долга"))

    fl_27 = ET.SubElement(group_25_28, "FL_27_DebtOverdue")
    ET.SubElement(fl_27, "missFact_1")
    ET.SubElement(fl_27, "debtOverdueSum").text = debt_remains
    ET.SubElement(fl_27, "debtOverdueMainSum").text = debt_remains
    ET.SubElement(fl_27, "debtOverduePercentSum").text = "0.00"
    ET.SubElement(fl_27, "debtOverdueOtherSum").text = "0.00"

    miss_date = format_date(row.get("Дата создания (дата передачи цессии)", ""))

    ET.SubElement(fl_27, "debtOverdueStartDate").text = miss_date
    ET.SubElement(fl_27, "mainMissDate").text = miss_date
    ET.SubElement(fl_27, "percentMissDate").text = miss_date

    miss_days = calculate_days_difference(miss_date, date_doc_str)

    if event_type == "2_5":
        ET.SubElement(fl_27, "missDuration").text = "0"
    else:
        ET.SubElement(fl_27, "missDuration").text = miss_days

    ET.SubElement(fl_27, "repaidMissDuration").text = miss_days

    fl_28 = ET.SubElement(group_25_28, "FL_28_Payment")
    last_pay = format_sum(row.get("Сумма последнего возрата", ""))
    ET.SubElement(fl_28, "paymentSum").text = last_pay
    ET.SubElement(fl_28, "paymentMainSum").text = last_pay
    ET.SubElement(fl_28, "paymentPercentSum").text = "0.00"
    ET.SubElement(fl_28, "paymentOtherSum").text = "0.00"

    # Логика подсчёта: totalSum = totalMainSum + totalPercentSum + totalOtherSum
    ET.SubElement(fl_28, "totalSum").text = last_pay
    ET.SubElement(fl_28, "totalMainSum").text = last_pay
    ET.SubElement(fl_28, "totalPercentSum").text = "0.00"
    ET.SubElement(fl_28, "totalOtherSum").text = "0.00"

    if last_pay != "0.00":
        ET.SubElement(fl_28, "date").text = date_doc_str
    ET.SubElement(fl_28, "sizeCode").text = "3"
    ET.SubElement(fl_28, "scheduleCode").text = "3"
    ET.SubElement(fl_28, "lastMissPaySum").text = last_pay
    ET.SubElement(fl_28, "paySum24").text = "0.00"
    ET.SubElement(fl_28, "payCurrency").text = "RUB"


def build_suffix_2_11_2(event_elem: ET.Element):
    """Добавляет специфичные блоки (адреса, ИП, дееспособность) для события типа 'add' (2.11.2).

    Args:
        event_elem (ET.Element): Родительский элемент события.
    """
    logger.trace("Формирование блоков для <FL_2_11_2>")

    fl_8 = ET.SubElement(event_elem, "FL_8_AddrReg")
    ET.SubElement(fl_8, "code").text = "3"

    fl_9 = ET.SubElement(event_elem, "FL_9_AddrFact")
    ET.SubElement(fl_9, "exist_0")

    fl_11 = ET.SubElement(event_elem, "FL_11_IndividualEntrepreneur")
    ET.SubElement(fl_11, "regFact_0")

    fl_12 = ET.SubElement(event_elem, "FL_12_Capacity")
    ET.SubElement(fl_12, "code").text = "1"


def build_suffix_2_3(event_elem: ET.Element, row: dict, date_doc_str: str, bureau: str):
    """Добавляет специфичные блоки (задолженность, учет) для события типа 'edit' (2.3).

    Args:
        event_elem (ET.Element): Родительский элемент события.
        row (dict): Словарь с данными записи.
        date_doc_str (str): Дата формирования документа.
        bureau (str): Идентификатор БКИ.
    """
    logger.trace("Формирование блоков для <FL_2_3>")

    group_25_28 = ET.SubElement(event_elem, "FL_25_26_27_28_Group")

    last_pay = format_sum(row.get("Сумма последнего возрата", ""))
    last_pay_exist_tag = "lastPayExist_0" if last_pay == "0.00" else "lastPayExist_1"

    ET.SubElement(group_25_28, last_pay_exist_tag)
    ET.SubElement(group_25_28, "calcDate").text = date_doc_str
    ET.SubElement(group_25_28, "exist_1")

    debt_remains = format_sum(row.get("Остаток долга"))

    fl_25 = ET.SubElement(group_25_28, "FL_25_Debt")
    ET.SubElement(fl_25, "debtSum").text = debt_remains
    ET.SubElement(fl_25, "debtMainSum").text = debt_remains
    ET.SubElement(fl_25, "debtPercentSum").text = "0.00"
    ET.SubElement(fl_25, "debtOtherSum").text = "0.00"
    ET.SubElement(fl_25, "graceUnconfExist_0")
    ET.SubElement(fl_25, "currency").text = "RUB"

    fl_26 = ET.SubElement(group_25_28, "FL_26_DebtDue")
    ET.SubElement(fl_26, "debtDueExist_0")

    build_fl_27_28(group_25_28, row, date_doc_str, "2_3")

    ET.SubElement(ET.SubElement(event_elem, "FL_20_JointDebtors"), "exist_0")
    ET.SubElement(ET.SubElement(event_elem, "FL_36_1_ProvisionPaymentOffset"), "exist_0")

    fl_54 = ET.SubElement(event_elem, "FL_54_Accounting")
    ET.SubElement(fl_54, "exist_1")
    ET.SubElement(fl_54, "minInterest").text = "0.00"
    ET.SubElement(fl_54, "maxInterest").text = "0.00"
    ET.SubElement(fl_54, "supportExist_0")
    ET.SubElement(fl_54, "calcDate").text = date_doc_str

    fl_56 = ET.SubElement(event_elem, "FL_56_Participation")
    ET.SubElement(fl_56, "role").text = "1"
    ET.SubElement(fl_56, "kindCode").text = "99"
    ET.SubElement(fl_56, "uid").text = row.get("Уникальный идентификатор договора (сделки) БАНКА", "")
    ET.SubElement(fl_56, "fundDate").text = format_date(row.get("Дата согласия на обработку ПДН (дата договора)", ""))
    ET.SubElement(fl_56, "overdueExist_1")
    ET.SubElement(fl_56, "stopExist_0")


def build_suffix_2_5(event_elem: ET.Element, row: dict, date_doc_str: str, bureau: str, extra_param: str | None):
    """Добавляет специфичные блоки (завершение договора) для события типа 'close' (2.5).

    Args:
        event_elem (ET.Element): Родительский элемент события.
        row (dict): Словарь с данными записи.
        date_doc_str (str): Дата формирования документа.
        bureau (str): Идентификатор БКИ.
        extra_param (str | None): Код завершения договора.
    """
    logger.trace("Формирование блоков для <FL_2_5>")

    group_25_28 = ET.SubElement(event_elem, "FL_25_26_27_28_Group")

    last_pay = format_sum(row.get("Сумма последнего возрата", ""))
    last_pay_exist_tag = "lastPayExist_0" if last_pay == "0.00" else "lastPayExist_1"

    ET.SubElement(group_25_28, last_pay_exist_tag)
    ET.SubElement(group_25_28, "calcDate").text = format_date(row.get("Дата последнего возврата", ""))
    ET.SubElement(group_25_28, "exist_0")

    build_fl_27_28(group_25_28, row, date_doc_str, "2_5")

    fl_38 = ET.SubElement(event_elem, "FL_38_ContractEnd")
    ET.SubElement(fl_38, "date").text = date_doc_str

    contract_code = extra_param if extra_param else "1"
    ET.SubElement(fl_38, "code").text = contract_code

    fl_56 = ET.SubElement(event_elem, "FL_56_Participation")
    ET.SubElement(fl_56, "role").text = "1"
    ET.SubElement(fl_56, "kindCode").text = "99"
    ET.SubElement(fl_56, "uid").text = row.get("Уникальный идентификатор договора (сделки) БАНКА", "")
    ET.SubElement(fl_56, "fundDate").text = format_date(row.get("Дата согласия на обработку ПДН (дата договора)", ""))

    ET.SubElement(fl_56, "overdueExist_0")
    ET.SubElement(fl_56, "stopExist_1")


def build_data_block(root: ET.Element, data_dict: dict, date_doc_str: str, bureau: str):
    """Формирует глобальный блок <Data> со всеми субъектами и их событиями.

    Итерируется по распарсенным данным, игнорирует техническую нулевую строку,
    строит титульные данные и определяет тип события на основе поля статуса долга.

    Args:
        root (ET.Element): Корневой элемент <Document>.
        data_dict (dict): Полный словарь распарсенных данных из Excel.
        date_doc_str (str): Дата формирования документа.
        bureau (str): Идентификатор целевого БКИ.
    """
    logger.trace("Формирование блока <Data>")

    data_elem = ET.SubElement(root, "Data")

    event_counter = 1

    for key, row in data_dict.items():
        if key == 0:
            continue

        logger.trace(f"Текущий субъект: {event_counter}")
        subject_fl = ET.SubElement(data_elem, "Subject_FL")
        build_title_block(subject_fl, row)

        events = ET.SubElement(subject_fl, "Events")

        status_raw = row.get("Статус долга (Операция)", "").strip().lower()
        if status_raw.startswith("add"):
            build_event(events, row, date_doc_str, bureau, event_counter, event_type="2_11_2")
            event_counter += 1
        elif status_raw.startswith("edit"):
            build_event(events, row, date_doc_str, bureau, event_counter, event_type="2_3")
            event_counter += 1
        elif status_raw.startswith("close"):
            parts = status_raw.split()
            close_digit = parts[1] if len(parts) > 1 else None
            build_event(events, row, date_doc_str, bureau, event_counter, event_type="2_5", extra_param=close_digit)
            event_counter += 1


def finalize_and_save_xml(bureau: str, attr: dict, filename: str, data_dict: dict, config: dict, date_doc_str: str):
    """Обертка для финальной сборки XML-дерева и его сохранения на диск.

    Args:
        bureau (str): Идентификатор БКИ.
        attr (dict): Словарь атрибутов для корневого тега <Document>.
        filename (str): Имя итогового файла.
        data_dict (dict): Данные из Excel.
        config (dict): Конфигурация организации.
        date_doc_str (str): Дата формирования документа.
    """
    logger.trace("Сборка XML-дерева и сохранения на диск")

    root = ET.Element("Document", attr)
    build_source_block(root, config, date_doc_str)
    build_data_block(root, data_dict, date_doc_str, bureau=bureau)
    save_xml(root, filename)


def generate_xml_okb(data_dict: dict, config: dict, now: datetime, date_doc_str: str):
    """Подготавливает атрибуты и генерирует XML-файл в формате для ОКБ."""
    logger.info("Генерация XML для ОКБ.")
    org = config.get("organization", {})
    okb_conf = config.get("bureaus", {}).get("okb", {})

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
        "dateDoc": date_doc_str,
        "subjectsCount": subjects_count,
        "groupBlocksCount": subjects_count,
        "regNumberDocInaccept": reg_num,
    }

    finalize_and_save_xml("okb", attr, f"{reg_num}.xml", data_dict, config, date_doc_str)


def generate_xml_scoring(data_dict: dict, config: dict, now: datetime, date_doc_str: str, run_counter: int):
    """Подготавливает атрибуты и генерирует XML-файл в формате для Скоринга."""
    logger.info("Генерация XML для Скоринга.")
    org = config.get("organization", {})
    subjects_count = str(len(data_dict) - 1)

    reg_num = str(run_counter)

    attr = {
        "schemaVersion": "4.1",
        "inn": org.get("inn", ""),
        "ogrn": org.get("ogrn", ""),
        "sourceID": "DMH",
        "regNumberDoc": reg_num,
        "dateDoc": date_doc_str,
        "subjectsCount": subjects_count,
        "groupBlocksCount": subjects_count,
        "regNumberDocInaccept": reg_num,
    }

    date_file_str = now.strftime("%Y%m%d")
    filename = f"DMH_FCH_{date_file_str}_{reg_num}.xml"
    finalize_and_save_xml("scoring", attr, filename, data_dict, config, date_doc_str)


def generate_xml_kbrs(data_dict: dict, config: dict, now: datetime, date_doc_str: str, run_counter: int):
    """Подготавливает атрибуты и генерирует XML-файл в формате для КБРС."""
    logger.info("Генерация XML для КБРС.")
    org = config.get("organization", {})
    subjects_count = str(len(data_dict) - 1)

    reg_num = f"KBRS_1136_{now.strftime('%Y%m%d')}_{run_counter}"

    attr = {
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
        "xsi:noNamespaceSchemaLocation": "Main.xsd",
        "schemaVersion": "4.1",
        "inn": org.get("inn", ""),
        "ogrn": org.get("ogrn", ""),
        "sourceID": "1136",
        "regNumberDoc": reg_num,
        "dateDoc": date_doc_str,
        "subjectsCount": subjects_count,
        "groupBlocksCount": subjects_count,
        "regNumberDocInaccept": reg_num,
    }

    finalize_and_save_xml("kbrs", attr, f"{reg_num}.xml", data_dict, config, date_doc_str)


def generate_xml_nbki(data_dict: dict, config: dict, now: datetime, date_doc_str: str):
    """Подготавливает атрибуты и генерирует XML-файл в формате для НБКИ."""
    logger.info("Генерация XML для НБКИ.")
    org = config.get("organization", {})
    subjects_count = str(len(data_dict) - 1)
    reg_num = f"SJ01SS000001_{now.strftime('%Y%m%d')}_{now.strftime('%H%M%S')}"

    attr = {
        "schemaVersion": "4.1",
        "dateDoc": date_doc_str,
        "inn": org.get("inn", ""),
        "ogrn": org.get("ogrn", ""),
        "sourceID": "SJ01SS000001",
        "regNumberDoc": reg_num,
        "subjectsCount": subjects_count,
        "groupBlocksCount": subjects_count,
        "regNumberDocInaccept": reg_num,
    }

    finalize_and_save_xml("nbki", attr, f"{reg_num}.xml", data_dict, config, date_doc_str)


def run_conversion(file_path: Path, config_path: Path, is_debug: bool = False):
    """Оркестрирует весь процесс конвертации из Excel в XML для разных БКИ.

    Парсит входной файл, загружает конфигурацию и последовательно запускает
    генерацию XML-файлов для всех поддерживаемых бюро. В зависимости от режима,
    использует боевой или отладочный счетчик запусков.

    Args:
        file_path (Path): Путь к исходному Excel-файлу.
        config_path (Path): Путь к JSON-файлу конфигурации.
        is_debug (bool, optional): Флаг режима отладки (использует статичный
            счетчик и не обновляет конфиг). По умолчанию False.
    """
    logger.info("Запуск процесса конвертации в XML.")

    data_dict = excel_parser.parse_active_sheet(file_path)

    if len(data_dict) <= 1:
        logger.warning("Нет данных для конвертации.")
        return

    now = datetime.now()
    date_doc_str = now.strftime("%Y-%m-%d")
    logger.trace(f"Записывается дата: {date_doc_str}. Время: {now}")

    config_data = load_config(config_path)

    if is_debug:
        logger.debug("Используется тестовый счетчик: 1111")
        run_counter = 1111
    else:
        run_counter = config_data.get("run_counter")
        logger.debug(f"Текущий счётчик: {run_counter}")

    generate_xml_okb(data_dict, config_data, now, date_doc_str)
    logger.info("Файл для ОКБ сформирован.")
    generate_xml_scoring(data_dict, config_data, now, date_doc_str, run_counter)
    logger.info("Файл для Скоринга сформирован.")
    generate_xml_kbrs(data_dict, config_data, now, date_doc_str, run_counter)
    logger.info("Файл для КБРС сформирован.")
    generate_xml_nbki(data_dict, config_data, now, date_doc_str)
    logger.info("Файл для НБКИ сформирован.")

    if not is_debug:
        config_data["run_counter"] = run_counter + 1
        save_config(config_data, config_path)

    logger.info("Конвертация завершена.")
