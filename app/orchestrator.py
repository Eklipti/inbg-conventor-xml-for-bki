from datetime import datetime
from pathlib import Path

from loguru import logger

from app import aggregation, excel_parser
from app.convertor import finalize_and_save_xml
from app.utils import (
    load_config,
    save_config,
)


def generate_xml_okb(
    data_dict: dict, config: dict, now: datetime, date_doc_str: str, output_dir: Path, save_file: bool = True
):
    """Подготавливает атрибуты и генерирует XML-файл в формате для ОКБ.

    Args:
        data_dict (dict): Словарь с распарсенными данными.
        config (dict): Конфигурация организации.
        now (datetime): Текущие дата и время.
        date_doc_str (str): Строковое представление даты документа.
        output_dir (Path): Директория для сохранения файла.
        save_file (bool, optional): Флаг сохранения файла на диск. По умолчанию True.
    """
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

    finalize_and_save_xml("okb", attr, f"{reg_num}.xml", data_dict, config, date_doc_str, output_dir, save_file)


def generate_xml_scoring(
    data_dict: dict,
    config: dict,
    now: datetime,
    date_doc_str: str,
    run_counter: int,
    output_dir: Path,
    save_file: bool = True,
):
    """Подготавливает атрибуты и генерирует XML-файл в формате для Скоринга.

    Args:
        data_dict (dict): Словарь с распарсенными данными.
        config (dict): Конфигурация организации.
        now (datetime): Текущие дата и время.
        date_doc_str (str): Строковое представление даты документа.
        run_counter (int): Текущий счетчик запусков.
        output_dir (Path): Директория для сохранения файла.
        save_file (bool, optional): Флаг сохранения файла на диск. По умолчанию True.
    """
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
    finalize_and_save_xml("scoring", attr, filename, data_dict, config, date_doc_str, output_dir, save_file)


def generate_xml_kbrs(
    data_dict: dict,
    config: dict,
    now: datetime,
    date_doc_str: str,
    run_counter: int,
    output_dir: Path,
    save_file: bool = True,
):
    """Подготавливает атрибуты и генерирует XML-файл в формате для КБРС.

    Args:
        data_dict (dict): Словарь с распарсенными данными.
        config (dict): Конфигурация организации.
        now (datetime): Текущие дата и время.
        date_doc_str (str): Строковое представление даты документа.
        run_counter (int): Текущий счетчик запусков.
        output_dir (Path): Директория для сохранения файла.
        save_file (bool, optional): Флаг сохранения файла на диск. По умолчанию True.
    """
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

    finalize_and_save_xml("kbrs", attr, f"{reg_num}.xml", data_dict, config, date_doc_str, output_dir, save_file)


def generate_xml_nbki(
    data_dict: dict, config: dict, now: datetime, date_doc_str: str, output_dir: Path, save_file: bool = True
):
    """Подготавливает атрибуты и генерирует XML-файл в формате для НБКИ.

    Args:
        data_dict (dict): Словарь с распарсенными данными.
        config (dict): Конфигурация организации.
        now (datetime): Текущие дата и время.
        date_doc_str (str): Строковое представление даты документа.
        output_dir (Path): Директория для сохранения файла.
        save_file (bool, optional): Флаг сохранения файла на диск. По умолчанию True.
    """
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

    finalize_and_save_xml("nbki", attr, f"{reg_num}.xml", data_dict, config, date_doc_str, output_dir, save_file)


def run_conversion(
    file_path: Path,
    config_path: Path,
    output_dir: Path = Path(),
    bki_list: list[str] | None = None,
    is_debug: bool = False,
    returns_path: Path | None = None,
) -> None:
    """Оркестрирует весь процесс конвертации из Excel в XML для разных БКИ.

    Парсит входной файл, загружает конфигурацию и последовательно запускает
    генерацию XML-файлов для всех поддерживаемых бюро. При наличии файла возвратов
    добавляет данные из него в основной файл.

    Args:
        file_path (Path): Путь к исходному Excel-файлу.
        config_path (Path): Путь к JSON-файлу конфигурации.
        output_dir (Path, optional): Директория для сохранения файлов.
        bki_list (list[str] | None, optional): Список требуемых БКИ.
        is_debug (bool, optional): Флаг режима отладки (использует статичный
            счетчик и не обновляет конфиг). По умолчанию False.
        returns_path (Path | None, optional): Опциональный путь к файлу возвратов.
    """
    config_data = load_config(config_path)

    if is_debug:
        logger.debug("Используется тестовый счетчик: 1111")
        run_counter = 1111
    else:
        run_counter = int(config_data.get("run_counter", 0))
        logger.debug(f"Текущий счётчик: {run_counter}")

    logger.info(f"Запуск процесса конвертации. Основной файл: {file_path}")

    try:
        if returns_path and returns_path.exists():
            logger.info(f"Файл возвратов: {returns_path}.")

            process_excel_returns_file_path = aggregation.process_excel_returns(returns_path, file_path)
            aggregation_file_path = aggregation.process_other_closures(returns_path, process_excel_returns_file_path)

            if not aggregation_file_path:
                logger.critical("Сбой агрегации файла возвратов.")
                return
        else:
            logger.debug("Файл возвратов не указан или не найден.")
            aggregation_file_path = file_path

        data_dict = excel_parser.parse_active_sheet(aggregation_file_path)
    except Exception as e:
        logger.critical(f"Критическая ошибка при чтении Excel: {e}")
        return

    if len(data_dict) <= 1:
        logger.warning(f"Файл {file_path.name} не содержит данных для обработки.")
        return

    now = datetime.now()
    date_doc_str = now.strftime("%Y-%m-%d")
    logger.info(f"Записывается дата: {date_doc_str}. Время: {now}")

    save_files = True
    if bki_list is None:
        active_bkis = ["okb", "scoring", "kbrs", "nbki"]
    elif len(bki_list) == 0:
        active_bkis = ["okb", "scoring", "kbrs", "nbki"]
        save_files = False
    else:
        active_bkis = bki_list

    if save_files and not output_dir.exists():
        output_dir.mkdir(parents=True, exist_ok=True)

    for bki in active_bkis:
        try:
            if bki == "okb":
                generate_xml_okb(data_dict, config_data, now, date_doc_str, output_dir, save_files)
                logger.success(f"Бюро {bki.upper()}: Конвертация успешно завершена.")

            if bki == "scoring":
                generate_xml_scoring(data_dict, config_data, now, date_doc_str, run_counter, output_dir, save_files)
                logger.success(f"Бюро {bki.upper()}: Конвертация успешно завершена.")

            if bki == "kbrs":
                generate_xml_kbrs(data_dict, config_data, now, date_doc_str, run_counter, output_dir, save_files)
                logger.success(f"Бюро {bki.upper()}: Конвертация успешно завершена.")

            if bki == "nbki":
                generate_xml_nbki(data_dict, config_data, now, date_doc_str, output_dir, save_files)
                logger.success(f"Бюро {bki.upper()}: Конвертация успешно завершена.")

        except Exception as e:
            logger.error(f"Ошибка при генерации XML для {bki.upper()}: {e}")

    logger.info("Процесс конвертации завершен.")

    if not is_debug and save_files:
        config_data["run_counter"] = run_counter + 1
        save_config(config_data, config_path)
