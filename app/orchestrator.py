from datetime import datetime
from pathlib import Path

from loguru import logger

from app import aggregation, excel_parser
from app.convertor import finalize_and_save_xml
from app.utils import (
    load_config,
    save_config, validate_excel, convert_xls_to_xlsx,
)


def generate_xml_bureau(
    bureau_type: str,
    data_dict: dict,
    config: dict,
    now: datetime,
    date_doc_str: str,
    output_dir: Path,
    run_counter: int = 0,
    save_file: bool = True,
) -> None:
    """Подготавливает атрибуты и генерирует XML-файл в формате для заданного БКИ.

    Args:
        bureau_type (str): Строковой идентификатор БКИ ('okb', 'scoring', 'kbrs', 'nbki').
        data_dict (dict): Словарь с распарсенными данными.
        config (dict): Конфигурация организации.
        now (datetime): Текущие дата и время.
        date_doc_str (str): Строковое представление даты документа.
        output_dir (Path): Директория для сохранения файла.
        run_counter (int, optional): Текущий счетчик запусков. По умолчанию 0.
        save_file (bool, optional): Флаг сохранения файла на диск. По умолчанию True.

    Returns:
        None
    """
    try:
        logger.info(f"Начало генерации XML для {bureau_type.upper()}.")

        org: dict = config.get("organization", {})
        inn: str = org.get("inn", "")
        ogrn: str = org.get("ogrn", "")

        if not inn or not ogrn:
            logger.warning(f"Для {bureau_type.upper()} отсутствуют ИНН ({inn}) или ОГРН ({ogrn}) в конфигурации.")

        subjects_count: str = str(max(0, len(data_dict) - 1))

        attr: dict = {
            "schemaVersion": "4.1",
            "inn": inn,
            "ogrn": ogrn,
            "dateDoc": date_doc_str,
            "subjectsCount": subjects_count,
            "groupBlocksCount": subjects_count,
        }

        if bureau_type == "okb":
            okb_conf: dict = config.get("bureaus", {}).get("okb", {})
            source_id: str = okb_conf.get("sourceID", "02173")
            reg_num: str = f"CHP_{source_id}_EFK_04-10_{now.strftime('%Y%m%d%H%M%S')}000"
            filename: str = f"{reg_num}.xml"
            attr.update(
                {
                    "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
                    "xsi:noNamespaceSchemaLocation": "Main.xsd",
                }
            )

        elif bureau_type == "scoring":
            source_id = "DMH"
            reg_num = str(run_counter)
            filename = f"DMH_FCH_{now.strftime('%Y%m%d')}_{reg_num}.xml"
            logger.trace("Специфичные атрибуты Скоринга сформированы.")

        elif bureau_type == "kbrs":
            source_id = "1136"
            reg_num = f"KBRS_{source_id}_{now.strftime('%Y%m%d')}_{run_counter}"
            filename = f"{reg_num}.xml"
            attr.update(
                {
                    "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
                    "xsi:noNamespaceSchemaLocation": "Main.xsd",
                }
            )
            logger.trace("Специфичные атрибуты КБРС сформированы.")

        elif bureau_type == "nbki":
            source_id = "SJ01SS000001"
            reg_num = f"{source_id}_{now.strftime('%Y%m%d')}_{now.strftime('%H%M%S')}"
            filename = f"{reg_num}.xml"
            logger.trace("Специфичные атрибуты НБКИ сформированы.")

        else:
            logger.error(f"Неподдерживаемый тип БКИ: {bureau_type}. Прерывание генерации.")
            return

        attr["sourceID"] = source_id
        attr["regNumberDoc"] = reg_num
        attr["regNumberDocInaccept"] = reg_num

        finalize_and_save_xml(bureau_type, attr, filename, data_dict, config, date_doc_str, output_dir, save_file)

    except Exception as e:
        logger.exception(f"Критическая ошибка в процессе генерации XML для {bureau_type.upper()}: {e}")


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

    if returns_path.suffix.lower() == ".xls":
        logger.info("Обнаружен формат .xls для файла возвратов, запуск конвертации в .xlsx.")
        returns_path = convert_xls_to_xlsx(returns_path)

    if file_path.suffix.lower() == ".xls":
        logger.info("Обнаружен формат .xls для основного файла, запуск конвертации в .xlsx.")
        file_path = convert_xls_to_xlsx(file_path)

    if not validate_excel(file_path):
        logger.warning(f"Основной файл {file_path.name} не прошел валидацию.")
        return None

    if is_debug:
        logger.debug("Используется тестовый счетчик: 1111")
        run_counter = 1111
    else:
        run_counter = int(config_data.get("run_counter", 0))
        logger.debug(f"Текущий счётчик: {run_counter}")

    logger.success(f"Запуск процесса конвертации.")
    logger.info(f"Основной файл: {file_path}")

    try:
        if returns_path and returns_path.exists():
            logger.info(f"Файл возвратов: {returns_path}.")

            process_excel_returns_file_path = aggregation.process_excel_returns(returns_path, file_path)
            aggregation_file_path = aggregation.process_other_closures(returns_path, process_excel_returns_file_path)

            if not aggregation_file_path:
                logger.critical("Сбой агрегации файла возвратов.")
                return None
        else:
            logger.debug("Файл возвратов не указан или не найден.")
            aggregation_file_path = file_path

        data_dict = excel_parser.parse_active_sheet(aggregation_file_path)
    except Exception as e:
        logger.critical(f"Критическая ошибка при чтении Excel: {e}")
        return None

    if len(data_dict) <= 1:
        logger.warning(f"Файл {file_path.name} не содержит данных для обработки.")
        return None

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
            generate_xml_bureau(
                bureau_type=bki,
                data_dict=data_dict,
                config=config_data,
                now=now,
                date_doc_str=date_doc_str,
                output_dir=output_dir,
                run_counter=run_counter,
                save_file=save_files,
            )
            logger.success(f"Бюро {bki.upper()}: Конвертация успешно завершена.")
        except Exception as e:
            logger.error(f"Ошибка при генерации XML для {bki.upper()}: {e}")

    logger.info("Процесс конвертации завершен.")

    if not is_debug and save_files:
        config_data["run_counter"] = run_counter + 1
        save_config(config_data, config_path)

    return None
