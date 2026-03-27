import argparse
from pathlib import Path

from loguru import logger

import convertor
from logger_config import setup_app_logging


def run():
    """
    Парсит аргументы командной строки и запускает консольную версию приложения.

    Ожидает параметры входного файла, конфигурации, а также флаги отладки
    и подробного вывода. После инициализации передает управление модулю convertor.
    """

    parser = argparse.ArgumentParser(description="Утилита для обработки Excel файлов.")

    parser.add_argument("-j", "--json", type=str, default="config.json", help="Путь к конфигурационному файлу JSON")
    parser.add_argument("-i", "--input", type=str, required=True, help="Путь к файлу xls/xlsx для обработки")
    parser.add_argument("-v", "--verbose", action="store_true", help="Включить подробный вывод логов")
    parser.add_argument(
        "-d", "--debug", action="store_true", help="Режим отладки: фиксированный номер 1111, конфиг не обновляется"
    )

    args = parser.parse_args()
    setup_app_logging(args.verbose)

    logger.trace("Используется CLI режим.")
    if args.debug:
        logger.debug("Включен режим отладки.")

    file_path = Path(args.input)
    config_path = Path(args.json)

    convertor.run_conversion(file_path, config_path, is_debug=args.debug)
