import argparse
import logging
import sys
from pathlib import Path

import convertor


def setup_logger(is_verbose: bool):
    """Настройка логирования. Если не verbose, то скрываем всё, кроме критических ошибок."""
    logger = logging.getLogger("AppLogger")
    level = logging.DEBUG if is_verbose else logging.CRITICAL

    logger.setLevel(level)
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter("%(levelname)s: %(message)s")
    handler.setFormatter(formatter)

    if logger.hasHandlers():
        logger.handlers.clear()

    logger.addHandler(handler)
    return logger


def run():
    parser = argparse.ArgumentParser(description="Утилита для обработки Excel файлов.")

    parser.add_argument("-j", "--json", type=str, default="config.json", help="Путь к конфигурационному файлу JSON")
    parser.add_argument("-i", "--input", type=str, required=True, help="Путь к файлу xls/xlsx для обработки")
    parser.add_argument("-v", "--verbose", action="store_true", help="Включить подробный вывод логов")
    parser.add_argument(
        "-d", "--debug", action="store_true", help="Режим отладки: фиксированный номер 1111, конфиг не обновляется"
    )

    args = parser.parse_args()
    logger = setup_logger(args.verbose)

    logger.debug("Запуск в консольном режиме...")
    if args.debug:
        logger.debug("Включен режим отладки.")

    file_path = Path(args.input)
    config_path = Path(args.json)

    convertor.run_conversion(file_path, config_path, logger, is_debug=args.debug)
