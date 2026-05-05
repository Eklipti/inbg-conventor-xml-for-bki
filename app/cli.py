import argparse
from pathlib import Path

from loguru import logger

import app.convertor as convertor
from app.logger_config import setup_app_logging


def run():
    """
    Парсит аргументы командной строки и запускает консольную версию приложения.

    Ожидает параметры входного файла, конфигурации, а также флаги отладки
    и подробного вывода. После инициализации передает управление модулю convertor.
    """

    parser = argparse.ArgumentParser(description="Утилита для обработки Excel файлов.")

    parser.add_argument("-j", "--json", type=str, default="config.json", help="Путь к конфигурационному файлу JSON")
    parser.add_argument("-i", "--input", type=str, required=True, help="Путь к файлу xls/xlsx для обработки")
    parser.add_argument(
        "-l",
        "--log-level",
        type=str,
        choices=["TRACE", "DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        default="INFO",
        help="Уровень логирования (по умолчанию INFO)",
    )
    parser.add_argument("-o", "--output", type=str, default=".", help="Путь к директории для сохранения файлов")
    parser.add_argument(
        "-r", "--returns", type=str, default=None, help="Путь к опциональному файлу возвратов (xls/xlsx)"
    )
    parser.add_argument(
        "--bki",
        nargs="*",
        choices=["okb", "scoring", "kbrs", "nbki"],
        help="Список БКИ. Если указать флаг без значений (--bki), процесс пройдет без сохранения файлов.",
    )
    parser.add_argument(
        "-d", "--debug", action="store_true", help="Режим отладки: фиксированный номер 1111, конфиг не обновляется"
    )

    args = parser.parse_args()
    setup_app_logging(args.log_level)

    logger.trace("Используется CLI режим.")
    if args.debug:
        logger.debug("Включен режим отладки.")

    file_path = Path(args.input)
    config_path = Path(args.json)
    output_dir = Path(args.output)
    returns_path = Path(args.returns) if args.returns else None

    try:
        convertor.run_conversion(
            file_path=file_path,
            config_path=config_path,
            output_dir=output_dir,
            bki_list=args.bki,
            is_debug=args.debug,
            returns_path=returns_path,
        )
    except Exception as e:
        logger.error(f"Ошибка в процессе конвертации: {e}")
