import logging
from pathlib import Path
import excel_parser

def run_conversion(file_path: Path, logger: logging.Logger):
    logger.debug("Запуск основного процесса конвертации...")

    data_dict = excel_parser.parse_active_sheet(file_path, logger)
    sample_data = dict(list(data_dict.items())[:5])
    logger.debug(f"Тестовый вывод собранных данных (первые 5 элементов): {sample_data}")
    
    # TODO: Здесь будет происходить дальнейшая работа с data_dict и добавление новой логики