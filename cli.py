import argparse
import logging
import sys
from pathlib import Path
import openpyxl

def setup_logger(is_verbose: bool):
    """Настройка логирования. Если не verbose, то скрываем всё, кроме критических ошибок."""
    logger = logging.getLogger("AppLogger")
    level = logging.DEBUG if is_verbose else logging.CRITICAL
    
    logger.setLevel(level)
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter('%(levelname)s: %(message)s')
    handler.setFormatter(formatter)
    
    if logger.hasHandlers():
        logger.handlers.clear()
        
    logger.addHandler(handler)
    return logger

def validate_excel(file_path: Path, logger: logging.Logger) -> bool:
    """Проверка существования файла и наличия нужных листов."""
    if not file_path.exists():
        logger.error(f"Файл не найден: {file_path}")
        return False
        
    if file_path.suffix.lower() == '.xls':
        logger.warning("Обнаружено расширение .xls. Пробуем обработать файл как .xlsx")

    try:
        # read_only=True ускоряет чтение, нужны только названия листов
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        wb.close()
    except Exception as e:
        logger.error(f"Не удалось прочитать файл. Убедитесь, что это корректный формат Excel. Ошибка: {e}")
        return False

    required_sheets = {"Активные", "Закрытые"}
    current_sheets = set(sheet_names)
    
    if not required_sheets.issubset(current_sheets):
        logger.error(f"В файле отсутствуют обязательные листы. Ожидалось: {required_sheets}, найдено: {current_sheets}")
        return False
        
    logger.info("Валидация файла прошла успешно. Все необходимые листы присутствуют.")
    return True

def run():
    parser = argparse.ArgumentParser(description="Утилита для обработки Excel файлов.")
    parser.add_argument("file", type=str, help="Путь к файлу xls/xlsx для обработки")
    parser.add_argument("-v", "--verbose", action="store_true", help="Включить подробный вывод логов")
    
    args = parser.parse_args()
    logger = setup_logger(args.verbose)
    
    logger.debug("Запуск в консольном режиме...")
    
    file_path = Path(args.file)
    if not validate_excel(file_path, logger):
        sys.exit(1)
