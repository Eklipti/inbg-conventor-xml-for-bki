import sys

from loguru import logger


def setup_app_logging(is_verbose: bool = False):
    """Глобальная настройка логирования приложения."""

    level = "TRACE" if is_verbose else "INFO"

    logger.remove()

    logger.add(
        sys.stdout,
        level=level,
        colorize=True,
        format="<green>{time:YYYY-MM-DD HH:mm:ss}</green>"
               " - <black>[{module:^12}]</black>"
               " - <level>[{level:^7}]</level>"
               " - <level>{message}</level>"
    )