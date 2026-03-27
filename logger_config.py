import sys

from loguru import logger


def setup_app_logging(is_verbose: bool = False):
    """Глобальная настройка логирования приложения."""

    level = "TRACE" if is_verbose else "INFO"

    logger.remove()

    logger.add(
        sys.stdout,
        level=level,
        format="{time:YYYY-MM-DD HH:mm:ss} - [{level}] - [{module}] - {message}"
    )