import sys
from pathlib import Path

from loguru import logger

LOG_FORMAT: str = "{time:YYYY-MM-DD HH:mm:ss.SSS} - [{module:^12}] - [{level:^7}] - {message}"

LOG_FORMAT_COLOR: str = (
    "<green>{time:YYYY-MM-DD HH:mm:ss.SSS}</green>"
    " - <black>[{module:^12}]</black>"
    " - <level>[{level:^7}]</level>"
    " - <level>{message}</level>"
)


def setup_app_logging(log_level: str = "INFO") -> None:
    """Глобальная настройка логирования приложения.

    Args:
        log_level (str): Уровень логирования для отображения. По умолчанию "INFO".
    """
    logger.remove()

    # Консольный вывод
    logger.add(
        sys.stdout,
        level=log_level,
        colorize=True,
        format=LOG_FORMAT_COLOR,
    )

    # Файловый вывод (всегда TRACE)
    log_path: Path = Path("logs") / "{time:YYYY-MM-DD_HH-mm-ss}.log"

    logger.add(
        log_path,
        level="TRACE",
        rotation="1 month",
        format=LOG_FORMAT,
        colorize=False,
    )
