import logging
import sys


def setup_app_logging(is_verbose: bool = False):
    """Глобальная настройка логирования приложения."""

    level = (
        logging.DEBUG if is_verbose else logging.INFO
    )  # В gui.py у вас был INFO, можно поставить CRITICAL как было в cli.py

    formatter = logging.Formatter(
        fmt="%(asctime)s - [%(levelname)s] - [%(module)s] - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )

    stdout_handler = logging.StreamHandler(sys.stdout)
    stdout_handler.setFormatter(formatter)

    root_logger = logging.getLogger()
    root_logger.setLevel(level)

    if root_logger.hasHandlers():
        root_logger.handlers.clear()

    root_logger.addHandler(stdout_handler)
