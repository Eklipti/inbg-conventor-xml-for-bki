import sys

from loguru import logger

from app import cli, gui


def main():
    """Главная точка входа в приложение.

    Маршрутизирует запуск приложения в зависимости от переданных аргументов.
    Если аргументы командной строки отсутствуют, запускается графический
    интерфейс (GUI). В противном случае запускается консольная версия (CLI).
    """
    logger.trace("Запуск приложения")
    try:
        if len(sys.argv) == 1:
            logger.trace("Аргументы не переданы. Запуск GUI режима.")
            gui.run()
        else:
            logger.trace(f"Переданы аргументы: {sys.argv[1:]}. Запуск CLI режима.")
            cli.run()
    except Exception as e:
        logger.critical(f"Критическая ошибка при работе приложения: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
