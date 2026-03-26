import sys
import logging

import cli
import gui


logger = logging.getLogger(__name__)


def main():
    """Главная точка входа в приложение.

    Маршрутизирует запуск приложения в зависимости от переданных аргументов.
    Если аргументы командной строки отсутствуют, запускается графический
    интерфейс (GUI). В противном случае запускается консольная версия (CLI).
    """
    if len(sys.argv) == 1:
        logger.info("Выбран GUI режим.")
        gui.run()
    else:
        logger.info("Выбран CLI режим.")
        cli.run()


if __name__ == "__main__":
    main()
