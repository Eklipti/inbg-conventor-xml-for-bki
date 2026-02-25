import logging
from pathlib import Path
from cli import setup_logger

def run():
    # В будущем здесь будет инициализация Tkinter, PyQt или другого фреймворка
    print("Открыто графическое окно (GUI Mode).")
    
    # Эмуляция: пользователь выбрал файл и поставил галочку "Режим отладки"
    # test_file_path = Path("test.xlsx")
    # debug_checkbox_checked = True 
    
    # Использование логики:
    # logger = setup_logger(debug_checkbox_checked)
    # is_valid = validate_excel(test_file_path, logger)
    # print(f"Результат валидации: {is_valid}")