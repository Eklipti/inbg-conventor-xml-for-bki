import sys
import threading
import tkinter.messagebox as messagebox
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk
from loguru import logger

import convertor

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class TextBoxLogSink:
    """
    Обработчик логов, который перенаправляет текстовый вывод в виджет CustomTkinter.

    Args:
        textbox (ctk.CTkTextbox): Текстовый виджет для отображения логов.
        app_instance (ctk.CTk): Экземпляр главного окна для потокобезопасного обновления UI.
    """

    def __init__(self, textbox, app_instance):
        self.textbox = textbox
        self.app_instance = app_instance

    def write(self, message):
        """
        Форматирует запись лога и передает её в основной поток для вывода в UI.

        Args:
            message (loguru.Message): Объект сообщения лога.
        """
        msg = str(message)
        self.app_instance.after(0, self._append_text, msg)

    def _append_text(self, msg):
        """Добавляет текст сообщения в виджет. Должно вызываться только в основном потоке.

        Args:
            msg (str): Текст сообщения.
        """
        self.textbox.configure(state="normal")
        self.textbox.insert("end", msg)
        self.textbox.configure(state="disabled")
        self.textbox.see("end")

class App(ctk.CTk):
    """Главное окно графического интерфейса приложения (GUI)."""

    def __init__(self):
        super().__init__()

        self.title("Conventor EFKI")
        self.geometry("700x550")
        self.minsize(600, 500)

        # === Grid ===
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(4, weight=1)

        # === EXCEL ===
        self.lbl_input = ctk.CTkLabel(self, text="Excel файл (*.xls, *.xlsx):")
        self.lbl_input.grid(row=0, column=0, padx=20, pady=(20, 5), sticky="w")

        self.entry_input = ctk.CTkEntry(self, placeholder_text="Выберите файл.")
        self.entry_input.grid(row=0, column=1, padx=(0, 20), pady=(20, 5), sticky="ew")

        self.btn_browse_input = ctk.CTkButton(self, text="Обзор", width=100, command=self.browse_input)
        self.btn_browse_input.grid(row=0, column=2, padx=(0, 20), pady=(20, 5))

        # === JSON ===
        self.lbl_config = ctk.CTkLabel(self, text="Файл конфигурации (*.json):")
        self.lbl_config.grid(row=1, column=0, padx=20, pady=5, sticky="w")

        self.entry_config = ctk.CTkEntry(self)
        self.entry_config.insert(0, "config.json")
        self.entry_config.grid(row=1, column=1, padx=(0, 20), pady=5, sticky="ew")

        self.btn_browse_config = ctk.CTkButton(self, text="Обзор", width=100, command=self.browse_config)
        self.btn_browse_config.grid(row=1, column=2, padx=(0, 20), pady=5)

        # === ОПЦИИ ===
        self.frame_options = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_options.grid(row=2, column=0, columnspan=3, padx=20, pady=10, sticky="ew")

        self.var_verbose = ctk.BooleanVar(value=True)
        self.chk_verbose = ctk.CTkCheckBox(
            self.frame_options, text="Подробные логи (Verbose)", variable=self.var_verbose
        )
        self.chk_verbose.pack(side="left", padx=(0, 20))

        self.var_debug = ctk.BooleanVar(value=False)
        self.chk_debug = ctk.CTkCheckBox(self.frame_options, text="Режим отладки (Debug)", variable=self.var_debug)
        self.chk_debug.pack(side="left")

        # === ЗАПУСК ===
        self.btn_run = ctk.CTkButton(
            self, text="КОНВЕРТАЦИЯ", height=40, font=("Arial", 14, "bold"), command=self.start_conversion
        )
        self.btn_run.grid(row=3, column=0, columnspan=3, padx=20, pady=10, sticky="ew")

        # === КОНСОЛЬ ЛОГОВ ===
        self.log_textbox = ctk.CTkTextbox(self, state="disabled", wrap="word", font=("Consolas", 12))
        self.log_textbox.grid(row=4, column=0, columnspan=3, padx=20, pady=(10, 20), sticky="nsew")

        # Для замены консольного логгера по умолчанию
        logger.remove()

        logger.add(
            sys.stderr,
            level="ERROR",
            format="<red>{time:HH:mm:ss} - [{module:^12}] - [{level:^7}] - {message}</red>"
        )

        def dynamic_gui_filter(record):
            if self.var_verbose.get():
                return True
            return record["level"].no >= 20

        gui_sink = TextBoxLogSink(self.log_textbox, self)
        logger.add(
            gui_sink,
            level="DEBUG",
            filter=dynamic_gui_filter,
            colorize=False,
            format="{time:HH:mm:ss} - [{module:^12}] - [{level:^7}] - {message}",
        )

    def browse_input(self):
        """Открывает диалоговое окно для выбора входного Excel-файла."""
        filepath = filedialog.askopenfilename(
            title="Выберите файл Excel", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filepath:
            self.entry_input.delete(0, "end")
            self.entry_input.insert(0, filepath)

    def browse_config(self):
        """Открывает диалоговое окно для выбора JSON-файла конфигурации."""
        filepath = filedialog.askopenfilename(
            title="Выберите конфигурационный файл", filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filepath:
            self.entry_config.delete(0, "end")
            self.entry_config.insert(0, filepath)

    def start_conversion(self):
        """Инициирует процесс конвертации из графического интерфейса.

        Проверяет наличие пути к входному файлу, блокирует кнопку запуска,
        очищает окно логов и запускает процесс конвертации в фоновом потоке.
        """
        input_path = self.entry_input.get().strip()
        config_path = self.entry_config.get().strip()
        is_debug = self.var_debug.get()

        if not input_path:
            self.show_log_message("ОШИБКА: Пожалуйста, выберите Excel файл для обработки!\n")
            return

        self.btn_run.configure(state="disabled", text="Выполнение.")

        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

        thread = threading.Thread(
            target=self.run_process_thread, args=(input_path, config_path, is_debug), daemon=True
        )
        thread.start()

    def run_process_thread(self, input_path, config_path, is_debug):
        """Выполняет конвертацию файлов в отдельном фоновом потоке.

        Вызывает логику модуля convertor и обрабатывает возможные ошибки с выводом всплывающих окон.

        Args:
            input_path (str): Путь к входному Excel-файлу.
            config_path (str): Путь к JSON-файлу конфигурации.
            is_debug (bool): Флаг режима отладки.
        """
        if is_debug:
            logger.debug("Включен режим отладки.")

        try:
            convertor.run_conversion(Path(input_path), Path(config_path), is_debug=is_debug)
            message = "Конвертация успешно завершена!\nФайлы сохранены в папке с программой"
            self.after(0, lambda: messagebox.showinfo("Готово", message))

        except Exception as e:
            logger.critical(f"Непредвиденная ошибка в процессе конвертации: {e}")
            import traceback

            logger.debug(traceback.format_exc())

            error_msg = f"Произошла ошибка при конвертации:\n{e}"
            self.after(0, lambda: messagebox.showerror("Ошибка", error_msg))

        finally:
            self.after(0, lambda: self.btn_run.configure(state="normal", text="КОНВЕРТАЦИЯ"))

    def show_log_message(self, msg):
        """Выводит текстовое сообщение напрямую в консоль логов GUI.

        Args:
            msg (str): Текст сообщения.
        """
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", msg)
        self.log_textbox.configure(state="disabled")
        self.log_textbox.see("end")


def run():
    """Точка входа для запуска графического интерфейса приложения."""
    logger.trace("Используется GUI режим.")
    app = App()
    app.mainloop()


if __name__ == "__main__":
    run()
