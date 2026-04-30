from __future__ import annotations

import threading
import tkinter.messagebox as messagebox
from pathlib import Path
from tkinter import filedialog
from typing import TYPE_CHECKING

import customtkinter as ctk
from loguru import logger

import convertor
from logger_config import LOG_FORMAT, setup_app_logging

if TYPE_CHECKING:
    from loguru import Record

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

        # === ФАЙЛ ВОЗВРАТОВ ===
        self.lbl_returns = ctk.CTkLabel(self, text="Файл возвратов (*.xls, *.xlsx) (Опционально):")
        self.lbl_returns.grid(row=1, column=0, padx=20, pady=5, sticky="w")

        self.entry_returns = ctk.CTkEntry(self, placeholder_text="Выберите файл возвратов")
        self.entry_returns.grid(row=1, column=1, padx=(0, 20), pady=5, sticky="ew")

        self.btn_browse_returns = ctk.CTkButton(self, text="Обзор", width=100, command=self.browse_returns)
        self.btn_browse_returns.grid(row=1, column=2, padx=(0, 20), pady=5)

        # === JSON ===
        self.lbl_config = ctk.CTkLabel(self, text="Файл конфигурации (*.json):")
        self.lbl_config.grid(row=2, column=0, padx=20, pady=5, sticky="w")

        self.entry_config = ctk.CTkEntry(self)
        self.entry_config.insert(0, "config.json")
        self.entry_config.grid(row=2, column=1, padx=(0, 20), pady=5, sticky="ew")

        self.btn_browse_config = ctk.CTkButton(self, text="Обзор", width=100, command=self.browse_config)
        self.btn_browse_config.grid(row=2, column=2, padx=(0, 20), pady=5)

        # === ПАПКА СОХРАНЕНИЯ ===
        self.lbl_output = ctk.CTkLabel(self, text="Папка сохранения:")
        self.lbl_output.grid(row=3, column=0, padx=20, pady=5, sticky="w")

        self.entry_output = ctk.CTkEntry(self, placeholder_text="По умолчанию: папка с проектом")
        self.entry_output.grid(row=3, column=1, padx=(0, 20), pady=5, sticky="ew")

        self.btn_browse_output = ctk.CTkButton(self, text="Обзор", width=100, command=self.browse_output)
        self.btn_browse_output.grid(row=3, column=2, padx=(0, 20), pady=5)

        # === БКИ ===
        self.frame_bki = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_bki.grid(row=4, column=0, columnspan=3, padx=20, pady=(10, 0), sticky="ew")

        ctk.CTkLabel(self.frame_bki, text="Выбор БКИ:").pack(side="left", padx=(0, 10))

        self.var_okb = ctk.BooleanVar(value=True)
        self.chk_okb = ctk.CTkCheckBox(self.frame_bki, text="ОКБ", variable=self.var_okb)
        self.chk_okb.pack(side="left", padx=(0, 10))

        self.var_scoring = ctk.BooleanVar(value=True)
        self.chk_scoring = ctk.CTkCheckBox(self.frame_bki, text="Скоринг", variable=self.var_scoring)
        self.chk_scoring.pack(side="left", padx=(0, 10))

        self.var_kbrs = ctk.BooleanVar(value=True)
        self.chk_kbrs = ctk.CTkCheckBox(self.frame_bki, text="КБРС", variable=self.var_kbrs)
        self.chk_kbrs.pack(side="left", padx=(0, 10))

        self.var_nbki = ctk.BooleanVar(value=True)
        self.chk_nbki = ctk.CTkCheckBox(self.frame_bki, text="НБКИ", variable=self.var_nbki)
        self.chk_nbki.pack(side="left")

        # === ОПЦИИ ===
        self.frame_options = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_options.grid(row=5, column=0, columnspan=3, padx=20, pady=10, sticky="ew")

        self.var_log_level = ctk.StringVar(value="INFO")
        self.cmb_log_level = ctk.CTkOptionMenu(
            self.frame_options,
            values=["TRACE", "DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
            variable=self.var_log_level,
        )
        self.cmb_log_level.pack(side="left", padx=(0, 20))

        self.var_debug = ctk.BooleanVar(value=False)
        self.chk_debug = ctk.CTkCheckBox(self.frame_options, text="Режим отладки", variable=self.var_debug)
        self.chk_debug.pack(side="left")

        # === ЗАПУСК ===
        self.btn_run = ctk.CTkButton(
            self, text="КОНВЕРТАЦИЯ", height=40, font=("Arial", 14, "bold"), command=self.start_conversion
        )
        self.btn_run.grid(row=6, column=0, columnspan=3, padx=20, pady=10, sticky="ew")

        # === КОНСОЛЬ ЛОГОВ ===
        self.log_textbox = ctk.CTkTextbox(self, state="disabled", wrap="word", font=("Consolas", 12))
        self.log_textbox.grid(row=7, column=0, columnspan=3, padx=20, pady=(10, 20), sticky="nsew")

        def dynamic_gui_filter(record: Record) -> bool:
            """Фильтрует логи для GUI в зависимости от выбранного уровня.

            Args:
                record (Record): Объект записи лога от loguru.

            Returns:
                bool: True, если лог должен быть отображен, иначе False.
            """
            level_map: dict[str, int] = {
                "TRACE": 5,
                "DEBUG": 10,
                "INFO": 20,
                "WARNING": 30,
                "ERROR": 40,
                "CRITICAL": 50,
            }
            current_level: str = self.var_log_level.get()
            min_level: int = level_map.get(current_level, 20)

            log_level_no: int = record["level"].no
            return bool(log_level_no >= min_level)

        gui_sink = TextBoxLogSink(self.log_textbox, self)
        logger.add(
            gui_sink,
            level="TRACE",  # итоговый срез делает фильтр выше
            filter=dynamic_gui_filter,
            colorize=False,
            format=LOG_FORMAT,
        )

    def browse_input(self):
        """Открывает диалоговое окно для выбора входного Excel-файла."""
        filepath = filedialog.askopenfilename(
            title="Выберите файл Excel", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filepath:
            self.entry_input.delete(0, "end")
            self.entry_input.insert(0, filepath)

    def browse_returns(self) -> None:
        """Открывает диалоговое окно для выбора файла возвратов."""
        filepath = filedialog.askopenfilename(
            title="Выберите файл возвратов Excel", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filepath:
            self.entry_returns.delete(0, "end")
            self.entry_returns.insert(0, filepath)

    def browse_output(self):
        """Открывает диалоговое окно для выбора папки сохранения."""
        dirpath = filedialog.askdirectory(title="Выберите папку для сохранения")
        if dirpath:
            self.entry_output.delete(0, "end")
            self.entry_output.insert(0, dirpath)

    def browse_config(self):
        """Открывает диалоговое окно для выбора JSON-файла конфигурации."""
        filepath = filedialog.askopenfilename(
            title="Выберите конфигурационный файл", filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filepath:
            self.entry_config.delete(0, "end")
            self.entry_config.insert(0, filepath)

    def start_conversion(self) -> None:
        """Инициирует процесс конвертации из графического интерфейса.

        Проверяет наличие пути к входному файлу, блокирует кнопку запуска,
        очищает окно логов и запускает процесс конвертации в фоновом потоке.
        """
        input_path = self.entry_input.get().strip()
        config_path = self.entry_config.get().strip()
        output_path = self.entry_output.get().strip() or "."
        returns_path = self.entry_returns.get().strip()
        is_debug = self.var_debug.get()

        bki_list = []
        if self.var_okb.get():
            bki_list.append("okb")
        if self.var_scoring.get():
            bki_list.append("scoring")
        if self.var_kbrs.get():
            bki_list.append("kbrs")
        if self.var_nbki.get():
            bki_list.append("nbki")

        if not input_path:
            logger.error("Попытка запуска без выбора входного файла")
            self.show_log_message("ОШИБКА: Пожалуйста, выберите Excel файл для обработки!\n")
            return

        if not bki_list:
            logger.error("Попытка запуска без выбора БКИ")
            self.show_log_message("ОШИБКА: Выберите минимум один БКИ!\n")
            return

        logger.info("Отправлена команда на конвертацию.")
        self.btn_run.configure(state="disabled", text="Выполнение.")

        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

        thread = threading.Thread(
            target=self.run_process_thread,
            args=(input_path, config_path, output_path, bki_list, is_debug, returns_path),
            daemon=True,
        )
        thread.start()

    def run_process_thread(
        self,
        input_path: str,
        config_path: str,
        output_path: str,
        bki_list: list[str],
        is_debug: bool,
        returns_path: str,
    ) -> None:
        """Выполняет конвертацию файлов в отдельном фоновом потоке.

        Вызывает логику модуля convertor и обрабатывает возможные ошибки с выводом всплывающих окон.

        Args:
            input_path (str): Путь к входному Excel-файлу.
            config_path (str): Путь к JSON-файлу конфигурации.
            is_debug (bool): Флаг режима отладки.
        """
        logger.trace("Фоновый поток конвертации запущен")
        if is_debug:
            logger.debug("Включен режим отладки в GUI.")

        try:
            r_path = Path(returns_path) if returns_path else None
            convertor.run_conversion(
                Path(input_path), Path(config_path), Path(output_path), bki_list, is_debug=is_debug, returns_path=r_path
            )
            message = "Конвертация успешно завершена!\nФайлы сохранены."
            self.after(0, lambda: messagebox.showinfo("Готово", message))

        except Exception as e:
            logger.critical(f"Критический сбой в потоке конвертации: {e}")
            logger.exception(e)

            error_msg = f"Произошла ошибка при конвертации:\n{e}"
            self.after(0, lambda: messagebox.showerror("Ошибка", error_msg))

        finally:
            logger.trace("Фоновый поток завершил работу")
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


def run() -> None:
    """Точка входа для запуска графического интерфейса приложения."""
    setup_app_logging("INFO")
    logger.trace("Инициализация GUI")
    try:
        app = App()
        logger.debug("Главное окно GUI создано")
        app.mainloop()
    except Exception as e:
        logger.critical(f"Приложение GUI аварийно завершилось: {e}")


if __name__ == "__main__":
    run()
