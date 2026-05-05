from __future__ import annotations

import json
import threading
import tkinter.messagebox as messagebox
from pathlib import Path
from tkinter import filedialog
from typing import TYPE_CHECKING, Any

import customtkinter as ctk
from loguru import logger

from app import orchestrator
from app.logger_config import LOG_FORMAT, setup_app_logging

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

        # === Корневая система вкладок ===
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.tabview = ctk.CTkTabview(self, command=self.tab_changed)
        self.tabview.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.tab_main = self.tabview.add("Конвертер")
        self.tab_json = self.tabview.add("Редактор кодов")

        # === Конвертер ===
        self.tab_main.grid_columnconfigure(1, weight=1)
        self.tab_main.grid_rowconfigure(7, weight=1)

        # === EXCEL ===
        self.lbl_input = ctk.CTkLabel(self.tab_main, text="Excel файл (*.xls, *.xlsx):")
        self.lbl_input.grid(row=0, column=0, padx=20, pady=(20, 5), sticky="w")

        self.entry_input = ctk.CTkEntry(self.tab_main, placeholder_text="Выберите файл.")
        self.entry_input.grid(row=0, column=1, padx=(0, 20), pady=(20, 5), sticky="ew")

        self.btn_browse_input = ctk.CTkButton(self.tab_main, text="Обзор", width=100, command=self.browse_input)
        self.btn_browse_input.grid(row=0, column=2, padx=(0, 20), pady=(20, 5))

        # === ФАЙЛ ВОЗВРАТОВ ===
        self.lbl_returns = ctk.CTkLabel(self.tab_main, text="Файл возвратов (*.xls, *.xlsx) (Опционально):")
        self.lbl_returns.grid(row=1, column=0, padx=20, pady=5, sticky="w")

        self.entry_returns = ctk.CTkEntry(self.tab_main, placeholder_text="Выберите файл возвратов")
        self.entry_returns.grid(row=1, column=1, padx=(0, 20), pady=5, sticky="ew")

        self.btn_browse_returns = ctk.CTkButton(self.tab_main, text="Обзор", width=100, command=self.browse_returns)
        self.btn_browse_returns.grid(row=1, column=2, padx=(0, 20), pady=5)

        # === JSON ===
        self.lbl_config = ctk.CTkLabel(self.tab_main, text="Файл конфигурации (*.json):")
        self.lbl_config.grid(row=2, column=0, padx=20, pady=5, sticky="w")

        self.entry_config = ctk.CTkEntry(self.tab_main)
        self.entry_config.insert(0, "config.json")
        self.entry_config.grid(row=2, column=1, padx=(0, 20), pady=5, sticky="ew")

        self.btn_browse_config = ctk.CTkButton(self.tab_main, text="Обзор", width=100, command=self.browse_config)
        self.btn_browse_config.grid(row=2, column=2, padx=(0, 20), pady=5)

        # === ПАПКА СОХРАНЕНИЯ ===
        self.lbl_output = ctk.CTkLabel(self.tab_main, text="Папка сохранения:")
        self.lbl_output.grid(row=3, column=0, padx=20, pady=5, sticky="w")

        self.entry_output = ctk.CTkEntry(self.tab_main, placeholder_text="По умолчанию: папка с проектом")
        self.entry_output.grid(row=3, column=1, padx=(0, 20), pady=5, sticky="ew")

        self.btn_browse_output = ctk.CTkButton(self.tab_main, text="Обзор", width=100, command=self.browse_output)
        self.btn_browse_output.grid(row=3, column=2, padx=(0, 20), pady=5)

        # === БКИ ===
        self.frame_bki = ctk.CTkFrame(self.tab_main, fg_color="transparent")
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
        self.frame_options = ctk.CTkFrame(self.tab_main, fg_color="transparent")
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
            self.tab_main, text="КОНВЕРТАЦИЯ", height=40, font=("Arial", 14, "bold"), command=self.start_conversion
        )
        self.btn_run.grid(row=6, column=0, columnspan=3, padx=20, pady=10, sticky="ew")

        # === КОНСОЛЬ ЛОГОВ ===
        self.log_textbox = ctk.CTkTextbox(self.tab_main, state="disabled", wrap="word", font=("Consolas", 12))
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
        logger.add(gui_sink, level="TRACE", filter=dynamic_gui_filter, colorize=False, format=LOG_FORMAT)

        # === Редактор JSON ===
        self.tab_json.grid_columnconfigure(0, weight=1)
        self.tab_json.grid_rowconfigure(2, weight=1)

        self.frame_json_file = ctk.CTkFrame(self.tab_json, fg_color="transparent")
        self.frame_json_file.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        self.frame_json_file.grid_columnconfigure(1, weight=1)

        self.lbl_mapping = ctk.CTkLabel(self.frame_json_file, text="Файл таблицы (*.json):")
        self.lbl_mapping.grid(row=0, column=0, padx=(0, 10))

        self.entry_mapping = ctk.CTkEntry(self.frame_json_file)
        self.entry_mapping.insert(0, "dictionary.json")
        self.entry_mapping.grid(row=0, column=1, sticky="ew", padx=(0, 10))

        self.btn_browse_mapping = ctk.CTkButton(
            self.frame_json_file, text="Обзор", width=100, command=self.browse_mapping
        )
        self.btn_browse_mapping.grid(row=0, column=2)

        self.frame_json_header = ctk.CTkFrame(self.tab_json, fg_color="transparent")
        self.frame_json_header.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        self.frame_json_header.grid_columnconfigure(1, weight=1)

        lbl_code = ctk.CTkLabel(self.frame_json_header, text="Код", font=("Arial", 12, "bold"), width=150)
        lbl_code.grid(row=0, column=0, padx=5, sticky="w")

        lbl_name = ctk.CTkLabel(self.frame_json_header, text="Наименование", font=("Arial", 12, "bold"))
        lbl_name.grid(row=0, column=1, padx=5, sticky="w")

        self.scroll_json = ctk.CTkScrollableFrame(self.tab_json)
        self.scroll_json.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
        self.scroll_json.grid_columnconfigure(1, weight=1)

        self.btn_add_json_row = ctk.CTkButton(self.tab_json, text="Добавить строку", command=self.add_json_row)
        self.btn_add_json_row.grid(row=3, column=0, pady=10)

        self.json_entries = []
        self.json_filepath = ""

    def browse_input(self):
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
            # Обновляем таблицу при выборе нового файла
            if self.tabview.get() == "Редактор JSON":
                self.load_json_data()

    def start_conversion(self):
        input_path = self.entry_input.get().strip()
        config_path = self.entry_config.get().strip()
        output_path = self.entry_output.get().strip() or "."
        returns_path = self.entry_returns.get().strip()
        is_debug = self.var_debug.get()

        bki_list = [
            bki
            for bki, var in zip(
                ["okb", "scoring", "kbrs", "nbki"],
                [self.var_okb, self.var_scoring, self.var_kbrs, self.var_nbki],
                strict=False,
            )
            if var.get()
        ]

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

    def run_process_thread(self, input_path, config_path, output_path, bki_list, is_debug, returns_path):
        logger.trace("Фоновый поток конвертации запущен")
        if is_debug:
            logger.debug("Включен режим отладки в GUI.")

        try:
            r_path = Path(returns_path) if returns_path else None
            orchestrator.run_conversion(
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

    def tab_changed(self) -> None:
        """Обрабатывает переключение вкладок для загрузки данных JSON.

        Если активна вкладка "Редактор кодов", инициирует загрузку данных из файла.
        """
        if self.tabview.get() == "Редактор кодов":
            self.load_json_data()

    def browse_mapping(self) -> None:
        """Открывает диалоговое окно для выбора JSON-файла словаря."""
        filepath = filedialog.askopenfilename(
            title="Выберите файл словаря", filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filepath:
            self.entry_mapping.delete(0, "end")
            self.entry_mapping.insert(0, filepath)
            self.load_json_data()

    def load_json_data(self) -> None:
        """Загружает данные из конфигурационного JSON-файла и строит таблицу интерфейса.

        Удаляет старые виджеты, проверяет существование файла через pathlib,
        автоматически создает файл с референсными данными при его отсутствии
        и обрабатывает возможные ошибки чтения.
        """

        try:
            for widget in self.scroll_json.winfo_children():
                widget.destroy()
            self.json_entries.clear()
        except Exception as e:
            logger.error(f"Ошибка при очистке виджетов: {e}")

        self.json_filepath = self.entry_mapping.get().strip()

        if not self.json_filepath:
            logger.warning("Путь к JSON файлу пуст, загрузка прервана.")
            return

        json_path = Path(self.json_filepath)

        if not json_path.exists():
            logger.info(f"Файл {json_path} не найден. Создан новый.")
            try:
                with json_path.open("w", encoding="utf-8") as f:
                    json.dump({"пример 1": "какой-то текст"}, f, ensure_ascii=False, indent=4)
            except Exception as e:
                logger.error(f"Не удалось создать файл {json_path}: {e}")
                return

        try:
            with json_path.open("r", encoding="utf-8") as f:
                data = json.load(f)
                if not isinstance(data, dict):
                    logger.error("Структура JSON файла не является словарем.")
                    self.show_log_message("ОШИБКА: Структура JSON файла не является словарем.\n")
                    return
        except json.JSONDecodeError as e:
            logger.error(f"Ошибка парсинга JSON: {e}")
            self.show_log_message(f"ОШИБКА: Некорректный формат JSON файла. {e}\n")
            return
        except Exception as e:
            logger.error(f"Непредвиденная ошибка чтения JSON файла '{json_path}': {e}")
            return

        for k, v in data.items():
            self._create_json_row(k, str(v))

    def _create_json_row(self, key_text: str = "", val_text: str = "") -> None:
        """Создает строку с полями ввода для ключа и значения в интерфейсе.

        Args:
            key_text (str, optional): Текст для поля ключа. По умолчанию "".
            val_text (str, optional): Текст для поля значения. По умолчанию "".
        """
        row_idx = len(self.json_entries)

        var_k = ctk.StringVar(value=key_text)
        var_v = ctk.StringVar(value=val_text)

        var_k.trace_add("write", self.save_json)
        var_v.trace_add("write", self.save_json)

        entry_k = ctk.CTkEntry(self.scroll_json, textvariable=var_k, width=150)
        entry_k.grid(row=row_idx, column=0, padx=(0, 5), pady=2, sticky="w")

        entry_v = ctk.CTkEntry(self.scroll_json, textvariable=var_v)
        entry_v.grid(row=row_idx, column=1, padx=5, pady=2, sticky="ew")

        btn_del = ctk.CTkButton(
            self.scroll_json,
            text="X",
            width=30,
            fg_color="#d9534f",
            hover_color="#c9302c",
        )

        btn_del.configure(
            command=lambda k=var_k, v=var_v, ek=entry_k, ev=entry_v, b=btn_del: self._delete_json_row(k, v, ek, ev, b)
        )
        btn_del.grid(row=row_idx, column=2, padx=(5, 0), pady=2)

        self.json_entries.append(
            {"var_k": var_k, "var_v": var_v, "widget_k": entry_k, "widget_v": entry_v, "widget_b": btn_del}
        )

    def _delete_json_row(
        self,
        var_k: ctk.StringVar,
        var_v: ctk.StringVar,
        entry_k: ctk.CTkEntry,
        entry_v: ctk.CTkEntry,
        btn_del: ctk.CTkButton,
    ) -> None:
        """Удаляет строку таблицы из интерфейса и инициирует перезапись JSON.

        Args:
            var_k (ctk.StringVar): Переменная ключа удаляемой строки.
            var_v (ctk.StringVar): Переменная значения удаляемой строки.
            entry_k (ctk.CTkEntry): Виджет поля ввода ключа.
            entry_v (ctk.CTkEntry): Виджет поля ввода значения.
            btn_del (ctk.CTkButton): Виджет кнопки удаления.
        """
        self.json_entries = [row for row in self.json_entries if row["widget_k"] != entry_k]

        try:
            entry_k.destroy()
            entry_v.destroy()
            btn_del.destroy()
        except Exception as e:
            logger.warning(f"Ошибка при уничтожении виджетов: {e}")

        self.save_json()

    def add_json_row(self) -> None:
        """Добавляет новую пустую строку со стандартными значениями в конец списка."""
        try:
            self._create_json_row("новый_код", "новое значение")
            self.save_json()
        except Exception as e:
            logger.error(f"Ошибка при добавлении новой строки: {e}")

    def save_json(self, *args: Any) -> None:
        """Считывает данные из интерфейса и сохраняет их в JSON-файл.

        Игнорирует строки с пустыми ключами для предотвращения повреждения структуры.

        Args:
            *args (Any): Дополнительные аргументы, передаваемые tkinter при вызове через trace_add.
        """
        if not hasattr(self, "json_filepath") or not self.json_filepath:
            logger.debug("Пропуск сохранения: путь к файлу JSON не инициализирован.")
            return

        data: dict[str, str] = {}
        for row in self.json_entries:
            try:
                k = row["var_k"].get().strip()
                v = row["var_v"].get().strip()
                if k:
                    data[k] = v
            except Exception as e:
                logger.warning(f"Ошибка при чтении значений из виджета во время сохранения: {e}")

        json_path = Path(self.json_filepath)
        try:
            with json_path.open("w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            logger.trace(f"JSON файл успешно сохранен: {json_path.name}")
        except PermissionError:
            logger.error(f"Отказано в доступе при попытке сохранить {json_path}")
        except Exception as e:
            logger.critical(f"Критическая ошибка при записи JSON файла '{json_path}': {e}")


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
