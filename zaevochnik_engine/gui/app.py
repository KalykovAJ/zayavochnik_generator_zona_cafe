"""
Графический интерфейс генератора заявочников.

Раньше программа запускалась только из консоли (run_zaevochnik.py) и
всегда генерировала все заявочники, отмеченные 'Да' в колонке
'Генерировать' на листе 'Задания' файла настроек. Чтобы собрать только
что-то одно, приходилось лезть в Excel и временно менять эту колонку.

Теперь при запуске открывается окно со списком ВСЕХ возможных
заявочников (каждая строка листа 'Задания' - это "категория x сеть"),
где можно отметить галочками нужные комбинации (или нажать "Выбрать
всё" для сразу всех сетей и категорий) и запустить генерацию, видя
прогресс-бар и подробный журнал снизу. Сам файл настроек при этом не
меняется.
"""

import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox

from zaevochnik_engine import settings_locator
from zaevochnik_engine.config_loader import (
    load_settings,
    load_all_job_definitions,
    build_excel_styles,
)
from zaevochnik_engine.pipeline import build_zaevochnik


class _QueueWriter:
    """Подменяет sys.stdout в фоновом потоке генерации: print() внутри
    pipeline.py и других модулей вместо консоли попадает в очередь,
    откуда его подхватывает и показывает в журнале главный поток GUI.

    ВАЖНО: раньше здесь стояла проверка `if message.strip():`, которая
    отбрасывала пустые/из одних пробелов сообщения - а именно таким
    сообщением print() передаёт свой завершающий перевод строки ("\\n"
    пишется ОТДЕЛЬНЫМ вызовом write()). Из-за этого все строки журнала
    склеивались в одну сплошную "простыню" без переносов. Теперь
    ЛЮБОЕ сообщение (включая "\\n") кладётся в очередь как есть.
    """

    def __init__(self, log_queue: "queue.Queue"):
        self.log_queue = log_queue

    def write(self, message):
        if message:
            self.log_queue.put(message)

    def flush(self):
        pass


def _resource_dir() -> str:
    """Папка с ресурсами (иконка и т.п.), с учётом запуска из PyInstaller."""
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class ZaevochnikApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Генератор заявочников")
        self.geometry("820x640")
        self.minsize(680, 500)

        self.log_queue: "queue.Queue" = queue.Queue()
        self.settings_path = None
        self.settings = None
        self.output_folder = None
        self.job_vars = []  # список (tk.BooleanVar, job_dict)
        self.worker_thread = None
        self._log_line_buffer = ""  # буфер незавершённой (без "\n") строки журнала

        self._build_layout()
        self._set_app_icon()

        self.after(100, self._poll_log_queue)
        self.after(150, self._initial_load)

    # ------------------------------------------------------------------
    # Построение интерфейса
    # ------------------------------------------------------------------

    def _set_app_icon(self):
        icon_path = os.path.join(_resource_dir(), "icon", "zaevochnik.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except tk.TclError:
                pass

    def _build_layout(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        self.settings_label_var = tk.StringVar(value="Файл настроек: поиск...")
        ttk.Label(top, textvariable=self.settings_label_var, wraplength=590).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(
            top, text="Изменить файл настроек…", command=self._change_settings_file
        ).pack(side="right")

        mid = ttk.LabelFrame(self, text="Что генерировать (категория товара — сеть)", padding=10)
        mid.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        btn_row = ttk.Frame(mid)
        btn_row.pack(fill="x", pady=(0, 6))
        ttk.Button(btn_row, text="Выбрать всё", command=lambda: self._set_all(True)).pack(side="left")
        ttk.Button(btn_row, text="Снять всё", command=lambda: self._set_all(False)).pack(
            side="left", padx=6
        )
        ttk.Button(btn_row, text="Обновить список", command=self._reload_jobs).pack(
            side="left", padx=6
        )

        canvas_frame = ttk.Frame(mid, relief="sunken", borderwidth=1)
        canvas_frame.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(canvas_frame, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.jobs_frame = ttk.Frame(self.canvas)
        self.jobs_frame.bind(
            "<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.jobs_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Прокрутка колесом мыши только над списком категорий/сетей.
        # РАНЬШЕ колесо было привязано через bind_all() - это глобальная
        # привязка на всё окно, поэтому прокрутка колесом мыши срабатывала
        # даже над журналом внизу и крутила список категорий вместо него.
        # Теперь привязка включается только пока курсор над этим списком
        # (canvas_frame) и снимается, как только курсор уходит с него -
        # в остальное время (например, над журналом) действует обычная
        # прокрутка того виджета, что под курсором.
        canvas_frame.bind("<Enter>", self._activate_jobs_scroll)
        canvas_frame.bind("<Leave>", self._deactivate_jobs_scroll)

        # --- прогресс-бар ---
        progress_frame = ttk.Frame(self, padding=(10, 0, 10, 4))
        progress_frame.pack(fill="x")
        self.progress_style = ttk.Style(self)
        self._set_progress_style("default")
        self.progress = ttk.Progressbar(
            progress_frame, orient="horizontal", mode="determinate",
            style="Zaevochnik.Horizontal.TProgressbar",
        )
        self.progress.pack(side="left", fill="x", expand=True)
        self.progress_label_var = tk.StringVar(value="")
        ttk.Label(progress_frame, textvariable=self.progress_label_var, width=16, anchor="e").pack(
            side="left", padx=(8, 0)
        )

        bottom = ttk.Frame(self, padding=(10, 0, 10, 8))
        bottom.pack(fill="x")
        self.generate_btn = ttk.Button(
            bottom, text="Сгенерировать выбранное", command=self._on_generate_clicked
        )
        self.generate_btn.pack(side="left")
        self.open_folder_btn = ttk.Button(
            bottom,
            text="Открыть папку с заявочниками",
            command=self._open_output_folder,
            state="disabled",
        )
        self.open_folder_btn.pack(side="left", padx=6)
        self.clear_log_btn = ttk.Button(bottom, text="Очистить журнал", command=self._clear_log)
        self.clear_log_btn.pack(side="left", padx=6)
        self.status_var = tk.StringVar(value="")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="right")

        # --- журнал: моноширинный шрифт + цветовые теги для читаемости ---
        log_frame = ttk.LabelFrame(self, text="Журнал", padding=6)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_text = tk.Text(
            log_frame,
            height=14,
            state="disabled",
            wrap="word",
            font=("Consolas", 10),
            background="#fbfbfb",
            borderwidth=0,
        )
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")
        self._configure_log_tags()

    def _configure_log_tags(self):
        # заголовок задания ("--- Генерация: ... ---") - жирным синим,
        # с небольшим отступом сверху, чтобы задания визуально разделялись
        self.log_text.tag_configure(
            "header", foreground="#1c4e8a", font=("Consolas", 10, "bold"), spacing1=6
        )
        self.log_text.tag_configure(
            "success", foreground="#1e8449", font=("Consolas", 10, "bold")
        )
        self.log_text.tag_configure(
            "error", foreground="#c0392b", font=("Consolas", 10, "bold")
        )
        self.log_text.tag_configure("warning", foreground="#b8860b")
        # промежуточные технические шаги пайплайна - мельче и сероватым,
        # с отступом слева, чтобы не спорили за внимание с заголовками/ошибками
        self.log_text.tag_configure(
            "step", foreground="#8a8a8a", lmargin1=18, lmargin2=18
        )
        self.log_text.tag_configure("normal", foreground="#222222")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _activate_jobs_scroll(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _deactivate_jobs_scroll(self, event):
        self.canvas.unbind_all("<MouseWheel>")

    def _set_progress_style(self, kind: str):
        colors = {"default": "#1c62b0", "success": "#1e8449", "warning": "#b8860b"}
        self.progress_style.configure(
            "Zaevochnik.Horizontal.TProgressbar", background=colors.get(kind, colors["default"])
        )

    # ------------------------------------------------------------------
    # Загрузка настроек и заданий
    # ------------------------------------------------------------------

    def _initial_load(self):
        try:
            self.settings_path = settings_locator.find_settings_file()
        except FileNotFoundError as exc:
            messagebox.showerror("Файл настроек не найден", str(exc))
            self.settings_label_var.set("Файл настроек не найден.")
            return
        self.settings_label_var.set(f"Файл настроек: {self.settings_path}")
        self._reload_jobs()

    def _change_settings_file(self):
        """Кнопка 'Изменить файл настроек'. РАНЬШЕ здесь вызывался
        forget_remembered_path() + повторный автопоиск - но если старый
        файл всё ещё лежал там же (папка программы/Рабочий стол), поиск
        молча находил его заново и диалог ни разу не открывался. Теперь
        всегда сразу открывается диалог выбора файла."""
        chosen = settings_locator.browse_for_settings_file()
        if not chosen:
            return  # пользователь закрыл диалог, ничего не выбрав - не трогаем текущий файл
        self.settings_path = chosen
        self.settings_label_var.set(f"Файл настроек: {self.settings_path}")
        self._reload_jobs()

    def _reload_jobs(self):
        if not self.settings_path:
            return
        try:
            self.settings = load_settings(self.settings_path)
            all_jobs = load_all_job_definitions(self.settings_path)
        except Exception as exc:
            messagebox.showerror("Ошибка чтения настроек", str(exc))
            return

        self.output_folder = self.settings["general"].get("Папка для заявочников (путь)")

        for widget in self.jobs_frame.winfo_children():
            widget.destroy()
        self.job_vars = []

        if not all_jobs:
            ttk.Label(self.jobs_frame, text="На листе 'Задания' нет ни одной строки.").pack(
                anchor="w", padx=6, pady=6
            )
            return

        for job in all_jobs:
            var = tk.BooleanVar(value=job["default_checked"])
            label = f"{job['category']}  —  {job['network']}    (файл: {job['output_filename']})"
            cb = ttk.Checkbutton(self.jobs_frame, text=label, variable=var)
            cb.pack(anchor="w", padx=6, pady=2)
            self.job_vars.append((var, job))

    def _set_all(self, value: bool):
        for var, _ in self.job_vars:
            var.set(value)

    # ------------------------------------------------------------------
    # Генерация
    # ------------------------------------------------------------------

    def _on_generate_clicked(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return
        selected = [job for var, job in self.job_vars if var.get()]
        if not selected:
            messagebox.showwarning(
                "Ничего не выбрано", "Отметьте хотя бы один заявочник для генерации."
            )
            return

        self._clear_log()
        self.generate_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="disabled")
        self.status_var.set("Генерация...")
        self._set_progress_style("default")
        self.progress.configure(maximum=len(selected), value=0)
        self.progress_label_var.set(f"0 из {len(selected)}")

        self.worker_thread = threading.Thread(
            target=self._run_generation, args=(selected,), daemon=True
        )
        self.worker_thread.start()

    def _run_generation(self, selected_jobs):
        old_stdout = sys.stdout
        sys.stdout = _QueueWriter(self.log_queue)
        errors = 0
        total = len(selected_jobs)
        try:
            general = self.settings["general"]
            output_folder = general["Папка для заявочников (путь)"]
            os.makedirs(output_folder, exist_ok=True)
            source_file = general["Файл справочника (путь)"]

            for idx, job in enumerate(selected_jobs, start=1):
                category_name = job["category"]
                network_id = job["network"]

                category = self.settings["categories"].get(category_name)
                if category is None:
                    print(f"⚠ Пропуск: категория '{category_name}' не найдена на листе 'Категории'.")
                    errors += 1
                    self.log_queue.put(("__PROGRESS__", idx, total))
                    continue

                network = self.settings["networks"].get(network_id)
                if network is None:
                    print(f"⚠ Пропуск: сеть '{network_id}' не найдена на листе 'Сети'.")
                    errors += 1
                    self.log_queue.put(("__PROGRESS__", idx, total))
                    continue

                output_path = os.path.join(output_folder, job["output_filename"])
                excel_styles = build_excel_styles(general, network, category_name)

                print(f"\n--- Генерация: категория='{category_name}', сеть='{network_id}' ---")
                try:
                    build_zaevochnik(
                        source_path=source_file,
                        output_path=output_path,
                        source_sheet=category_name,
                        sheet_name=category_name,
                        header_row=category["header_row"],
                        filter_col=network["filter_column"],
                        filter_val=[True],
                        columns_to_keep=category["columns_to_keep"],
                        column_mapping=category["column_mapping"],
                        weight_column_name=category["weight_column_name"],
                        excel_styles=excel_styles,
                        validation_prompts=self.settings["validation_prompts"],
                    )
                except Exception as exc:
                    print(f"❌ Ошибка при генерации '{category_name}' / '{network_id}': {exc}")
                    errors += 1

                self.log_queue.put(("__PROGRESS__", idx, total))
        except Exception as exc:
            print(f"❌ Непредвиденная ошибка: {exc}")
            errors += 1
        finally:
            sys.stdout = old_stdout

        self.log_queue.put(("__DONE__", errors))

    # ------------------------------------------------------------------
    # Журнал: буферизация по строкам + цветовая раскраска по содержимому
    # ------------------------------------------------------------------

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self._log_line_buffer = ""

    def _tag_for_line(self, line: str) -> str:
        stripped = line.strip()
        if stripped.startswith("❌"):
            return "error"
        if stripped.startswith("⚠"):
            return "warning"
        if stripped.startswith("🎉"):
            return "success"
        if stripped.startswith("---"):
            return "header"
        if re.match(r"^Шаг\s+\d+", stripped):
            return "step"
        return "normal"

    def _insert_log_line(self, line_with_newline: str):
        tag = self._tag_for_line(line_with_newline)
        self.log_text.configure(state="normal")
        self.log_text.insert("end", line_with_newline, (tag,))
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _feed_log_chunk(self, chunk: str):
        """Накапливает приходящие от print() кусочки текста и вставляет в
        журнал целыми строками (чтобы можно было раскрасить строку целиком
        по её содержимому, а не по случайному фрагменту)."""
        self._log_line_buffer += chunk
        while "\n" in self._log_line_buffer:
            line, self._log_line_buffer = self._log_line_buffer.split("\n", 1)
            self._insert_log_line(line + "\n")

    def _flush_log_buffer(self):
        if self._log_line_buffer:
            self._insert_log_line(self._log_line_buffer)
            self._log_line_buffer = ""

    def _poll_log_queue(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                if isinstance(item, tuple):
                    if item[0] == "__PROGRESS__":
                        _, done, total = item
                        self.progress.configure(value=done)
                        self.progress_label_var.set(f"{done} из {total}")
                    elif item[0] == "__DONE__":
                        errors = item[1]
                        self._flush_log_buffer()
                        self.generate_btn.configure(state="normal")
                        self.open_folder_btn.configure(state="normal")
                        if errors == 0:
                            self.status_var.set("Готово")
                            self._set_progress_style("success")
                            messagebox.showinfo("Готово", "Генерация успешно завершена.")
                        else:
                            self.status_var.set(f"Готово, ошибок: {errors}")
                            self._set_progress_style("warning")
                            messagebox.showwarning(
                                "Готово с ошибками",
                                f"Генерация завершена, но с ошибками: {errors}. "
                                f"Подробности - в журнале (отмечены ❌/⚠).",
                            )
                else:
                    self._feed_log_chunk(item)
        except queue.Empty:
            pass
        self.after(100, self._poll_log_queue)

    def _open_output_folder(self):
        if not self.output_folder or not os.path.isdir(self.output_folder):
            messagebox.showwarning("Папка не найдена", "Папка с заявочниками ещё не создана.")
            return
        if sys.platform == "win32":
            os.startfile(self.output_folder)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", self.output_folder])
        else:
            subprocess.Popen(["xdg-open", self.output_folder])


def main():
    app = ZaevochnikApp()
    app.mainloop()


if __name__ == "__main__":
    main()