"""
Поиск и запоминание пути к файлу "Настройки.xlsx".

Раньше путь к файлу настроек был жёстко зашит в run_zaevochnik.py
(константа SETTINGS_FILE с путём на Рабочий стол одного конкретного
пользователя). На любом другом компьютере или при переносе файла
программа переставала работать, пока кто-то вручную не правил путь
в коде.

Теперь путь ищется автоматически, в следующем порядке:

  1. Место, которое программа запомнила с прошлого раза (файл
     settings_location.json в папке %APPDATA%\\Zaevochnik).
  2. Папка самой программы (там же, где лежит .exe или run_zaevochnik.py).
  3. Стандартные пользовательские папки: Рабочий стол, Документы, Загрузки.
  4. На Windows - корень каждого подключенного диска (C:\\, D:\\, ...).
  5. Если нигде не нашлось - один раз показывается диалоговое окно
     "Укажите файл настроек". Выбранный путь запоминается, и в
     следующий раз поиск с шагов 2-5 уже не понадобится.

ВАЖНО: кнопка "Изменить файл настроек" в интерфейсе НЕ должна просто
забывать запомненный путь и запускать этот же автопоиск заново - если
файл по-прежнему лежит там же (например, в папке программы), поиск
молча найдёт его снова и диалог ни разу не откроется. Поэтому для
случая "пользователь сам хочет сменить файл" есть отдельная функция
browse_for_settings_file(), которая всегда сразу открывает диалог
выбора файла, без автопоиска.
"""

import json
import os
import string
import sys
from typing import Optional

SETTINGS_FILE_NAME = "Настройки.xlsx"
APP_FOLDER_NAME = "Zaevochnik"


def _app_data_dir() -> str:
    """Папка для хранения запомненного пути. На Windows - %APPDATA%\\Zaevochnik,
    на других системах - ~/.zaevochnik (на случай запуска не под Windows)."""
    base = os.getenv("APPDATA")
    if base:
        path = os.path.join(base, APP_FOLDER_NAME)
    else:
        path = os.path.join(os.path.expanduser("~"), f".{APP_FOLDER_NAME.lower()}")
    os.makedirs(path, exist_ok=True)
    return path


def _remember_file_path() -> str:
    return os.path.join(_app_data_dir(), "settings_location.json")


def _load_remembered_path() -> Optional[str]:
    remember_file = _remember_file_path()
    if not os.path.exists(remember_file):
        return None
    try:
        with open(remember_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        path = data.get("settings_path")
        if path and os.path.exists(path):
            return path
    except (json.JSONDecodeError, OSError):
        pass
    return None


def _save_remembered_path(path: str) -> None:
    remember_file = _remember_file_path()
    try:
        with open(remember_file, "w", encoding="utf-8") as f:
            json.dump({"settings_path": path}, f, ensure_ascii=False, indent=2)
    except OSError:
        pass  # не критично - в следующий раз просто заново спросим/найдём


def forget_remembered_path() -> None:
    """Сбрасывает запомненный путь, ничего больше не делая (без диалога и
    без повторного автопоиска). Используется внутренне и для диагностики -
    для смены файла из интерфейса используйте browse_for_settings_file()."""
    remember_file = _remember_file_path()
    if os.path.exists(remember_file):
        os.remove(remember_file)


def _program_folder() -> str:
    """Папка, где лежит запущенный .exe (PyInstaller) или .py файл."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    # settings_locator.py лежит в zaevochnik_engine/, а нужна папка проекта уровнем выше
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _user_shell_folders():
    home = os.path.expanduser("~")
    for name in ("Desktop", "Рабочий стол", "Documents", "Документы", "Downloads", "Загрузки"):
        yield os.path.join(home, name)


def _windows_drive_roots():
    if os.name != "nt":
        return
    for letter in string.ascii_uppercase:
        root = f"{letter}:\\"
        if os.path.exists(root):
            yield root


def _search_candidates():
    yield os.path.join(_program_folder(), SETTINGS_FILE_NAME)

    for folder in _user_shell_folders():
        yield os.path.join(folder, SETTINGS_FILE_NAME)

    for root in _windows_drive_roots():
        yield os.path.join(root, SETTINGS_FILE_NAME)


def _open_file_dialog() -> Optional[str]:
    """Просто открывает диалог 'Открыть файл', без всяких информационных
    окон до него. Используется и автопоиском (после info-подсказки), и
    ручным вызовом browse_for_settings_file() (без неё)."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        return None

    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title=f"Укажите файл {SETTINGS_FILE_NAME}",
        filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")],
    )
    root.destroy()
    return path or None


def _ask_user_for_file_with_notice() -> Optional[str]:
    """Показывает объясняющее окно ('файл не найден автоматически'), а
    затем открывает диалог выбора. Используется только внутри
    find_settings_file(), когда автопоиск ничего не нашёл."""
    try:
        import tkinter as tk
        from tkinter import messagebox
    except ImportError:
        return None

    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(
        "Файл настроек не найден",
        f"Не удалось автоматически найти файл '{SETTINGS_FILE_NAME}'.\n\n"
        "Пожалуйста, укажите его расположение вручную. Это нужно будет "
        "сделать только один раз - программа запомнит выбор."
    )
    root.destroy()
    return _open_file_dialog()


def find_settings_file(interactive: bool = True) -> str:
    """Возвращает путь к файлу настроек, используя автопоиск и
    запоминание, описанные в начале файла.

    interactive=False - не показывать диалоговые окна (например, для
    запуска по расписанию без участия пользователя); если файл не
    найден автоматически, будет поднято FileNotFoundError.
    """
    remembered = _load_remembered_path()
    if remembered:
        return remembered

    for candidate in _search_candidates():
        if os.path.exists(candidate):
            _save_remembered_path(candidate)
            return candidate

    if interactive:
        chosen = _ask_user_for_file_with_notice()
        if chosen and os.path.exists(chosen):
            _save_remembered_path(chosen)
            return chosen

    raise FileNotFoundError(
        f"Файл '{SETTINGS_FILE_NAME}' не найден автоматически, и путь "
        f"не был указан вручную."
    )


def browse_for_settings_file(remember: bool = True) -> Optional[str]:
    """Сразу открывает диалог 'Открыть файл' - БЕЗ автопоиска по папкам и
    БЕЗ проверки запомненного пути. Для этого и существует: пользователь
    явно нажал "Изменить файл настроек" и хочет выбрать файл вручную,
    даже если старый файл настроек всё ещё лежит там же, где лежал.

    Возвращает выбранный путь (и запоминает его, если remember=True),
    либо None, если пользователь закрыл диалог, ничего не выбрав.
    """
    chosen = _open_file_dialog()
    if chosen and os.path.exists(chosen):
        if remember:
            _save_remembered_path(chosen)
        return chosen
    return None