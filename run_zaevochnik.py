"""
Единая точка запуска генерации заявочников.

Раньше для каждой сети (МП/БП/ПН) и категории товара (Кофе/Заморозка)
нужен был отдельный файл run_*.py с параметрами, зашитыми в код. Потом
это свели к одному файлу запуска, который читает всё из общего файла
настроек "Настройки.xlsx".

Дальше это было улучшено ещё дважды:

  1. Путь к "Настройки.xlsx" больше не зашит в коде. Раньше здесь была
     константа SETTINGS_FILE с абсолютным путём на Рабочий стол одного
     конкретного пользователя - на любом другом компьютере программа
     не запускалась, пока путь не поправят вручную. Теперь путь ищется
     автоматически (запомненное с прошлого раза место -> папка
     программы -> Рабочий стол/Документы/Загрузки -> корни дисков), а
     если нигде не находится - один раз спрашивается через диалоговое
     окно и запоминается. См. zaevochnik_engine/settings_locator.py.

  2. По умолчанию при запуске открывается окно, где можно отметить
     галочками, какие именно заявочники (категория x сеть) собрать
     сейчас - одну сеть, одну категорию, любой набор или сразу всё -
     не редактируя сам файл настроек. См. zaevochnik_engine/gui/app.py.

"Настройки.xlsx" остаётся ОБЩИМ файлом для этой программы и для
программы сборки шаблонов заказов АЗС (та же самая книга, те же самые
листы "Общие"/"Сети") - его нужно поменять только один раз, и это
применится к обеим программам.

Чтобы добавить новую сеть или новую категорию - открываете
"Настройки.xlsx" и правите нужный лист. Код трогать не нужно.

Запуск:
    python run_zaevochnik.py            - открыть графический интерфейс (обычный способ)
    python run_zaevochnik.py --console  - собрать без окна всё, что отмечено
                                           'Да' на листе 'Задания' (для запуска
                                           по расписанию, напр. Планировщиком заданий Windows)
"""

import os
import sys

from zaevochnik_engine import settings_locator
from zaevochnik_engine.config_loader import load_settings, build_excel_styles
from zaevochnik_engine.pipeline import build_zaevochnik


def run_console():
    """Режим без окна: собирает всё, что отмечено 'Да' на листе 'Задания'.
    Это прежнее (единственное) поведение программы - оставлено для запуска
    по расписанию, когда открывать окно и нажимать кнопки некому."""
    settings_path = settings_locator.find_settings_file()
    settings = load_settings(settings_path)
    general = settings["general"]

    source_file = general["Файл справочника (путь)"]
    output_folder = general["Папка для заявочников (путь)"]
    os.makedirs(output_folder, exist_ok=True)

    jobs = settings["jobs"]
    if not jobs:
        print("В листе 'Задания' нет ни одной строки с 'Генерировать' = Да. Нечего собирать.")
        return

    for job in jobs:
        category_name = job["category"]
        network_id = job["network"]

        category = settings["categories"].get(category_name)
        if category is None:
            print(f"⚠ Пропуск: категория '{category_name}' не найдена на листе 'Категории'.")
            continue

        network = settings["networks"].get(network_id)
        if network is None:
            print(f"⚠ Пропуск: сеть '{network_id}' не найдена на листе 'Сети'.")
            continue

        output_path = os.path.join(output_folder, job["output_filename"])
        excel_styles = build_excel_styles(general, network, category_name)

        print(f"\n--- Генерация заявочника: категория='{category_name}', сеть='{network_id}' ---")
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
                validation_prompts=settings["validation_prompts"],
            )
        except Exception as exc:
            # Одна некорректная строка в "Задания" (например, сеть, которой
            # нет в этой категории справочника) не должна останавливать
            # генерацию остальных заявочников из списка.
            print(f"❌ Ошибка при генерации '{category_name}' / '{network_id}': {exc}")


def main():
    if "--console" in sys.argv:
        run_console()
        return

    from zaevochnik_engine.gui.app import main as run_gui
    run_gui()


if __name__ == "__main__":
    main()
