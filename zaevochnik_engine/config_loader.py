"""
Загрузчик настроек из Excel-файла "Настройки_заявочников.xlsx".

Раньше, чтобы добавить новую сеть АЗС или новую категорию товара,
нужно было создавать новый файл запуска (run_*.py) и редактировать код
на Python. Теперь ВСЯ такая информация - какие сети существуют, какие
у них фирменные цвета, какие категории товара есть в справочнике, и
какие заявочники вообще нужно генерировать - хранится в обычном Excel-
файле, который может редактировать любой пользователь, не трогая код.

Файл настроек содержит 5 листов:
  - "Общие"     - общие параметры (путь к справочнику, куда сохранять
                  заявочники, шрифты/цвета по умолчанию и т.п.);
  - "Подсказки" - тексты всплывающих подсказок при вводе количества;
  - "Категории" - список категорий товара (= листов справочника), их
                  структура колонок;
  - "Сети"      - список сетей АЗС, их фирменные цвета и то, как их
                  найти в справочнике (флаговая колонка);
  - "Задания"   - какие заявочники (категория x сеть) реально нужно
                  собирать и куда сохранять готовый файл.
"""

from typing import Optional

import pandas as pd


def _clean(value):
    """Возвращает None вместо NaN/пустых значений ячеек Excel."""
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    text = str(value).strip()
    return text if text else None


def _pick(value, default):
    """value, если оно задано в Excel, иначе default (общее значение)."""
    cleaned = _clean(value)
    return cleaned if cleaned is not None else default


def _row_to_dict(df: pd.DataFrame) -> dict:
    """Лист из двух колонок (Параметр | Значение) -> обычный словарь."""
    result = {}
    for _, row in df.iterrows():
        key = _clean(row.iloc[0])
        if key is None:
            continue
        result[key] = row.iloc[1]
    return result


def _split_columns_list(raw: str) -> list:
    return [part.strip() for part in str(raw).split(",") if part.strip()]


def load_general(settings_path: str) -> dict:
    df = pd.read_excel(settings_path, sheet_name="Общие", header=None)
    return _row_to_dict(df)


def load_validation_prompts(settings_path: str) -> dict:
    df = pd.read_excel(settings_path, sheet_name="Подсказки")
    prompts = {}
    for _, row in df.iterrows():
        key = _clean(row["Тип (unit/pack)"])
        if key is None:
            continue
        prompts[key] = {
            "title": row["Заголовок"],
            "message": row["Сообщение"],
        }
    return prompts


def load_categories(settings_path: str) -> dict:
    df = pd.read_excel(settings_path, sheet_name="Категории")
    categories = {}
    for _, row in df.iterrows():
        name = _clean(row["Категория (лист справочника)"])
        # Строки-примечания под таблицей (или случайно пустые строки) не
        # содержат корректный номер строки шапки - пропускаем их, а не
        # падаем с ошибкой.
        if name is None or _clean(row["Строка шапки таблицы"]) is None:
            continue
        categories[name] = {
            "header_row": int(row["Строка шапки таблицы"]),
            "columns_to_keep": _split_columns_list(row["Колонки для заявочника (через запятую)"]),
            "column_mapping": {
                "type_col": row["Колонка типа заказа"],
                "qty_col": row["Колонка \"Кол-во заказа\""],
                "mult_col": row["Колонка мультипликатора упаковки"],
                "total_col": row["Колонка \"Итого (шт)\""],
                "weight_col": row["Колонка веса единицы"],
            },
            "weight_column_name": row["Колонка \"Итого (вес)\""],
        }
    return categories


def load_networks(settings_path: str, general: dict) -> dict:
    df = pd.read_excel(settings_path, sheet_name="Сети")
    networks = {}
    for _, row in df.iterrows():
        net_id = _clean(row["Сеть (ID)"])
        # Строки-примечания под таблицей не содержат название бренда -
        # пропускаем их, а не падаем с ошибкой.
        if net_id is None or _clean(row.get("Название сети (бренд)")) is None:
            continue
        networks[net_id] = {
            "brand_name": _pick(row["Название сети (бренд)"], net_id),
            "filter_column": net_id,
            "primary": _pick(row.get("Цвет заголовка таблицы (HEX)"), general.get("Цвет заголовка по умолчанию (HEX)")),
            "bg_pack": _pick(row.get("Цвет фона \"Упаковка\" (HEX)"), general.get("Цвет фона \"Упаковка\" по умолчанию (HEX)")),
            "header_font_color": _pick(row.get("Цвет текста заголовка таблицы (HEX)"), general.get("Цвет текста заголовка по умолчанию (HEX)")),
            "top_header_bg": _pick(row.get("Цвет фона верхней шапки (HEX)"), general.get("Цвет фона верхней шапки по умолчанию (HEX)")),
            "text_yellow": _pick(row.get("Цвет жёлтого текста (HEX)"), general.get("Цвет жёлтого текста по умолчанию (HEX)")),
            "text_white": _pick(row.get("Цвет белого текста (HEX)"), general.get("Цвет белого текста по умолчанию (HEX)")),
            "fuel_station": _pick(row.get("Ярлык номера точки"), None),
        }
        if networks[net_id]["fuel_station"] is None:
            raise ValueError(
                f"Для сети '{net_id}' на листе 'Сети' не заполнена колонка "
                f"'Ярлык номера точки' (например, 'АЗС №' или 'MARKET №'). "
                f"Раньше у этого параметра было значение по умолчанию с листа "
                f"'Общие', теперь его нужно задавать для каждой сети отдельно."
            )
    return networks


def load_jobs(settings_path: str) -> list:
    df = pd.read_excel(settings_path, sheet_name="Задания")
    jobs = []
    for _, row in df.iterrows():
        category = _clean(row["Категория"])
        network = _clean(row["Сеть"])
        enabled = _clean(row["Генерировать (Да/Нет)"])
        if category is None or network is None:
            continue
        if not enabled or enabled.strip().lower() not in ("да", "yes", "true", "1"):
            continue
        jobs.append({
            "category": category,
            "network": network,
            "output_filename": _clean(row["Имя файла вывода"]) or f"Заявочник {category} {network}.xlsx",
        })
    return jobs


def load_all_job_definitions(settings_path: str) -> list:
    """Все строки листа 'Задания', включая те, где 'Генерировать' = Нет.

    load_jobs() (выше) нужен консольному режиму - он молча берёт только
    включённые строки, как и раньше. Этой функции пользуется графический
    интерфейс: пользователь должен видеть ВСЕ возможные заявочники
    (категория x сеть) и сам отмечать галочками, что собрать именно
    сейчас, не редактируя сам файл настроек. Колонка 'Генерировать
    (Да/Нет)' в этом случае используется только как исходное состояние
    галочки при открытии программы.
    """
    df = pd.read_excel(settings_path, sheet_name="Задания")
    jobs = []
    for _, row in df.iterrows():
        category = _clean(row["Категория"])
        network = _clean(row["Сеть"])
        if category is None or network is None:
            continue
        enabled = _clean(row["Генерировать (Да/Нет)"])
        default_checked = bool(enabled) and enabled.strip().lower() in ("да", "yes", "true", "1")
        jobs.append({
            "category": category,
            "network": network,
            "output_filename": _clean(row["Имя файла вывода"]) or f"Заявочник {category} {network}.xlsx",
            "default_checked": default_checked,
        })
    return jobs


def load_settings(settings_path: str) -> dict:
    general = load_general(settings_path)
    return {
        "general": general,
        "validation_prompts": load_validation_prompts(settings_path),
        "categories": load_categories(settings_path),
        "networks": load_networks(settings_path, general),
        "jobs": load_jobs(settings_path),
    }


def build_excel_styles(general: dict, network: dict, category_name: str) -> dict:
    """Собирает EXCEL_STYLES для конкретной пары (сеть, категория) —
    аналог старого build_excel_styles() из config.py, но данные вместо
    Python-констант берутся из файла настроек."""
    return {
        "row_heights": {
            "header": int(general["Высота строки шапки таблицы (px)"]),
            "data": int(general["Высота строки данных (px)"]),
        },
        "colors": {
            "primary": network["primary"],
            "bg_pack": network["bg_pack"],
            "border": general["Цвет рамки таблицы (HEX)"],
        },
        "fonts": {
            "header": {
                "name": general["Шрифт (обычный)"],
                "size": int(general["Размер шрифта (обычный)"]),
                "bold": True,
                "color": network["header_font_color"],
            },
            "regular": {"name": general["Шрифт (обычный)"], "size": int(general["Размер шрифта (обычный)"]), "bold": False},
            "bold": {"name": general["Шрифт (обычный)"], "size": int(general["Размер шрифта (обычный)"]), "bold": True},
        },
        "alignments": {
            "header": {"horizontal": "center", "vertical": "center", "wrap_text": True},
            "default": {"horizontal": "center", "vertical": "center"},
            "text_left": {"horizontal": "left", "vertical": "center"},
        },
        "top_header": {
            "font_name": general["Шрифт шапки верхнего блока (бренд)"],
            "border": general["Цвет рамки таблицы (HEX)"],
            "company_font_size": int(general["Размер шрифта названия компании"]),
            "label_font_size": int(general["Размер шрифта подписей шапки"]),
            "sheet_password": str(general["Пароль защиты листа"]),
            "bg_color": network["top_header_bg"],
            "text_color_yellow": network["text_yellow"],
            "text_color_white": network["text_white"],
            "company_name": f"{network['brand_name']} {category_name}",
            "fuel_station": network["fuel_station"],
        },
    }