import os
from typing import List, Union

import pandas as pd


def read_source(file_path: str, sheet_names: Union[str, List[str]]) -> pd.DataFrame:
    """
    Читает указанный лист справочника (обычно один - каждый заявочник строится
    из одного конкретного листа, см. SOURCE_SHEET в run_*.py).

    sheet_names ОБЯЗАТЕЛЕН и не имеет значения по умолчанию: раньше исходник
    читался как pd.read_excel(file_path) без указания листа, что молча брало
    первый лист файла. Пока в справочнике был один лист ("Заморозка") это
    работало случайно правильно, но как только в справочник добавится второй
    лист ("Кофе"), такое чтение начало бы возвращать не те данные для новых
    заявочников по кофе.

    Список листов тоже поддерживается (на случай, если когда-нибудь
    понадобится собрать заявочник сразу из нескольких листов), но текущий
    сценарий другой: заявочник "Зона кафе" (заморозка) и будущий заявочник
    "Кофе" - это два РАЗНЫХ файла запуска с разным SOURCE_SHEET и разным
    OUTPUT_FILE, каждый читает свой один лист.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Файл данных не найден по пути: {file_path}")

    sheets = [sheet_names] if isinstance(sheet_names, str) else list(sheet_names)
    if not sheets:
        raise ValueError("Не указан ни один лист-источник (sheet_names пуст).")

    frames = [pd.read_excel(file_path, sheet_name=sheet) for sheet in sheets]
    return frames[0] if len(frames) == 1 else pd.concat(frames, ignore_index=True)
