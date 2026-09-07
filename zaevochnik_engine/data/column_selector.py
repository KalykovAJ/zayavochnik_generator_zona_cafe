from typing import List

import pandas as pd


def select_final_columns(df: pd.DataFrame, columns_to_keep: List[str]) -> pd.DataFrame:
    """
    Оставляет в DataFrame ТОЛЬКО явно перечисленные колонки — белый список,
    а не чёрный.

    Раньше в run_*.py был список COLUMNS_TO_DROP_FINALLY ("что убрать"). Из-за
    этого при добавлении новой сети в справочник (нового флагового столбца,
    например "АЗС4") пришлось бы вручную дописывать этот новый столбец в
    COLUMNS_TO_DROP_FINALLY КАЖДОГО уже существующего run_*.py — иначе чужой
    флаговый столбец случайно попал бы в заявочник.

    С белым списком ("что оставить") это невозможно в принципе: run_cafe_bp.py
    как перечислял свои 4 нужные колонки, так и продолжит их перечислять,
    и появление любой новой колонки в справочнике (флаг новой сети, служебная
    пометка и т.п.) никак не повлияет на уже готовые файлы запуска.

    Также это автоматически чинит расхождение, которое было в старом коде:
    в комментариях run_cafe_*.py говорилось, что "Код" и "Поставщик" не нужны
    в заявочнике, но по факту COLUMNS_TO_DROP_FINALLY их не удалял. Теперь
    такая колонка либо явно в белом списке (значит нужна), либо нет — молчаливых
    расхождений между комментарием и кодом больше не возникает.
    """
    missing = [col for col in columns_to_keep if col not in df.columns]
    if missing:
        raise ValueError(
            f"Колонки {missing} перечислены в COLUMNS_TO_KEEP, но отсутствуют "
            f"в справочнике-источнике. Проверьте актуальный список столбцов "
            f"листа(ов)-источника или поправьте COLUMNS_TO_KEEP."
        )
    return df[columns_to_keep].copy()


def add_generated_columns(df: pd.DataFrame, qty_col: str, total_col: str, weight_total_col: str) -> pd.DataFrame:
    """Добавляет пустые служебные колонки, которые движок заполнит формулами
    (см. excel/formulas.py) уже после записи листа в Excel."""
    df = df.copy()
    df[qty_col] = ""
    df[total_col] = ""
    df[weight_total_col] = ""
    return df
