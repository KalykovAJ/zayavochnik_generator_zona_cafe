from typing import List, Optional, Union

import pandas as pd


def filter_by_network(
        df: pd.DataFrame,
        filter_column: Optional[Union[str, List[str]]],
        filter_value: Optional[Union[str, List[str]]]
) -> pd.DataFrame:
    """Оставляет только строки, где значение в flag-колонке(ах) сети совпадает
    с filter_value (например, filter_column="БП", filter_value=[True]).

    Если указанной flag-колонки нет в справочнике вообще - это ошибка
    конфигурации (например, в "Настройки.xlsx" в задании указали сеть,
    которой нет на этом листе справочника, как "SP" на листе "Кофе"),
    а не сигнал "не фильтровать". Поэтому в отличие от старой версии
    здесь при отсутствии колонки бросается понятная ошибка, а не тихо
    возвращаются вообще все строки (что раньше могло бы случайно
    включить в заявочник товары чужих сетей).
    """
    if not filter_column or filter_value is None:
        return df

    columns_list = [filter_column] if isinstance(filter_column, str) else filter_column
    allowed_values = [filter_value] if isinstance(filter_value, str) else filter_value

    mask = pd.Series(True, index=df.index)
    for col in columns_list:
        if col not in df.columns:
            raise ValueError(
                f"Колонка-флаг сети '{col}' не найдена в справочнике. "
                f"Проверьте лист 'Сети' в файле настроек - для этой сети "
                f"на данном листе справочника нет такой колонки, значит "
                f"эту сеть нельзя использовать для этой категории товара."
            )
        mask &= df[col].isin(allowed_values)

    return df[mask].copy()
