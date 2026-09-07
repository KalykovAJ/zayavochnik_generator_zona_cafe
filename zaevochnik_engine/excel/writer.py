from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd


def write_dataframe(worksheet, df: pd.DataFrame, start_row: int) -> int:
    """
    Построчно записывает DataFrame в лист, начиная со start_row (первой идёт
    строка заголовков). Возвращает номер последней строки с данными —
    он нужен дальше формулам, стилям и авто-фильтру.
    """
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=start_row):
        for c_idx, value in enumerate(row, start=1):
            worksheet.cell(row=r_idx, column=c_idx, value=value)
    return start_row + len(df)
