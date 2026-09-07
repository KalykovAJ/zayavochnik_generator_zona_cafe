from openpyxl.utils import get_column_letter


def autofit_column_width(worksheet, col_indices, start_row: int, end_row: int,
                         min_width: int = 15, padding: int = 4, max_width: int = None):
    """
    Автоподбор ширины колонок по самому длинному текстовому значению
    (формулы игнорируются, т.к. на глаз они обычно короче итогового
    отображаемого результата).

    Раньше почти идентичный цикл был написан дважды: один раз в конце
    styler.py (для всех колонок таблицы) и второй раз в конце
    header_styler.py (для колонок A/B верхней шапки). Здесь это общая,
    параметризуемая версия обоих случаев.
    """
    for col_idx in col_indices:
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(start_row, end_row + 1):
            cell_val = worksheet.cell(row=row_idx, column=col_idx).value
            if cell_val is not None and not str(cell_val).startswith("="):
                max_len = max(max_len, len(str(cell_val)))

        width = max(max_len + padding, min_width)
        if max_width is not None:
            width = min(width, max_width)

        current_width = worksheet.column_dimensions[col_letter].width or 0
        worksheet.column_dimensions[col_letter].width = max(current_width, width)
