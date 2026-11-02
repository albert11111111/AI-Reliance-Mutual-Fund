"""按 DataFrame 列名设置 openpyxl 文本/数值格式与列宽，不用 Excel 列字母选业务列。"""

from openpyxl.utils import get_column_letter


def _letter(columns: list[str], name: str) -> str:
    try:
        return get_column_letter(columns.index(name) + 1)
    except ValueError as exc:
        raise KeyError(f"表头没有列: {name}") from exc


def style_sheet(
    ws,
    columns,
    *,
    text_cols=(),
    number_formats=None,
    widths=None,
):
    """冻结首行、开筛选，再按列名设格式。

    columns: 与写入顺序一致的列名（通常 list(df.columns)）
    text_cols: 整列文本格式 @
    number_formats: {列名: 格式串}，只改数据行
    widths: {列名: 宽度}
    """
    cols = list(columns)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for name in text_cols:
        letter = _letter(cols, name)
        for cell in ws[letter]:
            cell.number_format = "@"

    for name, fmt in (number_formats or {}).items():
        letter = _letter(cols, name)
        for cell in ws[letter][1:]:
            cell.number_format = fmt

    for name, width in (widths or {}).items():
        ws.column_dimensions[_letter(cols, name)].width = width
