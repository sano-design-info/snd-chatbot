# UNICODE文字の番号として、64からA~ となるので、数値変換変換ができるようにしている
MAGIC_NUMBER = 64


def rowcol_to_a1(row: int, col: int) -> str:
    """Translates a row and column cell address to A1 notation.
    :param row: The row of the cell to be converted.
        Rows start at index 1.
    :type row: int, str
    :param col: The column of the cell to be converted.
        Columns start at index 1.
    :type row: int, str
    :returns: a string containing the cell's coordinates in A1 notation.
    Example:
    >>> rowcol_to_a1(1, 1)
    A1

    # ref: gspread package: https://github.com/burnash/gspread
    # LICENSE: https://github.com/burnash/gspread/blob/master/LICENSE.txt
    """
    row = int(row)
    col = int(col)

    if row < 1 or col < 1:
        raise ValueError(f"The cell label is incorrect. ({row}, {col})")

    div = col
    column_label = ""

    while div:
        (div, mod) = divmod(div, 26)
        if mod == 0:
            mod = 26
            div -= 1
        column_label = chr(mod + MAGIC_NUMBER) + column_label

    label = "{}{}".format(column_label, row)

    return label


def a1_to_rowcol(a1_str: str) -> tuple[int, int]:
    """まだ未実装です
    convert A1 -> (1,1)"""
    return (0, 0)