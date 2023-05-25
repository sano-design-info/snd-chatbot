import pytest

from helper.rangeconvert import *


@pytest.mark.parametrize(
    ("rowcol", "expected"),
    [
        ((1, 1), "A1"),
        ((1, 2), "B1"),
        ((27, 1), "AA1"),
    ],
)
def test_rowcol_to_a1(rowcol, expected):
    assert rowcol_to_a1(rowcol) == expected


def test_rowcol_to_a1_exception():
    with pytest.raises(ValueError):
        rowcol_to_a1((0, 0))
