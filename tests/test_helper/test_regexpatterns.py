import pytest

from helper.regexpatterns import *


@pytest.mark.parametrize(
    ("targetstr"),
    [
        ("MA-0000"),
        ("MA-0000-0"),
        ("MA-0000-UPPER"),
        ("MA-0000-LH"),
        ("MA-0000-UPPER-RH"),
    ],
)
def test_match_msm_anken_number(targetstr):
    assert MSM_ANKEN_NUMBER.match(targetstr) is not None


@pytest.mark.parametrize(
    ("targetstr", "expected"),
    [
        ("MA-0000", ("MA-0000", "0000")),
        ("MA-0000-UPPER", ("MA-0000", "0000")),
        ("MA-0000-LH", ("MA-0000", "0000")),
        ("MA-0000-UPPER-RH", ("MA-0000", "0000")),
    ],
)
def test_group_msm_anken_number(targetstr, expected):
    assert MSM_ANKEN_NUMBER.match(targetstr).groups() == expected
