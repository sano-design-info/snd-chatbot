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
        ("MA-0000", ("MA-0000", "0000", "", None, None, None)),
        ("MA-0000-UPPER", ("MA-0000", "0000", "UPPER", "UPPER", None, None)),
        ("MA-0000-LH", ("MA-0000", "0000", "LH", None, "LH", None)),
        ("MA-0000-UPPER-RH", ("MA-0000", "0000", "UPPER-RH", "UPPER", "RH", None)),
        ("MA-0000-1", ("MA-0000", "0000", "1", None, None, "1")),
        ("MA-0000-UPPER-RH-1", ("MA-0000", "0000", "UPPER-RH-1", "UPPER", "RH", "1")),
    ],
)
def test_group_msm_anken_number(targetstr, expected):
    assert MSM_ANKEN_NUMBER.match(targetstr).groups() == expected

# get_billingのテスト
def test_match_billing_durarion():
    assert BILLING_DURARION.match("納期 1/2") is not None
    assert BILLING_DURARION.match("納期 10/22") is not None
    # グループ名
    assert BILLING_DURARION.match("納期 1/2").group("durarion") == "1/2"


# itemparserのテスト
