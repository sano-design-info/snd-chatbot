import pytest

from helper.regexpatterns import MSM_ANKEN_NUMBER, INVOICE_DURARION, RANGE_ADDR_PATTERN


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


# generate_invoiceのテスト
def test_match_invoice_durarion():
    assert INVOICE_DURARION.match("納期 1/2") is not None
    assert INVOICE_DURARION.match("納期 10/22") is not None
    # グループ名
    assert INVOICE_DURARION.match("納期 1/2").group("durarion") == "1/2"


# itemparserのテスト
def test_match_range_addr_pattern():
    assert RANGE_ADDR_PATTERN.match("sheetname!A1") is not None
    assert RANGE_ADDR_PATTERN.match("sheetname!A1").groups() == ("sheetname", "A", "1")
    # グループ名
    assert RANGE_ADDR_PATTERN.match("sheetname!A1").group("sheetname") == "sheetname"
    assert RANGE_ADDR_PATTERN.match("sheetname!A1").group("firstcolumn") == "A"
    assert RANGE_ADDR_PATTERN.match("sheetname!A1").group("firstrow") == "1"
