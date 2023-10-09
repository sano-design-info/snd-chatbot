# get_billing スクリプトのテスト
import json

from pathlib import Path
import pytest

from task.get_billing import (
    generate_dl_numbers,
    QuoteData,
    generate_billing_info_json,
    today_datetime,
    START_DATE_FORMAT,
    BillingInfo,
    generate_json_mfci_billing_item,
)


# generate_dl_numbers
@pytest.mark.parametrize(
    "dl_numbers, expected",
    [
        ("1,2,3,4", [1, 2, 3, 4]),
        ("1-4", [1, 2, 3, 4]),
        ("1,2,5-9", [1, 2, 5, 6, 7, 8, 9]),
    ],
)
def test_generate_dl_numbers(dl_numbers: str, expected: list[int]):
    assert generate_dl_numbers(dl_numbers) == expected


# BillingTargetQuote のテスト  only_katasikiとdurarionの値を検証する
@pytest.mark.parametrize(
    "quitedata,excepted_only_katasiki,excepted_durarion",
    [(("納期  5/12", 54000, "MA-9901 ガススプリング配管図"), "MA-9901", "5/12")],
)
def test_BillingTargetQuote(
    quitedata: tuple[str, int, str],
    excepted_only_katasiki: str,
    excepted_durarion: str,
):
    target = QuoteData(*quitedata)
    assert target.only_katasiki == excepted_only_katasiki
    assert target.durarion == excepted_durarion


# generate_billing_json_data のテスト
# 生成結果の各種データと期待値のjsonファイルを比較する
@pytest.mark.parametrize(
    "testdata, expected_jsonpath",
    [
        (
            (456000, "2023年05月請求分"),
            "tests/testdata/mfcloudinvoice_billing.json",
        ),
    ],
)
def test_generate_billing_json_data(testdata: tuple, expected_jsonpath: str):
    with open(Path(expected_jsonpath), "r", encoding="utf8") as f:
        excepted_jsondata = json.load(f)

    # 期待値jsonのbilling_dateを今日の日付に変更する
    excepted_jsondata["billing_date"] = today_datetime.strftime(START_DATE_FORMAT)
    result = generate_billing_info_json(
        BillingInfo(testdata[0], "ガススプリング配管図作製費", testdata[1])
    )

    assert result["department_id"] == excepted_jsondata["department_id"]
    assert result["billing_date"] == excepted_jsondata["billing_date"]
    # item側を生成する

    item_result = generate_json_mfci_billing_item(
        BillingInfo(testdata[0], "ガススプリング配管図作製費", testdata[1])
    )

    # items.name
    assert item_result["name"] == excepted_jsondata["items"][0]["name"]
    # items.unit_price
    assert item_result["price"] == excepted_jsondata["items"][0]["price"]


# set_border_style のテスト

# export_list のテスト
