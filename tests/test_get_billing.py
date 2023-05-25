# get_billing スクリプトのテスト

from pathlib import Path
import pytest

from get_billing import *


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
    target = BillingTargetQuote(*quitedata)
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
    excepted_jsondata["billing"]["billing_date"] = today_datetime.strftime(
        START_DATE_FORMAT
    )
    result = generate_billing_json_data(
        BillingData(testdata[0], "ガススプリング配管図作製費", testdata[1])
    )

    # 検証をする
    # TODO:2023-05-25 この検証方法が微妙に感じるので、もっと良い方法を考える
    # department_id
    assert (
        result["billing"]["department_id"]
        == excepted_jsondata["billing"]["department_id"]
    )
    # items.name
    assert (
        result["billing"]["items"][0]["name"]
        == excepted_jsondata["billing"]["items"][0]["name"]
    )
    # items.unit_price
    assert (
        result["billing"]["items"][0]["unit_price"]
        == excepted_jsondata["billing"]["items"][0]["unit_price"]
    )


# set_border_style のテスト

# export_list のテスト
