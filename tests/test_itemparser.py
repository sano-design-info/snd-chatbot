import pytest
from pathlib import Path
import json

from datetime import datetime
import dateutil.tz
from itemparser import *


# convert_gmail_datetimestr
@pytest.mark.parametrize(
    ("datetimestr, expected"),
    [
        (
            "Thu, 25 May 2023 14:00:44 +0900 (JST)",
            datetime(2023, 5, 25, 14, 0, 44, tzinfo=dateutil.tz.gettz("Asia/Tokyo")),
        ),
    ],
)
def test_convert_gmail_datetimestr(datetimestr, expected):
    assert convert_gmail_datetimestr(datetimestr) == expected


# ExpandedMessageItemのテスト
# gmail apiの users.messages.get で取得したメッセージの情報を保持するクラス
# テスト用のjsonファイルを用意して、それを読み込んでテストする
@pytest.mark.parametrize(
    ("jsonfilepath"),
    [
        ("tests/testdata/gmailapi_sample_onlytext.json"),
        ("tests/testdata/gmailapi_sample_html.json"),
    ],
)
def test_ExpandedMessageItem(jsonfilepath):
    jsonfile = Path(jsonfilepath)

    with jsonfile.open(mode="r", encoding="utf-8") as f:
        jsondata = json.load(f)

    item = ExpandedMessageItem(jsondata)

    # ヘッダのチェック。それぞれ属性で分けている
    assert item.id == jsondata["id"]
    assert item.title == jsondata["payload"]["headers"]["Subject"]
    assert item.from_address == jsondata["payload"]["headers"]["From"]
    assert item.to_address == jsondata["payload"]["headers"]["To"]
    assert item.cc_address == jsondata["payload"]["headers"]["Cc"]
    assert item.date == convert_gmail_datetimestr(
        jsondata["payload"]["headers"]["Date"]
    )

    # ペイロードのチェック

    # pass


# RenrakukoumokuInfo


# CsvFileInfo

# EstimateCalcSheetInfo

# MsmAnkenMap

# MsmAnkenMapList

# 関連する関数のテスト
# generate_update_valueranges

#
