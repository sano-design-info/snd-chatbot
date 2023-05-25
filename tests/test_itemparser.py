import pytest

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
def test_ExpandedMessageItem():
    pass


# RenrakukoumokuInfo


# CsvFileInfo

# EstimateCalcSheetInfo

# MsmAnkenMap

# MsmAnkenMapList

# 関連する関数のテスト
# generate_update_valueranges

#
