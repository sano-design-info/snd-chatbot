import pytest
from pathlib import Path
import json

from datetime import datetime
import dateutil.tz
import helper
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
    ("jsonfilepath", "expected"),
    [
        (
            "tests/testdata/gmailapi_sample_onlytext.json",
            {
                "id": "187b63167c05df9c",
                "title": "見積作成 メールスレッド登録テスト2 MA-0000",
                "from_address": '"佐野設計事務所 佐野浩士" \u003chiroshi.sano@sano-design.info\u003e',
                "to_address": "hiroshi.sano+test01@sano-design.info",
                "cc_address": "hiroshi.sano+test02@sano-design.info",
                # Tue, 25 Apr 2023 11:15:02 +0900
                "datetime": datetime(
                    2023, 4, 25, 11, 15, 2, tzinfo=dateutil.tz.gettz("Asia/Tokyo")
                ),
                "body": "44GT44Gh44KJ44Gu44Oh44O844Or44Gv6KaL56mN5L2c5oiQ44Go44Oh44O844Or44K544Os44OD44OJ55m76Yyy44OG44K544OI44Gn44GZ44CCMuWbnuebruOAgkND5YWl44KKDQo9PT09PT09PT09PT09PT09PT09PQ0K5qCq5byP5Lya56S-5L2Q6YeO6Kit6KiI5LqL5YuZ5omADQrkvZDph47jgIDmtanlo6sNCmVtYWlsIDogaGlyb3NoaS5zYW5vQHNhbm8tZGVzaWduLmluZm8NCldlYuOCteOCpOODiCA6IGh0dHA6Ly9zYW5vLWRlc2lnbi5pbmZvLw0KVEVMIDogMDU0NS01NS0xMTY2DQo9PT09PT09PT09PT09PT09PT09PQ0K",
            },
        ),
        (
            "tests/testdata/gmailapi_sample_html.json",
            {
                "id": "1866d9ce947765f4",
                "title": "【配管図作成依頼】　MA-0947 TP72(36)新生工業様　　　 【配管図依頼_依頼日230220_721199新生工業様_TP72　",
                "from_address": '"OKAZAKI, Chiharu (岡崎 千治)" \u003cchiharu.pzv1.okazaki@misumi.co.jp\u003e',
                "to_address": '"hiroshi.sano" \u003chiroshi.sano@sano-design.info\u003e',
                "cc_address": '"yuuki.sano" \u003cyuuki.sano@sano-design.info\u003e, "snd-team@sano-design.info" \u003csnd-team@sano-design.info\u003e, "gas-linked@ml.misumi.co.jp" \u003cgas-linked@ml.misumi.co.jp\u003e, "西川　勇" \u003cisamu.n3wx.nishikawa@misumi.co.jp\u003e, "渡邊　容子" \u003cyo.watanabe@misumi.co.jp\u003e, "福永　望" \u003cn.fukunaga@misumi.co.jp\u003e',
                # "Mon, 20 Feb 2023 06:57:15 +0000" #JSTへ変換 -> Mon, 20 Feb 2023 15:57:15 +0900
                "datetime": datetime(
                    2023, 2, 20, 15, 57, 15, tzinfo=dateutil.tz.gettz("Asia/Tokyo")
                ),
                "body": "5qCq5byP5Lya56S-5L2Q6YeO6Kit6KiI5LqL5YuZ5omADQrkvZDph47mp5gNCg0KDQrjgYrkuJboqbHjgavjgarjgaPjgabjgYrjgorjgb7jgZnjgIINCuODn-OCueODn-OBruWyoeW0juOBp-OBmeOAgg0KDQoNCjIvMjDmlrDnlJ_lt6Xmpa3mp5jjgYvjgonjga7phY3nrqHlm7PkvZzmiJDkvp3poLzjgafjgZnjgILvvIgxLzTvvIkNCg0KDQoNCg0KDQrjgZTlr77lv5zjgpLjgYrpoZjjgYTjgZfjgb7jgZnjgIINCg0KDQoNCuODu-OAjE1BLSAwOTQ344CAIFRQNzIoMzYp44CNDQoNCuW4jOacm-e0jeacn--8mjIvMjLvvIjmsLTvvIkNCg0KDQoNCg0KDQrku6XkuIrjgIHjgYrmiYvmlbDjgYrjgYvjgZHoh7TjgZfjgb7jgZnjgIINCg0K44KI44KN44GX44GP44GK6aGY44GE6Ie044GX44G-44GZ44CCDQoNCg0KKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqDQoNCuagquW8j-S8muekvuODn-OCueODnw0KDQrml6XmnKzkvIHmpa3kvZPjgIDph5Hlnovkuovmpa3jgrDjg6vjg7zjg5cNCg0K6YeR5Z6L5LqL5qWt6YOo44CA44OX44Os44K55qmf5bel5ZOB5LqL5qWt44OB44O844OgDQoNCuWyoeW0juOAgOWNg-ayuw0KDQoNCuOAkjEwMi04NTgz44CA5p2x5Lqs6YO95Y2D5Luj55Sw5Yy65Lmd5q615Y2XMS02LTXkuZ3mrrXkvJrppKjjg4bjg6njgrkNCuOAgOOAgOOAgOOAgOOAgOOAgOOAgOOAgOOAgO-8iOKYheKYheKYheKYheenu-i7ouOBl-OBvuOBl-OBn--8geKYheKYheKYheKYheKYhe-8iQ0KDQpFLW1haWwgY2hpaGFydS5wenYxLm9rYXpha2lAbWlzdW1pLmNvLmpwPG1haWx0bzpjaGloYXJ1LnB6djEub2themFraUBtaXN1bWkuY28uanA-DQoNCuS8muekvuaQuuW4ryAwNzAtMzE2NS00ODQxDQoNCioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKg0KDQoNCkZyb206IOemj-awuOOAgOacmyA8bi5mdWt1bmFnYUBtaXN1bWkuY28uanA-DQpTZW50OiBNb25kYXksIEZlYnJ1YXJ5IDIwLCAyMDIzIDI6MjkgUE0NClRvOiDkuK3ls7bjgIDlgaUgPHRha2VzaGkucnltZC5uYWthamltYUBtaXN1bWkuY28uanA-OyBPS0FaQUtJLCBDaGloYXJ1ICjlsqHltI4g5Y2D5rK7KSA8Y2hpaGFydS5wenYxLm9rYXpha2lAbWlzdW1pLmNvLmpwPg0KQ2M6IOa4oemCiuOAgOWuueWtkCA8eW8ud2F0YW5hYmVAbWlzdW1pLmNvLmpwPjsgZ2FzLWxpbmtlZEBtbC5taXN1bWkuY28uanA7IOmVt-mWgOOAgOmZveaCpiA8eS5uYWdhdG9AbWlzdW1pLmNvLmpwPg0KU3ViamVjdDog44CQ6YWN566h5Zuz5L6d6aC8X-S-nemgvOaXpTIzMDIyMF83MjExOTnmlrDnlJ_lt6Xmpa3mp5hfVFA3MuOAgA0KDQrkuK3ls7bjgZXjgpPjgIHlsqHltI7jgZXjgpMNCg0K44GK55ay44KM5qeY44Gn44GZ44CCDQrphY3nrqHlm7PkvZzmiJDkvp3poLzpoZjjgYTjgb7jgZnjgIINCg0KDQoNCjcyMTE5OSDvvIjmoKrvvInmlrDnlJ_lt6Xmpa3mp5gNCg0K5p2x44OX44Os5ZCR44GRWVk45qGI5Lu277yI5p2x44OX44Os44Kk44Oz44OJ6YeP55Sj77yJDQoNCg0K44OH44O844K_44GvR3JlZW5Gb3Jlc3TjgpLnorroqo3jgZfjgabjgY_jgaDjgZXjgYTjgIINCmh0dHBzOi8vZ2YubWlzdW1pLmNvLmpwL2NvcmFsL3BhZ2VzL2V4dGVybmFsLyMvaW5kZXguaHRtbD94YmVpZD04NDdjYzU1YjcwMzIxMDhlZWU2ZGQ4OTdmM2JjYThhNQ0KVFA3MigzNikt44Ks44K56YWN566h5Zuz5L6d6aC8LTIwMjMwMjEzLmR4Zg0K4oC75YWI6YCx44GU6YCj57Wh5beu44GX5LiK44GS44Gf5q6L44KKNOWei-WIhuOBp-OBmeOAgg0KDQoNCg0K5Zue562U5biM5pyb57SN5pyf77ya5pyA55-t5biM5pybDQoNCuS6iOWumuaXpeOCj-OBi-OBo-OBn-OCiemAo-e1oemhmOOBhOOBvuOBmeOAgg0KDQoqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioNCuagquW8j-S8muekvuODn-OCueODnw0K44CA5pel5pys5LyB5qWt5L2T44CA44Oe44O844Kx44OG44Kj44Oz44Kw44Kw44Or44O844OX44CA44Ki44Kr44Km44Oz44OI44Oe44O844Kx44OG44Kj44Oz44Kw5o6o6YCy5a6kDQrjgIDjgIDjgIDph5Hlnovjg57jg7zjgrHjg4bjgqPjg7PjgrDjg4Hjg7zjg6Ag44OX44Os44K544K744Kv44K344On44Oz44CA5bqD5bO25Za25qWt5omADQrnpo_msLjjgIDmnJsNClRFTO-8mjAxMjAtMzQzMDY2DQpGQVjvvJowNTcwLTAzNDM1NQ0KTU9CSUxF77yaMDkwLTk2ODMtNDgxNQ0KRS1tYWls77yabi5mdWt1bmFnYUBtaXN1bWkuY28uanA8bWFpbHRvOm4uZnVrdW5hZ2FAbWlzdW1pLmNvLmpwPg0KVVJM77yaaHR0cHM6Ly9qcC5taXN1bWktZWMuY29tLw0K4piF44GU55m65rOo44O744GK6KaL56mN44KK44GvV09T44KS44GU5Yip55So44GP44Gg44GV44GE77yB4piFDQoqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioNCg0K",
            },
        ),
    ],
)
def test_ExpandedMessageItem(jsonfilepath, expected):
    jsonfile = Path(jsonfilepath)

    with jsonfile.open(mode="r", encoding="utf-8") as f:
        jsondata = json.load(f)

    item = ExpandedMessageItem(jsondata)

    # ヘッダのチェック。それぞれ属性で分けている
    assert item.id == expected["id"]
    assert item.title == expected["title"]
    assert item.from_address == expected["from_address"]
    assert item.to_address == expected["to_address"]
    assert item.cc_address == expected["cc_address"]
    assert item.datetime_ == expected["datetime"]

    # ペイロードのチェック
    assert item.body == helper.decode_base64url(expected["body"]).decode("utf8")


# RenrakukoumokuInfo: 次回

# CsvFileInfo: 次回

# EstimateCalcSheetInfo: 次回、google apiを使うのでモック化が必要

# MsmAnkenMap: 次回、EstimateCalcSheetInfoを使うため

# MsmAnkenMapList: 次回、EstimateCalcSheetInfoを使うため

# 関連する関数のテスト
# generate_update_valueranges:次回。pandasでデータ構造の例を用意して生成結果が正しいか確認できればよし

#
