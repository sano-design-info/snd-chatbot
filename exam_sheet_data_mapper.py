from googleapiclient.discovery import build
from api import googleapi
from api.googleapi.sheet_data_mapper import write_data_to_sheet

# 認証情報の取得
google_cred = googleapi.get_cledential(googleapi.API_SCOPES)
gsheet_service = build("sheets", "v4", credentials=google_cred)

# https://docs.google.com/spreadsheets/d/1ptsDjomKNIK-YdfH-gxLoYzlYFfTF179jFh_X4ykzXo/edit#gid=0
# GoogleスプレッドシートのIDとシート名
SPREADSHEET_ID = "1ptsDjomKNIK-YdfH-gxLoYzlYFfTF179jFh_X4ykzXo"
SHEET_NAME = "sheet1"

# 推定データ
estimate_data = {
    "estimate_id": "TEST-0999",
    "title": "テスト図作製費",
    "memo": "",
    "quote_date": "2023-12-09",
    "expired_date": "2023-12-10",
    "item_table": [
        {
            "name": "MA-9999 TEST制作図",
            "detail": "納期: 12/31",
            "price": 0,
            "quantity": 1,
            "zeiritu": "10%",
        },
    ],
}

# セルマッピング情報
cell_mapping = {
    "singlecell": {
        "estimate_id": "sheet1!A2",
        "title": "sheet1!A3",
        "memo": "sheet1!A4",
        "quote_date": "sheet1!A5",
    },
    "tables": {
        "item_table": {
            "columns": {
                "name": "sheet1!B",
                "detail": "sheet1!C",
                "price": "sheet1!D",
                "quantity": "sheet1!E",
            },
            "startRow": 13,
            "endRow": 23,
        }
    },
}

# データをシートに書き込む
write_data_to_sheet(gsheet_service, SPREADSHEET_ID, estimate_data, cell_mapping)
