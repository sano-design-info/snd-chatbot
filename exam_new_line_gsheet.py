from googleapiclient.discovery import build
# from googleapiclient.errors import HttpError

from api import googleapi
from helper import load_config

config = load_config.CONFIG

SCRIPT_CONFIG = config.get("generate_quotes")
# 見積書のファイル一覧を記録するGoogleスプレッドシートのID
QUOTE_FILE_LIST_GSHEET_ID = SCRIPT_CONFIG.get("QUOTE_FILE_LIST_GSHEET_ID")

# Googleのtokenを用意
GOOGLE_CREDENTIAL = config.get("google").get("CRED_FILEPATH")
google_cred = googleapi.get_cledential(googleapi.API_SCOPES)
gsheet_service = build("sheets", "v4", credentials=google_cred)


def get_quote_id() -> str:
    # 見積管理シートへ行を入れて見積番号を取得する

    new_quote_line = googleapi.append_sheet(
        gsheet_service,
        QUOTE_FILE_LIST_GSHEET_ID,
        "見積書管理",
        [['=TEXT(ROW()-1,"0000")', "", "", ""]],
        "USER_ENTERED",
        "INSERT_ROWS",
        True,
    )
    print(new_quote_line)
    # 追加できた行のA列の値を取得
    return new_quote_line.get("updates").get("updatedData").get("values")[0][0]


if __name__ == "__main__":
    print(get_quote_id())
