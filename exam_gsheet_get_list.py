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


def get_quote_gsheet_by_quote_list_gsheet(
    gsheet_service, per_page_length: int
) -> list | None:
    # Googleスプレッドシートの見積書一覧をみて、最後尾から必要な件数の見積書スプレッドシートのURLリストを取得する

    # 見積計算表の生成元シートの最大行範囲を取得
    range_name = "見積書管理!D2:D"
    result = (
        gsheet_service.spreadsheets()
        .values()
        .get(spreadsheetId=QUOTE_FILE_LIST_GSHEET_ID, range=range_name)
        .execute()
    )
    values = result.get("values", [])

    if not values:
        return None
    else:
        # 最後の行を取得
        last_row = len(values)
        print(f"{range_name} 最後の行: {last_row}")
        return values[last_row - per_page_length : last_row]


if __name__ == "__main__":
    print(get_quote_gsheet_by_quote_list_gsheet(gsheet_service, 100))
