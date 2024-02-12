from googleapiclient.discovery import build
# from googleapiclient.errors import HttpError

from api import googleapi
from helper import load_config

from task import get_billing

config = load_config.CONFIG

SCRIPT_CONFIG = config.get("generate_quotes")
# 見積書のファイル一覧を記録するGoogleスプレッドシートのID

# Googleのtokenを用意
GOOGLE_CREDENTIAL = config.get("google").get("CRED_FILEPATH")
google_cred = googleapi.get_cledential(googleapi.API_SCOPES)
gsheet_service = build("sheets", "v4", credentials=google_cred)

if __name__ == "__main__":
    print("run get_quote_gsheet_by_quote_list_gsheet")
    quote_gsheet_id_list_under_100 = [
        i[0].split("/")[-1]
        for i in get_billing.get_quote_gsheet_by_quote_list_gsheet(gsheet_service, 100)
    ]
    print(quote_gsheet_id_list_under_100)

    hinmoku_celladdrs = get_billing.get_hinmoku_celladdrs_by_gsheet()
    print(f"hinmoku_celladdrs: {hinmoku_celladdrs.values()}")

    quote_data_by_gsheet = [
        get_billing.get_values_by_range(
            gsheet_service, id, get_billing.get_hinmoku_celladdrs_by_gsheet()
        )
        for id in quote_gsheet_id_list_under_100
    ]
    print(quote_data_by_gsheet)
