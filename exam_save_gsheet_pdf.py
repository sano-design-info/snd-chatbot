from pathlib import Path
from googleapiclient.discovery import build
# from googleapiclient.errors import HttpError

from api import googleapi
from helper import load_config

config = load_config.CONFIG

SCRIPT_CONFIG = config.get("generate_quotes")

# GoogleスプレッドシートのテンプレートID
target_sheet_id = "1bIOb-ODmbQ4Zj8sXUDH59uwWYDafAFg0P9mJwNcaLns"

# Googleのtokenを用意
google_cred = googleapi.get_cledential(googleapi.API_SCOPES)
gdrive_service = build("drive", "v3", credentials=google_cred)


def save_pdf() -> str:
    # パラメーターセットをして、PDFを生成する

    # 見積書のPDFをダウンロード
    googleapi.export_pdf_by_driveexporturl(
        google_cred.token,
        target_sheet_id,
        Path("./quote_template.pdf"),
        {
            # "gid": "0",
            # "size": "7",
            "portrait": "false",
            # "fitw": "true",
            # "gridlines": "false",
        },
    )


if __name__ == "__main__":
    print(target_sheet_id)
    save_pdf()
    print(f"saved!: { Path('./quote_template.pdf')}")
