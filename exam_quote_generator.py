from datetime import datetime
import json

# from pathlib import Path
import sys
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from api import googleapi
from api.googleapi import sheet_data_mapper
from helper import load_config, EXPORTDIR_PATH


START_DATE_FORMAT = "%Y-%m-%d"

config = load_config.CONFIG

SCRIPT_CONFIG = config.get("generate_quotes")

# GoogleスプレッドシートのテンプレートID
target_sheet_id = "1bIOb-ODmbQ4Zj8sXUDH59uwWYDafAFg0P9mJwNcaLns"

# load config, credential
config = load_config.CONFIG

TORIHIKISAKI_NAME = config.get("general").get("TORIHIKISAKI_NAME")

SCRIPT_CONFIG = config.get("generate_quotes")

# 見積書のGoogleスプレッドシート関連
# 見積書のファイル一覧を記録するGoogleスプレッドシートのID
QUOTE_FILE_LIST_GSHEET_ID = SCRIPT_CONFIG.get("QUOTE_FILE_LIST_GSHEET_ID")
# GoogleスプレッドシートのテンプレートID
QUOTE_TEMPLATE_GSHEET_ID = SCRIPT_CONFIG.get("QUOTE_TEMPLATE_GSHEET_ID")
# テンプレートに入力するセルマッピング・JSONファイルのパス
QUOTE_TEMPLATE_CELL_MAPPING_JSON_PATH = SCRIPT_CONFIG.get(
    "QUOTE_TEMPLATE_CELL_MAPPING_JSON_PATH"
)
# 見積書のGoogleスプレッドシート保存先
QUOTE_GSHEET_SAVE_DIR_IDS = SCRIPT_CONFIG.get("QUOTE_GSHEET_SAVE_DIR_IDS")
# 見積書のPDF保存先
QUOTE_PDF_SAVE_DIR_IDS = SCRIPT_CONFIG.get("QUOTE_PDF_SAVE_DIR_IDS")

# 定数から全体に使う変数を作成
export_quote_dirpath = EXPORTDIR_PATH / "quote"
export_quote_dirpath.mkdir(parents=True, exist_ok=True)
with open(QUOTE_TEMPLATE_CELL_MAPPING_JSON_PATH, "r", encoding="utf-8") as f:
    quote_template_cell_mapping_dict = json.load(f)

# Googleのtokenを用意
google_cred = googleapi.get_cledential(googleapi.API_SCOPES)
gdrive_service = build("drive", "v3", credentials=google_cred)
gsheet_service = build("sheets", "v4", credentials=google_cred)


ANKEN_NUMBER = "MA-9991"
ANKEN_DURATION = datetime(2024, 2, 10)


def convert_dict_to_gsheet_tamplate(quote_number) -> None:
    # 見積書へ書き込むデータを作る
    # ここで作成するデータは、見積書のテンプレートに合わせたデータを作成する

    hinmoku = {
        "name": f"{ANKEN_NUMBER} ガススプリング配管図",
        "detail": f"納期 {ANKEN_DURATION:%m/%d}",
        "price": 100000,
        "quantity": 1,
        "zeiritu": "10%",
    }

    # LRは条件判断を行う
    quote_note = ""
    rh_flag = ANKEN_NUMBER.split("-")[-1]
    if rh_flag in ("RH", "LH") != 0:
        # RHの場合はLH, LHの場合はRHの備考文章を作成
        reverse_part_number = "MA-" + "-".join(ANKEN_NUMBER.split("-")[0:-1])
        if rh_flag == "RH":
            reverse_part_number = reverse_part_number + "-LH"
        else:
            reverse_part_number = reverse_part_number + "-RH"

        quote_note = f"本見積は{reverse_part_number}の対象側作図案件となります"

    # TODO:2024-02-05 これは設定にする
    quote_id_prefix = "DCCF6E"
    # 見積書番号を生成
    # TODO: 2024-02-05 ここのget_quote_idはMainTask側で処理して情報をanken_quoteに入れる
    qupte_id = f"{quote_id_prefix}-Q-{quote_number}"

    today_datetime = datetime.now()
    return {
        "customer_name": TORIHIKISAKI_NAME,
        "quote_id": qupte_id,
        "title": "ガススプリング配管図作製費",
        # 日付は実行時の日付を利用
        "quote_date": today_datetime.strftime(START_DATE_FORMAT),
        "note": quote_note,
        "item_table": [hinmoku],
    }


def get_quote_number_by_quote_manage_gsheet(updated_result_quote_manage_gsheet) -> str:
    # 追加できた行のA列の値を取得
    return (
        updated_result_quote_manage_gsheet.get("updates")
        .get("updatedData")
        .get("values")[0][0]
    )


def get_updated_cell_address_by_quote_manage_gsheet(
    updated_result_quote_manage_gsheet
) -> str:
    # 見積管理シートへ行を入れて見積番号を取得する
    return updated_result_quote_manage_gsheet.get("updates").get("updatedRange")


def generate_and_manage_quote() -> str:
    try:
        # [見積書作成を行う]
        # 見積書管理表から番号を生成
        print("見積書の作成を行います。")
        print("見積書の管理表から番号を生成します。")
        updated_quote_manage_gsheet = googleapi.append_sheet(
            gsheet_service,
            QUOTE_FILE_LIST_GSHEET_ID,
            "見積書管理",
            [['=TEXT(ROW()-1,"0000")', "", "", ""]],
            "USER_ENTERED",
            "INSERT_ROWS",
            True,
        )
        # 見積書番号を取得
        quote_id = get_quote_number_by_quote_manage_gsheet(updated_quote_manage_gsheet)

        # 見積書の情報を生成

        # googleスプレッドシートの見積書テンプレートを複製する
        print("見積書のテンプレートを複製します。")
        quote_file_id = googleapi.dupulicate_file(
            gdrive_service,
            QUOTE_TEMPLATE_GSHEET_ID,
            quote_filestem := f"見積書_{ANKEN_NUMBER}",
        )
        print(f"見積書のテンプレートを複製しました。: {quote_file_id}")

        # 見積書のファイル名と保存先を設定
        print("見積書のファイル名と保存先を設定します。")
        _ = googleapi.update_file(
            gdrive_service,
            file_id=quote_file_id,
            body=None,
            add_parents=QUOTE_GSHEET_SAVE_DIR_IDS,
            fields="id, parents",
        )

        # 見積書へanken_quoteの内容を記録
        print("見積書へ見積内容を記録します。")
        sheet_data_mapper.write_data_to_sheet(
            gsheet_service,
            quote_file_id,
            convert_dict_to_gsheet_tamplate(quote_id),
            quote_template_cell_mapping_dict,
        )

        # ファイル名:見積書_[納期].pdf
        quote_filename = f"{quote_filestem}.pdf"

        # 見積書のPDFをダウンロード
        print("見積書のPDFをダウンロードします。")
        googleapi.export_pdf_by_driveexporturl(
            google_cred.token,
            quote_file_id,
            export_quote_dirpath / quote_filename,
            {
                "gid": "0",
                "size": "7",
                "portrait": "true",
                "fitw": "true",
                "gridlines": "false",
            },
        )

        # 見積書のPDFをGoogleドライブへ保存
        print("見積書のPDFをGoogleドライブへ保存します。")
        upload_pdf_result = googleapi.upload_file(
            gdrive_service,
            export_quote_dirpath / quote_filename,
            "application/pdf",
            "application/pdf",
            QUOTE_PDF_SAVE_DIR_IDS,
        )

        # 見積書のGoogleスプレッドシートとPDFのURLを見積管理表に記録

        # 見積管理表を更新する。B列から[ファイル名, 見積書:Gsheet のIDからURL, 見積書:GDrive PDFのIDからURL]
        # 見積管理表に番号を追加したupdatedRows（updated_quote_manage_cell_address）を使うがB列以降を使う
        # 生成した見積番号のセルアドレスからB列に置き換えて取得。AのみをBにする
        # 例: updated_quote_manage_cell_address = "見積書管理!A2:D2" -> "見積書管理!B2:D2"
        print("見積管理表を更新します。")
        _ = googleapi.update_sheet(
            gsheet_service,
            QUOTE_FILE_LIST_GSHEET_ID,
            get_updated_cell_address_by_quote_manage_gsheet(
                updated_quote_manage_gsheet
            ).replace("A", "B"),
            [
                [
                    quote_filename,
                    f"http://docs.google.com/spreadsheets/d/{quote_file_id}",
                    f"http://drive.google.com/file/d/{upload_pdf_result.get('id')}",
                ]
            ],
        )
        print("見積書のPDFをダウンロードしました。")
        # - 生成後、今回選択したスプレッドシートは生成済みフォルダへ移動する
        # _ = googleapi.update_file(
        #     gdrive_service,
        #     file_id=anken_quote.calcsheet_source,
        #     add_parents=ARCHIVED_ESTIMATECALCSHEET_DIR_ID,
        #     remove_parents=anken_quote.calcsheet_parents,
        #     fields="id",
        # )

    except HttpError as error:
        sys.exit(f"何かのエラーでした。: {error}")


if __name__ == "__main__":
    generate_and_manage_quote()
