# coding: utf-8
import json
import os
import os.path
import sys
from datetime import datetime
from pathlib import Path
from pprint import pprint

import re

import click
import questionary

import dotenv
import toml
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from helper.mfcloud_api import MFCICledential, download_quote_pdf, generate_quote
from helper import google_api_helper, api_scopes
from post_process import (
    EstimateCalcSheetInfo,
    MsmAnkenMap,
    MsmAnkenMapList,
    generate_update_valueranges,
    get_schedule_table_area,
    update_schedule_sheet,
)

# load config, credential
dotenv.load_dotenv()
GOOGLE_CREDENTIAL = os.environ.get("CRED_FILEPATH")

# 2020-01-01 のフォーマットのみ受け付ける
START_DATE_FORMAT = "%Y-%m-%d"

CONFIG_FILE = Path("mfi_estimate_generator.toml")
config = toml.load(CONFIG_FILE)

MISUMI_TORIHIKISAKI_ID = config.get("mfci").get("TORIHIKISAKI_ID")
MITSUMORI_DIR_IDS = config.get("googledrive").get("MITSUMORI_DIR_IDS")

MISTUMORI_NUMBER_PATTERN = re.compile("^.*_MA-(?P<number>.*)")

GOOGLE_API_SCOPES = api_scopes.GOOGLE_API_SCOPES

# TODO:2023-03-28 これはもう使わないはずなので削除する。issue作ること
API_ENDPOINT = "https://invoice.moneyforward.com/api/v2/"

MITSUMORI_RANGES = config["mapping"]
MOVE_DIR_ID = config.get("googledrive").get("MOVE_DIR_ID")


table_search_range = os.environ["TABLE_SEARCH_RANGE"]


def print_quote_info(**kargs):
    """
    生成した見積情報をワンラインで表示する
    """

    print(
        f'見積情報: 型式:MA-{kargs["part_number"]} 日時:{kargs["duration"]} 価格:{kargs["price"]}'
    )


def fix_datetime(datetime_str: str) -> datetime:
    """
    日付の入力の区切り文字を修正する。
    """
    for splitecahr in (".", "/"):
        if splitecahr in datetime_str:
            return datetime.strptime(datetime_str, splitecahr.join(("%Y", "%m", "%d")))


def seikika_data(estimate_data: dict) -> dict:

    # 入る値は全てイミュータブルなのでcopyでヨシにした
    result_data = estimate_data.copy()

    # 日付文字列をdatetime化する、日付が未定の場合（入ってないや文字列の場合）はそのままスルーする
    if isinstance(fix_datetime(estimate_data["duration"]), datetime):
        duration_str = f"納期 {fix_datetime(estimate_data['duration']): %m/%d}"
    else:
        duration_str = f"納期 {estimate_data['duration']}"
    result_data["duration"] = duration_str

    result_data["price"] = int(re.sub(r"[\¥\,]", "", estimate_data["price"]))
    result_data["discount_flag"] = int(re.sub(r"[\¥\,]", "", estimate_data["price"]))

    return result_data


def generate_quote_json_data(**kargs) -> dict:
    """
    見積情報を元に、MFクラウド請求書APIで使う見積書作成のjsonを生成する
    """

    today_datetime = datetime.now().strftime(START_DATE_FORMAT)
    quote_json_template = """
    {
        "quote": {
            "department_id": "",
            "quote_date": "2020-05-08",
            "title": "string_Testtitle",
            "note": "",
            "memo": "",
            "tags": "佐野設計自動生成",
            "items": [
                {
                    "name": "品目",
                    "detail": "詳細",
                    "unit_price": 0,
                    "unit": "",
                    "quantity": 0,
                    "excise": true
                }
            ]
        }
    }
    """

    # jsonでロードする
    quote_data = json.loads(quote_json_template)

    # department_idはミスミのものを利用
    quote_data["quote"]["department_id"] = MISUMI_TORIHIKISAKI_ID

    # 各情報を入れる

    # 結果をjsonで返す

    quote_data["quote"]["title"] = "ガススプリング配管図作製費"

    quote_data["quote"]["quote_date"] = today_datetime
    quote_data["quote"]["items"][0]["name"] = "MA-{} ガススプリング配管図".format(
        kargs["part_number"]
    )
    quote_data["quote"]["items"][0]["quantity"] = 1
    quote_data["quote"]["items"][0]["detail"] = kargs["duration"]
    quote_data["quote"]["items"][0]["unit_price"] = kargs["price"]

    # LRは条件判断を行う
    rh_flag = kargs["part_number"].split("-")[-1]

    # TODO:2022-11-14 discount_flagはもう利用していないので、2022-11-14現在で最新の計算表v2で問題がなくなったらand条件を外す
    if rh_flag in ("RH", "LH") and int(kargs["discount_flag"]) != 0:

        # RHの場合はLH, LHの場合はRHの備考文章を作成
        reverse_part_number = "MA-" + "-".join(kargs["part_number"].split("-")[0:-1])

        if rh_flag == "RH":
            reverse_part_number = reverse_part_number + "-LH"
        else:
            reverse_part_number = reverse_part_number + "-RH"

        quote_data["quote"]["note"] = "本見積は{}の対象側作図案件となります".format(reverse_part_number)
    return quote_data


@click.command()
@click.option("--dry-run", is_flag=True, help="Dry Run Flag")
def main(dry_run):

    # Googleのtokenを用意
    google_cred = google_api_helper.get_cledential(GOOGLE_API_SCOPES)
    gdrive_serivice = build("drive", "v3", credentials=google_cred)
    gsheet_service = build("sheets", "v4", credentials=google_cred)

    # mfcloudのセッション作成
    mfci_cred = MFCICledential()
    mfcloud_invoice_session = mfci_cred.get_session()

    # google sheetのリストを取得
    googledrive_search_target_mimetype = "application/vnd.google-apps.spreadsheet"
    # - 特定のフォルダ（GSheet的にはグループ）の一覧を取得

    # テンプレフォルダ内もフィルターにいれる。その時にテンプレートは除外する
    dir_ids_query = " or ".join((f"'{id}' in parents " for id in MITSUMORI_DIR_IDS))
    try:
        contain_mitumori_folder_query = f"""
        ({dir_ids_query})
        and mimeType = '{googledrive_search_target_mimetype}'
        and trashed = false
        and name != "ミスミ配管図見積り計算表v2_MA-[ミスミ型番]"
        """

        mitumori_sheet_list = (
            gdrive_serivice.files()
            .list(
                q=contain_mitumori_folder_query,
                pageSize=10,
                fields="nextPageToken, files(id, name, parents)",
            )
            .execute()
        )
        # 取得結果を反転させる。順番を作成順にしたほうがわかりやすい
        target_items = list(reversed(mitumori_sheet_list.get("files", [])))

    except HttpError as error:
        sys.exit(f"Google Drive APIのエラーが発生しました。: {error}")

    # 一覧から該当する見積もり計算表を取得
    selected_estimate_calcsheets = questionary.checkbox(
        "見積もりを作成する見積もり計算表を選択してください。",
        choices=[
            {
                "name": f"{index:0>2}: {mitsumori_gsheet.get('name')}",
                "value": mitsumori_gsheet,
            }
            for index, mitsumori_gsheet in enumerate(target_items, 1)
        ],
    ).ask()

    # キャンセル処理を入れる
    if not selected_estimate_calcsheets:
        print("操作をキャンセルしました。終了します。")
        sys.exit(0)

    print(selected_estimate_calcsheets)

    # TODO:2023-03-29 選択結果を保持して複数の見積書作成を行う
    # 見積のデータを生成
    quote_datas_set = []
    for estimate_calcsheet in selected_estimate_calcsheets:
        estimate_data = {}

        # シートのファイル名から品番生成
        estimate_data["part_number"] = MISTUMORI_NUMBER_PATTERN.match(
            estimate_calcsheet.get("name")
        ).group("number")

        # GSheetから値取得
        try:
            # Call the Sheets API
            sheet = gsheet_service.spreadsheets()
            gsheet_result = (
                sheet.values()
                .batchGet(
                    spreadsheetId=estimate_calcsheet.get("id"),
                    ranges=list(MITSUMORI_RANGES.values()),
                )
                .execute()
            )
            gsheet_values = gsheet_result.get("valueRanges", [])

            if not gsheet_values:
                print("No data found.")
                return

            for key, value in zip(MITSUMORI_RANGES.keys(), gsheet_values):
                estimate_data[key] = value.get("values")[0][0]

        except HttpError as error:
            sys.exit(f"Google Sheet APIでエラーが発生しました。: {error}")

        # 型変換する
        estimate_data = seikika_data(estimate_data)

        # サマリーを表示する
        print_quote_info(**estimate_data)

        # 収集した見積情報を元にMFクラウドへ渡すjsonデータを生成
        quote_data = generate_quote_json_data(**estimate_data)

        quote_datas_set.append((quote_data, estimate_data, estimate_calcsheet))

    # dry-runはここまで。dry-runは結果をjsonで返す。
    if dry_run:
        print("[dry run]json dump:")
        pprint(quote_datas_set)
        sys.exit(0)

    for quote_data, estimate_data, estimate_calcsheet in quote_datas_set:
        generated_quote_result = generate_quote(mfcloud_invoice_session, quote_data)

        # errorなら終了する
        if "errors" in generated_quote_result:
            print("エラーが発生しました。詳細はレスポンスを確認ください")
            pprint(generated_quote_result)
            sys.exit(0)

        # PDFのファイル名はミスミの型式で行う
        filename_partname = estimate_data["part_number"]
        save_filepath = Path("./quote") / f"見積書_MA-{filename_partname}.pdf"

        download_quote_pdf(
            mfcloud_invoice_session,
            generated_quote_result["data"]["attributes"]["pdf_url"],
            save_filepath,
        )

        # - 生成後、今回選択したスプレッドシートは生成済みフォルダへ移動する
        try:
            previous_parents = ",".join(estimate_calcsheet.get("parents"))

            _ = (
                gdrive_serivice.files()
                .update(
                    fileId=estimate_calcsheet["id"],
                    addParents=MOVE_DIR_ID,
                    removeParents=previous_parents,
                    fields="id, parents",
                )
                .execute()
            )

        except HttpError as error:
            sys.exit(f"スプレッドシート移動時にエラーが発生しました: {error}")

        # スケジュール表の該当行に価格を追加する
        target_sheet_id = estimate_calcsheet.get("id")

        # 扱う行は一つなので登録も一つのみ
        estimate_calcsheet_info = EstimateCalcSheetInfo(target_sheet_id)
        msmanken_info = MsmAnkenMap(estimate_calcsheet_info=estimate_calcsheet_info)
        msmankenmaplist = MsmAnkenMapList()
        msmankenmaplist.msmankenmap_list.append(msmanken_info)

        export_pd = msmankenmaplist.generate_update_sheet_values()
        before_pd = get_schedule_table_area(table_search_range, google_cred)

        update_data = generate_update_valueranges(
            table_search_range, before_pd, export_pd
        )
        print(f"update result:{update_data}")

        update_schedule_sheet(update_data, google_cred)


if __name__ == "__main__":
    main()
