# coding: utf-8
import itertools
import json
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from pprint import pprint

import click
import questionary
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from helper import api_scopes, google_api_helper, load_config
from helper.mfcloud_api import MFCICledential, download_quote_pdf, generate_quote
from post_process import (
    EstimateCalcSheetInfo,
    MsmAnkenMap,
    MsmAnkenMapList,
    generate_update_valueranges,
    get_schedule_table_area,
    update_schedule_sheet,
)

# 2020-01-01 のフォーマットのみ受け付ける
START_DATE_FORMAT = "%Y-%m-%d"
GOOGLE_API_SCOPES = api_scopes.GOOGLE_API_SCOPES

# TODO:2023-03-28 これはもう使わないはずなので削除する。issue作ること
API_ENDPOINT = "https://invoice.moneyforward.com/api/v2/"

# load config, credential
config = load_config.CONFIG

MISUMI_TORIHIKISAKI_ID = config.get("mfci").get("TORIHIKISAKI_ID")
MITSUMORI_DIR_IDS = config.get("google").get("MITSUMORI_DIR_IDS")
MITSUMORI_RANGES = config["mapping"]
MOVE_DIR_ID = config.get("google").get("MOVE_DIR_ID")

GOOGLE_CREDENTIAL = config.get("google").get("CRED_FILEPATH")
table_search_range = config.get("google").get("TABLE_SEARCH_RANGE")

SCRIPT_CONFIG = config.get("generate_quotes")


@dataclass
class QuoteItem(EstimateCalcSheetInfo):
    """
    見積書の各案件事のグループ化を行うためのクラス
    """

    estimate_pdf_path: Path = field(init=False, default=None)
    mfci_quote_json: dict = field(init=False, default=None)

    def __post_init__(self):
        # 継承元のpost_initを呼ぶ
        super().__post_init__()
        # 収集した見積もり計算書の情報を元に見積書用のjsonを生成する
        self._convert_mfcloud_quote_json_data()

    def print_quote_info(self) -> None:
        """
        生成した見積情報をワンラインで表示する
        """
        print(f"見積情報: 型式:{self.anken_number} 日時:{self.duration} 価格:{self.price}")

    def _convert_mfcloud_quote_json_data(self) -> None:
        """
        見積情報を元に、MFクラウド請求書APIで使う見積書作成のjsonを生成する。結果はmfci_quote_jsonへ入れる
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

        # 結果をjsonで返す
        quote_data["quote"]["title"] = "ガススプリング配管図作製費"
        quote_data["quote"]["quote_date"] = today_datetime
        quote_data["quote"]["items"][0]["name"] = "{} ガススプリング配管図".format(
            self.anken_number
        )
        quote_data["quote"]["items"][0]["quantity"] = 1
        quote_data["quote"]["items"][0]["detail"] = f"納期 {self.duration:%m/%d}"
        quote_data["quote"]["items"][0]["unit_price"] = int(self.price)

        # LRは条件判断を行う
        rh_flag = self.anken_number.split("-")[-1]
        if rh_flag in ("RH", "LH") != 0:
            # RHの場合はLH, LHの場合はRHの備考文章を作成
            reverse_part_number = "MA-" + "-".join(self.anken_number.split("-")[0:-1])
            if rh_flag == "RH":
                reverse_part_number = reverse_part_number + "-LH"
            else:
                reverse_part_number = reverse_part_number + "-RH"

            quote_data["quote"]["note"] = "本見積は{}の対象側作図案件となります".format(
                reverse_part_number
            )
        self.mfci_quote_json = quote_data


@click.command()
@click.option("--dry-run", is_flag=True, help="Dry Run Flag")
def main(dry_run):
    # 一連の操作中に使うデータ構造を入れるリスト（グループ化はメール生成時に行う）
    quote_items: list[QuoteItem] = []

    # Googleのtokenを用意
    google_cred = google_api_helper.get_cledential(GOOGLE_API_SCOPES)
    gdrive_serivice = build("drive", "v3", credentials=google_cred)
    gmail_service = build("gmail", "v1", credentials=google_cred)

    # mfcloudのセッション作成
    mfci_cred = MFCICledential()
    mfcloud_invoice_session = mfci_cred.get_session()

    # google sheetのリストを取得
    googledrive_search_target_mimetype = "application/vnd.google-apps.spreadsheet"
    # - 特定のフォルダ（GSheet的にはグループ）の一覧を取得

    # テンプレフォルダ内もフィルターにいれる。その時にテンプレートは除外する
    dir_ids_query = " or ".join((f"'{id}' in parents " for id in MITSUMORI_DIR_IDS))
    # TODO: 2023-04-20ここの検索もgoogleapiのhelperに移動したい
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

    if not target_items:
        print("見積もり計算表が見つかりませんでした。終了します。")
        sys.exit(0)

    # TODO:2023-04-19 ここのnameはスプレッドシートの名称ではなくて案件番号にする方がいい。

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

    # MFクラウドの見積書jsonデータを作成させて関連結果含めてQuoteItemに入れる
    for estimate_calcsheet in selected_estimate_calcsheets:
        quote_item = QuoteItem(estimate_calcsheet.get("id"))
        quote_item.calcsheet_parents = estimate_calcsheet.get("parents")
        # サマリーを表示する
        quote_item.print_quote_info()
        # itemをリストアップ
        quote_items.append(quote_item)

    # dry-runはここまで。dry-runは結果をjsonで返す。
    if dry_run:
        print("[dry run]json dump:")
        pprint(quote_items)
        sys.exit(0)

    # MFクラウドで見積書作成
    for quote_item in quote_items:
        # 見積書作成
        generated_quote_result = generate_quote(
            mfcloud_invoice_session, quote_item.mfci_quote_json
        )

        # errorなら終了する
        if "errors" in generated_quote_result:
            print("エラーが発生しました。詳細はレスポンスを確認ください")
            pprint(generated_quote_result)
            sys.exit(0)

        # PDFのファイル名はミスミの型式をつける
        quote_item.estimate_pdf_path = (
            Path("./quote") / f"見積書_{quote_item.anken_number}.pdf"
        )
        download_quote_pdf(
            mfcloud_invoice_session,
            generated_quote_result["data"]["attributes"]["pdf_url"],
            quote_item.estimate_pdf_path,
        )

        # - 生成後、今回選択したスプレッドシートは生成済みフォルダへ移動する
        # TODO: 2023-04-20 ここは関数として切り出す -> helper.googleapi
        try:
            previous_parents = ",".join(quote_item.calcsheet_parents)

            _ = (
                gdrive_serivice.files()
                .update(
                    fileId=quote_item.calcsheet_source,
                    addParents=MOVE_DIR_ID,
                    removeParents=previous_parents,
                    fields="id, parents",
                )
                .execute()
            )

        except HttpError as error:
            sys.exit(f"スプレッドシート移動時にエラーが発生しました: {error}")

        # スケジュール表の該当行に価格や納期を追加する
        msmanken_info = MsmAnkenMap(estimate_calcsheet_info=quote_item)
        msmankenmaplist = MsmAnkenMapList()
        msmankenmaplist.msmankenmap_list.append(msmanken_info)

        export_pd = msmankenmaplist.generate_update_sheet_values()
        before_pd = get_schedule_table_area(table_search_range, google_cred)
        update_data = generate_update_valueranges(
            table_search_range, before_pd, export_pd
        )
        print(f"update result:{update_data}")

        update_schedule_sheet(update_data, google_cred)

    # メールの下書きを生成。案件のベース番号をもとにグルーピングをして一つのメールに複数の見積を添付する
    quote_groups = itertools.groupby(quote_items, lambda x: x.anken_base_number)

    for group_key, quote_iter in quote_groups:
        # quote_itemsをlistに変換する
        quote_items = list(quote_iter)

        # メール生成のテンプレは別のファイルに書く。
        # 納期はグループ内最初のQuoteItemのものを利用（案件に対して同じ納期を設定している前提）
        mail_template_body: str = SCRIPT_CONFIG.get("mail_template_body")
        replybody = mail_template_body.replace("{{nouki}}", quote_items[0].duration_str)

        # メールのスレッドを取得して、スレッドに返信する
        searchquery = f"label:snd-ミスミ (*{group_key}*)"

        threads = google_api_helper.search_threads(gmail_service, searchquery)
        if not threads:
            sys.exit("スレッドが見つかりませんでした。メール返信作成を中止します。")

        # TODO:2023-04-18 ここは複数スレッドがあった場合は選択制にする。出ない場合は一番上のものを使いますと、タイトルを出して確認させる。
        # メッセージが大抵一つだが、一番上を取り出す（一番上が最新のはず）
        message = google_api_helper.get_messages_by_threadid(
            gmail_service, threads[0].get("id", "")
        )[0]

        # 返信メッセージで下書きを生成
        print(
            google_api_helper.append_draft_in_thread(
                gmail_service,
                replybody,
                (quote_item.estimate_pdf_path for quote_item in quote_items),
                message["id"],
                threads[0].get("id", ""),
            )
        )


if __name__ == "__main__":
    main()
