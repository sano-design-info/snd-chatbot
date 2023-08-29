# coding: utf-8
import itertools
import json
import sys
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from pprint import pprint

import click
import questionary
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import api.googleapi
from api.mfcloud_api import (
    MFCIClient,
    download_quote_pdf,
    create_quote,
    create_item,
    attach_item_into_quote,
)
from helper import load_config, EXPORTDIR_PATH
from itemparser import (
    EstimateCalcSheetInfo,
    MsmAnkenMap,
    MsmAnkenMapList,
    generate_update_valueranges,
    get_schedule_table_area,
    update_schedule_sheet,
)

# load config, credential
config = load_config.CONFIG

MISUMI_TORIHIKISAKI_ID = config.get("mfci").get("TORIHIKISAKI_ID")
MITSUMORI_DIR_IDS = config.get("google").get("MITSUMORI_DIR_IDS")
MITSUMORI_RANGES = config["mapping"]
MOVE_DIR_ID = config.get("google").get("MOVE_DIR_ID")

GOOGLE_CREDENTIAL = config.get("google").get("CRED_FILEPATH")
table_search_range = config.get("google").get("TABLE_SEARCH_RANGE")

SCRIPT_CONFIG = config.get("generate_quotes")

export_qupte_dirpath = EXPORTDIR_PATH / "quote"
export_qupte_dirpath.mkdir(parents=True, exist_ok=True)

# 2020-01-01 のフォーマットのみ受け付ける
START_DATE_FORMAT = "%Y-%m-%d"


@dataclass
class AnkenQuote(EstimateCalcSheetInfo):
    """
    見積書の各案件事のグループ化を行うためのクラス
    """

    estimate_pdf_path: Path = field(init=False, default=None)
    mfci_quote_json: dict = field(init=False, default=None)
    mfci_quote_item_json: dict = field(init=False, default=None)

    def __post_init__(self):
        # 継承元のpost_initを呼ぶ
        super().__post_init__()
        # 収集した見積もり計算書の情報を元に見積書用のjsonを生成する
        self._convert_json_mfci_quote_item()
        self._convert_json_mfci_quote()

    def print_quote_info(self) -> None:
        """
        生成した見積情報をワンラインで表示する
        """
        print(f"見積情報: 型式:{self.anken_number} 日時:{self.duration} 価格:{self.price}")

    def _convert_json_mfci_quote_item(self) -> None:
        """
        見積情報を元に、MFクラウド請求書APIで使う見積書向け品目用のjson文字列を生成する。
        結果はmfci_quote_item_jsonへ入れる
        """
        item_json_template = """
        {
            "name": "品目",
            "detail": "詳細",
            "unit": "0",
            "price": 0,
            "quantity": 1,
            "excise": "ten_percent"
        }
        """

        # jsonでロードする
        quote_item = json.loads(item_json_template)

        # 結果をjsonで返す
        quote_item["name"] = "{} ガススプリング配管図".format(self.anken_number)
        quote_item["quantity"] = 1
        quote_item["detail"] = f"納期 {self.duration:%m/%d}"
        quote_item["price"] = int(self.price)

        self.mfci_quote_item_json = quote_item

    def _convert_json_mfci_quote(self) -> None:
        """
        見積情報を元に、MFクラウド請求書APIで使う見積書作成のjsonを生成する。
        結果はmfci_quote_jsonへ入れる
        """

        today_datetime = datetime.now()
        quote_json_template = """
        {
            "department_id": "",
            "title": "ガススプリング配管図作製費",
            "memo": "",
            "quote_date": "2022-12-09",
            "expired_date": "2022-12-10",
            "note": "",
            "tag_names": [
                "佐野設計自動生成"
            ]
        }
        """

        # jsonでロードする
        quote_data = json.loads(quote_json_template)

        # department_id
        quote_data["department_id"] = MISUMI_TORIHIKISAKI_ID
        # 日付は実行時の日付を利用
        quote_data["quote_date"] = today_datetime.strftime(START_DATE_FORMAT)
        # 有効期限は１週間後
        quote_data["expired_date"] = (today_datetime + timedelta(days=7)).strftime(
            START_DATE_FORMAT
        )

        # LRは条件判断を行う
        rh_flag = self.anken_number.split("-")[-1]
        if rh_flag in ("RH", "LH") != 0:
            # RHの場合はLH, LHの場合はRHの備考文章を作成
            reverse_part_number = "MA-" + "-".join(self.anken_number.split("-")[0:-1])
            if rh_flag == "RH":
                reverse_part_number = reverse_part_number + "-LH"
            else:
                reverse_part_number = reverse_part_number + "-RH"

            quote_data["note"] = f"本見積は{reverse_part_number}の対象側作図案件となります"

        self.mfci_quote_json = quote_data


@click.command()
@click.option("--dry-run", is_flag=True, help="Dry Run Flag")
def main(dry_run):
    # Googleのtokenを用意
    google_cred = api.googleapi.get_cledential(api.googleapi.API_SCOPES)
    gdrive_serivice = build("drive", "v3", credentials=google_cred)
    gmail_service = build("gmail", "v1", credentials=google_cred)
    sheet_service = build("sheets", "v4", credentials=google_cred)
    # 一連の操作中に使うデータ構造を入れるリスト（グループ化はメール生成時に行う）
    anken_quotes: list[AnkenQuote] = []

    # mfcloudのセッション作成
    mfci_cred = MFCIClient()
    mfci_session = mfci_cred.get_session()

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
        anken_quote = AnkenQuote(sheet_service, estimate_calcsheet.get("id"))
        anken_quote.calcsheet_parents = estimate_calcsheet.get("parents")
        # サマリーを表示する
        anken_quote.print_quote_info()
        # itemをリストアップ
        anken_quotes.append(anken_quote)

    # dry-runはここまで。dry-runは結果をjsonで返す。
    if dry_run:
        print("[dry run]json dump:")
        pprint(anken_quotes)
        sys.exit(0)

    # MFクラウドで見積書作成
    for anken_quote in anken_quotes:
        # 見積書作成
        # 品目をAPIで作成
        created_item_result = create_item(
            mfci_session, anken_quote.mfci_quote_item_json
        )
        # 空の見積書作成
        created_quote_result = create_quote(mfci_session, anken_quote.mfci_quote_json)
        # 最後に品目を見積書へ追加
        print(created_quote_result)
        attach_item_into_quote(
            mfci_session,
            created_quote_result["id"],
            created_item_result["id"],
        )

        # errorなら終了する
        if "errors" in created_quote_result:
            print("エラーが発生しました。詳細はレスポンスを確認ください")
            pprint(created_quote_result)
            sys.exit(0)

        # PDFのファイル名はミスミの型式をつける
        anken_quote.estimate_pdf_path = (
            export_qupte_dirpath / f"見積書_{anken_quote.anken_number}.pdf"
        )
        download_quote_pdf(
            mfci_session,
            created_quote_result["pdf_url"],
            anken_quote.estimate_pdf_path,
        )
        print(f"見積書のPDFをダウンロードしました。保存先:{anken_quote.estimate_pdf_path}")

        # - 生成後、今回選択したスプレッドシートは生成済みフォルダへ移動する
        # TODO: 2023-04-20 ここは関数として切り出す -> helper.googleapi
        try:
            previous_parents = ",".join(anken_quote.calcsheet_parents)

            _ = (
                gdrive_serivice.files()
                .update(
                    fileId=anken_quote.calcsheet_source,
                    addParents=MOVE_DIR_ID,
                    removeParents=previous_parents,
                    fields="id, parents",
                )
                .execute()
            )

        except HttpError as error:
            sys.exit(f"スプレッドシート移動時にエラーが発生しました: {error}")

        # スケジュール表の該当行に価格や納期を追加する
        msmanken_info = MsmAnkenMap(estimate_calcsheet_info=anken_quote)
        msmankenmaplist = MsmAnkenMapList()
        msmankenmaplist.msmankenmap_list.append(msmanken_info)

        export_pd = msmankenmaplist.generate_update_sheet_values()
        before_pd = get_schedule_table_area(table_search_range, sheet_service)
        update_data = generate_update_valueranges(
            table_search_range, before_pd, export_pd
        )
        print(f"update result:{update_data}")

        update_schedule_sheet(update_data, sheet_service)

    # メールの下書きを生成。案件のベース番号をもとにグルーピングをして一つのメールに複数の見積を添付する
    quote_groups = itertools.groupby(anken_quotes, lambda x: x.anken_base_number)

    for group_key, quote_iter in quote_groups:
        # quote_itemsをlistに変換する
        anken_quotes = list(quote_iter)

        # メール生成のテンプレは別のファイルに書く。
        # 納期はグループ内最初のQuoteItemのものを利用（案件に対して同じ納期を設定している前提）
        mail_template_body: str = SCRIPT_CONFIG.get("mail_template_body")
        replybody = mail_template_body.replace(
            "{{nouki}}", anken_quotes[0].duration_str
        )

        # メールのスレッドを取得して、スレッドに返信する
        searchquery = f"label:snd-ミスミ (*{group_key}*)"

        threads = api.googleapi.search_threads(gmail_service, searchquery)
        if not threads:
            sys.exit("スレッドが見つかりませんでした。メール返信作成を中止します。")

        # TODO:2023-04-18 ここは複数スレッドがあった場合は選択制にする。
        # 出ない場合は一番上のものを使いますと、タイトルを出して確認させる。
        # メッセージが大抵一つだが、一番上を取り出す（一番上が最新のはず）
        message = api.googleapi.get_messages_by_threadid(
            gmail_service, threads[0].get("id", "")
        )[0]

        # 返信メッセージで下書きを生成
        print(
            api.googleapi.append_draft_in_thread(
                gmail_service,
                replybody,
                (quote_item.estimate_pdf_path for quote_item in anken_quotes),
                message["id"],
                threads[0].get("id", ""),
            )
        )


if __name__ == "__main__":
    main()
