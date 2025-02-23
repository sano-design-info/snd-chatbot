import itertools
import json
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
# from pprint import pprint

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import chat.card
from api import googleapi
from api.googleapi import sheet_data_mapper

from helper import EXPORTDIR_PATH, chatcard, load_config
from itemparser import (
    EstimateCalcSheetInfo,
    MsmAnkenMap,
    MsmAnkenMapList,
    generate_update_valueranges,
    get_schedule_table_area,
    update_schedule_sheet,
)
from task import BaseTask, ProcessData

# load config, credential
config = load_config.CONFIG

# TODO:2024-02-05 この設定はitemparserに移動する
schedule_spreadsheet_table_range = config.get("general").get(
    "SCHEDULE_SPREADSHEET_TABLE_RANGE"
)
TORIHIKISAKI_NAME = config.get("general").get("TORIHIKISAKI_NAME")

SCRIPT_CONFIG = config.get("generate_quotes")
ESTIMATE_CALCSHEET_DIR_IDS = SCRIPT_CONFIG.get("ESTIMATE_CALCSHEET_DIR_IDS")
ARCHIVED_ESTIMATECALCSHEET_DIR_IDS = SCRIPT_CONFIG.get(
    "ARCHIVED_ESTIMATECALCSHEET_DIR_IDS"
)
MAIL_TEMPLATE_BODY_STR = SCRIPT_CONFIG.get("mail_template_body")

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

# TODO:2023-09-14 これは使っている部分へ戻す。これ以外で使っていないので、ここで定義する必要はない
# 2020-01-01 のフォーマットのみ受け付ける
START_DATE_FORMAT = "%Y-%m-%d"

# API Session
# TODO:2023-09-19 ここはtry exceptで囲む
# Googleのtokenを用意
google_cred = googleapi.get_cledential(googleapi.API_SCOPES)
gdrive_service = build("drive", "v3", credentials=google_cred)
gmail_service = build("gmail", "v1", credentials=google_cred)
gsheet_service = build("sheets", "v4", credentials=google_cred)

# チャット用の認証情報を取得
google_sa_cred = googleapi.get_cledential_by_serviceaccount(googleapi.CHAT_API_SCOPES)
chat_service = build("chat", "v1", credentials=google_sa_cred)

spacename = config.get("google").get("CHAT_SPACENAME")
bot_header = chatcard.bot_header


# TODO:2023-10-15 このデータクラスをasdictすると、EstimateCalcSheetInfoのgapiのserviseが変換できないと思われる。
# その時は、dataclass.InitVarを使って、初期化限定変数を作ればいいらしい
@dataclass
class AnkenQuote(EstimateCalcSheetInfo):
    """
    見積書の各案件事のグループ化を行うためのクラス
    """

    quote_pdf_path: Path = field(init=False, default=None)
    quote_gsheet_data: dict = field(init=False, default=None)
    updated_quote_manage_cell_address: str = field(init=False, default=None)

    def __post_init__(self):
        # 継承元のpost_initを呼ぶ
        super().__post_init__()
        # 収集した見積もり計算書の情報を元に見積書用のjsonを生成する
        # self._convert_json_mfci_quote_item()
        # self._convert_json_mfci_quote()

    def print_quote_info(self) -> None:
        """
        生成した見積情報をワンラインで表示する
        """
        print(
            f"見積情報: 型式:{self.anken_number} 日時:{self.duration} 価格:{self.price}"
        )

    def convert_dict_to_gsheet_tamplate(self, quote_number) -> None:
        # 見積書へ書き込むデータを作る
        # ここで作成するデータは、見積書のテンプレートに合わせたデータを作成する

        hinmoku = {
            "name": f"{self.anken_number} ガススプリング配管図",
            "detail": f"納期 {self.duration:%m/%d}",
            "price": int(self.price),
            "quantity": 1,
            "zeiritu": "10%",
        }

        # LRは条件判断を行う
        quote_note = ""
        rh_flag = self.anken_number.split("-")[-1]
        if rh_flag in ("RH", "LH") != 0:
            # RHの場合はLH, LHの場合はRHの備考文章を作成
            reverse_part_number = "MA-" + "-".join(self.anken_number.split("-")[0:-1])
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
        self.quote_gsheet_data = {
            "customer_name": TORIHIKISAKI_NAME,
            "quote_id": qupte_id,
            "title": "ガススプリング配管図作製費",
            # 日付は実行時の日付を利用
            "quote_date": today_datetime.strftime(START_DATE_FORMAT),
            "note": quote_note,
            "item_table": [hinmoku],
        }


def generate_anken_quote_list(estimate_calcsheets: list[dict]) -> list[AnkenQuote]:
    """
    ミスミの配管計算表を元に見積もりを作成するためのAnkenQuoteのリストを作成する

    Args:
        estimate_calcsheets (list[dict]): ミスミの配管計算表の情報
    return:
        list[AnkenQuote]: 見積もりを作成するためのAnkenQuoteのリスト
    """
    anken_quotes: list[AnkenQuote] = []
    for estimate_calcsheet in estimate_calcsheets:
        anken_quote = AnkenQuote(gsheet_service, estimate_calcsheet.get("id"))
        anken_quote.calcsheet_parents = estimate_calcsheet.get("parents")
        # itemをリストアップ
        anken_quotes.append(anken_quote)
    return anken_quotes


def update_msm_anken_schedule_sheet(
    anken_quote: AnkenQuote, gsheet_service
) -> list[dict]:
    """
    ミスミのスケジュール表を更新する
    Args:
        anken_quote (AnkenQuote): 見積もり情報
        gsheet_service (googleapiclient.discovery.Resource): Google Sheet APIのサービス
    return:
        list[dict]: 更新結果
    """
    msmanken_info = MsmAnkenMap(estimate_calcsheet_info=anken_quote)
    msmankenmaplist = MsmAnkenMapList()
    msmankenmaplist.msmankenmap_list.append(msmanken_info)

    export_pd = msmankenmaplist.generate_update_sheet_values()

    # TODO:2024-02-05 ここのtable_rangeは元のモジュールから呼び出したほうがシンプル
    before_pd = get_schedule_table_area(
        schedule_spreadsheet_table_range, gsheet_service
    )
    update_data = generate_update_valueranges(
        schedule_spreadsheet_table_range, before_pd, export_pd
    )

    print(f"update result:{update_data}")

    return update_schedule_sheet(update_data, gsheet_service)


class PrepareTask(BaseTask):
    def execute_task(self):
        # TODO:2023-09-28 [prepare start]

        # google sheetのリストを取得
        # - 特定のフォルダ（GSheet的にはグループ）の一覧を取得
        # - テンプレフォルダ内もフィルターにいれる。その時にテンプレートは除外する
        query_by_estimate_calcsheet = f"""
            ({" or ".join((f"'{id}' in parents " for id in ESTIMATE_CALCSHEET_DIR_IDS))})
            and mimeType = 'application/vnd.google-apps.spreadsheet'
            and trashed = false
            and name != "ミスミ配管図見積り計算表v2_MA-[ミスミ型番]"
        """

        try:
            estimate_calcsheet_list = list(
                reversed(
                    googleapi.get_file_list(
                        gdrive_service,
                        query_by_estimate_calcsheet,
                        page_size=10,
                        fields="files(id, name, parents)",
                    ).get("files", [])
                )
            )

        except HttpError as error:
            sys.exit(f"Google Drive APIのエラーが発生しました。: {error}")

        if not estimate_calcsheet_list:
            print("見積もり計算表が見つかりませんでした。終了します。")
            sys.exit(0)
        return estimate_calcsheet_list
        # TODO: 2023-10-13 この戻り値はjson形式なので、チャットのタスクとしてそのまま利用する

    def execute_task_by_chat(self):
        result = self.execute_task()

        if result is None:
            print("見積もり計算表が見つかりませんでした。終了します。")
            # google chatのエラーとして返す
            return chat.card.genactionresponse_dialog(
                "見積もり計算表が見つかりませんでした。終了します。"
            )

        estimate_list_checkbox = chat.card.genwidget_checkboxlist(
            "見積もり一覧",
            "estimate_list_checkbox",
            [
                chat.card.SelectionInputItem(
                    estimate_calcsheet.get("name"), json.dumps(estimate_calcsheet)
                )
                for estimate_calcsheet in result
            ],
        )

        # 設定カードのボディを生成
        config_body = chat.card.create_card(
            "config_card__generate_quote",
            header=bot_header,
            widgets=[
                estimate_list_checkbox,
                # ボタンを追加
                chat.card.genwidget_buttonlist(
                    [
                        chat.card.gencomponent_button(
                            "タスク実行", "run_task__generate_quotes"
                        ),
                        chat.card.gencomponent_button("キャンセル", "cancel_task"),
                    ]
                ),
            ],
        )
        return googleapi.create_chat_message(chat_service, spacename, config_body)


class MainTask(BaseTask):
    def execute_task(self, process_data: ProcessData | None = None):
        # 渡されたデータを展開する
        selected_estimate_calcsheets = process_data["task_data"].get(
            "selected_estimate_calcsheets"
        )

        # 一連の操作中に使うデータ構造を入れるリスト（グループ化はメール生成時に行う）
        anken_quotes: list[AnkenQuote] = generate_anken_quote_list(
            selected_estimate_calcsheets
        )

        # MFクラウドで見積書作成
        for anken_quote in anken_quotes:
            try:
                # [見積書作成を行う]
                # 見積書管理表から番号を生成
                updated_quote_manage_gsheet = googleapi.append_sheet(
                    gsheet_service,
                    QUOTE_FILE_LIST_GSHEET_ID,
                    "見積書管理",
                    [['=TEXT(ROW()-1,"0000")', "", "", ""]],
                    "USER_ENTERED",
                    "INSERT_ROWS",
                    True,
                )
                print(
                    f"見積書の管理表に番号を生成しました。: {updated_quote_manage_gsheet}"
                )
                # 見積書番号を取得
                quote_id = (
                    updated_quote_manage_gsheet.get("updates")
                    .get("updatedData")
                    .get("values")[0][0]
                )
                print(f"見積書の管理表から番号を生成しました。: {quote_id}")

                # 見積書の情報を生成
                anken_quote.convert_dict_to_gsheet_tamplate(quote_id)

                # googleスプレッドシートの見積書テンプレートを複製する
                quote_file_id = googleapi.dupulicate_file(
                    gdrive_service,
                    QUOTE_TEMPLATE_GSHEET_ID,
                    quote_filestem := f"見積書_{anken_quote.anken_number}",
                )

                # 見積書のファイル名と保存先を設定
                _ = googleapi.update_file(
                    gdrive_service,
                    file_id=quote_file_id,
                    body=None,
                    add_parents=QUOTE_GSHEET_SAVE_DIR_IDS,
                    fields="id, parents",
                )

                # 見積書へanken_quoteの内容を記録
                sheet_data_mapper.write_data_to_sheet(
                    gsheet_service,
                    quote_file_id,
                    anken_quote.quote_gsheet_data,
                    quote_template_cell_mapping_dict,
                )

                # ファイル名:見積書_[納期].pdf
                quote_filename = f"{quote_filestem}.pdf"
                anken_quote.quote_pdf_path = export_quote_dirpath / quote_filename

                # 見積書のPDFをダウンロード
                googleapi.export_pdf_by_driveexporturl(
                    google_cred.token,
                    quote_file_id,
                    anken_quote.quote_pdf_path,
                    {
                        "gid": "0",
                        "size": "7",
                        "portrait": "true",
                        "fitw": "true",
                        "gridlines": "false",
                    },
                )

                # 見積書のPDFをGoogleドライブへ保存
                upload_pdf_result = googleapi.upload_file(
                    gdrive_service,
                    anken_quote.quote_pdf_path,
                    "application/pdf",
                    "application/pdf",
                    QUOTE_PDF_SAVE_DIR_IDS,
                )

                # 見積書のGoogleスプレッドシートとPDFのURLを見積管理表に記録

                # 見積管理表を更新する。B列から[ファイル名, 見積書:Gsheet のIDからURL, 見積書:GDrive PDFのIDからURL]
                # 見積管理表に番号を追加したupdatedRows（updated_quote_manage_cell_address）を使うがB列以降を使う
                # 生成した見積番号のセルアドレスからB列に置き換えて取得。AのみをBにする
                # 例: updated_quote_manage_cell_address = "見積書管理!A2:D2" -> "見積書管理!B2:D2"
                _ = googleapi.update_sheet(
                    gsheet_service,
                    QUOTE_FILE_LIST_GSHEET_ID,
                    updated_quote_manage_gsheet.get("updates")
                    .get("updatedRange")
                    .replace("A", "B"),
                    [
                        [
                            quote_filename,
                            f"http://docs.google.com/spreadsheets/d/{quote_file_id}",
                            f"http://drive.google.com/file/d/{upload_pdf_result.get('id')}",
                        ]
                    ],
                )
                # TODO:2024-02-06 ここのestimate_pdf_pathは変数名が微妙なので、quote_pdf_pathとして変更する。影響範囲を確認すること
                print(
                    f"見積書のPDFをダウンロードしました。保存先:{anken_quote.quote_pdf_path}"
                )
                # 見積書生成後、今回選択した見積計算書スプレッドシートは生成済みフォルダへ移動する
                _ = googleapi.update_file(
                    gdrive_service,
                    file_id=anken_quote.calcsheet_source,
                    add_parents=ARCHIVED_ESTIMATECALCSHEET_DIR_IDS,
                    remove_parents=anken_quote.calcsheet_parents,
                    fields="id",
                )

            except HttpError as error:
                sys.exit(f"見積書生成中にエラーが発生しました: {error}")

            # スケジュール表の該当行に価格や納期を追加する
            update_msm_anken_schedule_sheet(anken_quote, gsheet_service)

        # TODO:2023-09-28 下書き生成は、上の見積書が生成できたら実行するタスクになる。
        # TODO: 2024-02-06 以前に複数の案件ベース番号があった時に、二番めのメールが作成できなかったことがあるので検証する

        # メールの下書きを生成。案件のベース番号をもとにグルーピングをして一つのメールに複数の見積を添付する
        quote_groups = itertools.groupby(anken_quotes, lambda x: x.anken_base_number)

        for group_key, quote_iter in quote_groups:
            # メールのスレッドを取得して、スレッドに返信する
            threads = googleapi.search_threads(
                gmail_service, f"label:snd-ミスミ (*{group_key}*)"
            )
            # スレッドが見つからない場合は終了
            if not threads:
                print("スレッドが見つかりませんでした。メール返信作成を中止します。")
                return {
                    "result": "スレッドが見つかりませんでした。メール返信作成を中止します。"
                }

            # TODO:2023-04-18 ここは複数スレッドがあった場合は選択制にする。
            # 出ない場合は一番上のものを使いますと、タイトルを出して確認させる。
            # メッセージが大抵一つだが、一番上を取り出す（一番上が最新のはず）
            message = googleapi.get_messages_by_threadid(
                gmail_service, threads[0].get("id", "")
            )[0]

            # メールの必要な情報を生成する
            # quote_itemsをlistに変換する
            anken_quotes = list(quote_iter)

            # メール生成のテンプレは別のファイルに書く。
            # 納期はグループ内最初のQuoteItemのものを利用（案件に対して同じ納期を設定している前提）
            mail_template_body: str = MAIL_TEMPLATE_BODY_STR
            replybody = mail_template_body.replace(
                "{{nouki}}", anken_quotes[0].duration_str
            )

            # 返信メッセージで下書きを生成
            return googleapi.append_draft_in_thread(
                gmail_service,
                replybody,
                (quote_item.quote_pdf_path for quote_item in anken_quotes),
                message["id"],
                threads[0].get("id", ""),
            )

    # チャット用のタスクメソッド
    def execute_task_by_chat(self, process_data: ProcessData | None = None):
        result = self.execute_task(process_data)
        # チャット用のメッセージを作成する
        send_message_body = chat.card.create_card(
            "result_card__generate_quote",
            header=bot_header,
            widgets=[
                chat.card.genwidget_textparagraph(
                    f"見積書を作成しました。: {result.get('id')}"
                ),
            ],
        )
        # send_message_body.update({"actionResponse": {"type": "NEW_MESSAGE"}})
        return googleapi.create_chat_message(chat_service, spacename, send_message_body)
