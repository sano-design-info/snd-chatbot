import json
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
from dateutil import parser
from dateutil.relativedelta import relativedelta
from googleapiclient.discovery import build
from openpyxl.styles import Border, Side
from zoneinfo import ZoneInfo

import chat.card
from api import googleapi
from api.mfcloud_api import (
    MFCIClient,
    attach_billingitem_into_billing,
    create_invoice_template_billing,
    create_item,
    download_billing_pdf,
    get_quotes,
)
from helper import EXPORTDIR_PATH, ROOTDIR, chatcard, load_config
from helper.regexpatterns import BILLING_DURARION, MSM_ANKEN_NUMBER
from task import BaseTask, ProcessData

# `2020-01-01` のフォーマットのみ受け付ける
START_DATE_FORMAT = "%Y-%m-%d"


config = load_config.CONFIG
MISUMI_TORIHIKISAKI_ID = config.get("mfci").get("TORIHIKISAKI_ID")
QUOTE_PER_PAGE = config.get("mfci").get("QUOTE_PER_PAGE")
SCRIPT_CONFIG = config.get("get_billing")

export_billing_dirpath = EXPORTDIR_PATH / "billing"
export_billing_dirpath.mkdir(parents=True, exist_ok=True)

# 本日の日付を全体で使うためにここで宣言
today_datetime = datetime.now()
BILLING_PDFFILEPATH = export_billing_dirpath / f"{today_datetime:%Y%m}_ミスミ配管請求書.pdf"
BILLING_LIST_EXCELPATH = (
    export_billing_dirpath / f"{today_datetime:%Y%m}_ミスミ配管納品一覧.xlsx"
)

# API Session
# TODO:2023-09-19 ここはtry exceptで囲む
google_cred = googleapi.get_cledential(googleapi.API_SCOPES)
gmail_service = build("gmail", "v1", credentials=google_cred)

# mfcloudのセッション作成
mfcl_session = MFCIClient().get_session()

# チャット用の認証情報を取得
google_sa_cred = googleapi.get_cledential_by_serviceaccount(googleapi.CHAT_API_SCOPES)
chat_service = build("chat", "v1", credentials=google_sa_cred)

spacename = config.get("google").get("CHAT_SPACENAME")
bot_header = chatcard.bot_header


def generate_dl_numbers(dl_numbers_str: str) -> list:
    """
    変換したい見積書の番号を指定したときに端的な番号にする
    例:1,2,3-5,7 => 1,2,3,4,5,7

    """
    dl_numbers = list()
    for dl_number_s in dl_numbers_str.split(","):
        if dl_number_s.isdigit():
            dl_numbers.append(int(dl_number_s))
        elif "-" in dl_numbers_str:
            dl_number_range = dl_number_s.split("-")
            now_number = int(dl_number_range[0])
            end_number = int(dl_number_range[1])
            while now_number <= end_number:
                dl_numbers.append(now_number)
                now_number += 1
    return dl_numbers


@dataclass
class QuoteData:
    """
    見積情報を簡素にまとめたデータ構造

    args:
        durarion_src: 見積書にある品目の詳細
        price: 見積書の合計金額
        hinmoku_title: 見積書にある品目のタイトル

    >>> qdd = {"durarion_src":"納期 10/20", "price":10000, "hinmoku_title":"MA-9901 ガススプリング配管図"}
    >>> qd = QuoteData(**qdd)

    """

    # durarionは最初にstrとして取り込んで、クラス生成時に正規表現で綺麗にする
    durarion_src: str = ""
    price: float = 0
    hinmoku_title: str = ""

    only_katasiki: str = field(init=False)
    durarion: str = field(init=False)

    def __post_init__(self):
        self.only_katasiki = MSM_ANKEN_NUMBER.match(self.hinmoku_title).group(0)
        self.durarion = BILLING_DURARION.match(self.durarion_src).group("durarion")


@dataclass
class BillingInfo:
    """作成する請求書情報のデータ構造"""

    price: float
    billing_title: str
    hinmoku_title: str


def generate_json_mfci_billing_item(billing_info: BillingInfo) -> dict:
    """
    請求情報を元に、MFクラウド請求書APIで使う請求書書向け品目用のjson文字列を生成する。
    結果は辞書形式で返す。
    """

    item_json_template = """
    {
        "name": "品目",
        "detail": "",
        "unit": "0",
        "price": 0,
        "quantity": 1,
        "excise": "ten_percent"
    }
    """

    # jsonでロードする
    billing_item = json.loads(item_json_template)

    # 結果をjsonで返す
    billing_item["name"] = "{} ガススプリング配管図".format(billing_info.hinmoku_title)
    billing_item["quantity"] = 1
    billing_item["price"] = int(billing_info.price)
    billing_item["name"] = billing_info.hinmoku_title

    return billing_item


def generate_billing_info_json(billing_info: BillingInfo) -> dict:
    """
    MFクラウド請求書APIで使う請求書作成のjsonを生成する
    結果は辞書形式で返す。
    """

    billing_json_template = """
    {
        "department_id": "",
        "billing_date": "2020-05-08",
        "due_date": "2020-05-08",
        "title": "請求書タイトル",
        "tags": "佐野設計自動生成",
        "note": "詳細は別添付の明細をご確認ください",
        "items": [
        ]
    }
    """
    # jsonでロードする
    billing_info_json = json.loads(billing_json_template)

    # department_idはミスミのものを利用
    billing_info_json["department_id"] = MISUMI_TORIHIKISAKI_ID

    # 各情報を入れる
    billing_info_json["title"] = billing_info.billing_title
    billing_info_json["billing_date"] = today_datetime.strftime(START_DATE_FORMAT)

    # due_dateは今月末にする
    # ref: https://zenn.dev/wtkn25/articles/python-relativedelta#%E6%9C%88%E5%88%9D%E3%80%81%E6%9C%88%E6%9C%AB%E3%80%81%E5%85%88%E6%9C%88%E5%88%9D%E3%80%81%E5%85%88%E6%9C%88%E6%9C%AB%E3%80%81%E7%BF%8C%E6%9C%88%E5%88%9D%E3%80%81%E7%BF%8C%E6%9C%88%E6%9C%AB
    billing_info_json["due_date"] = (
        today_datetime + relativedelta(months=+1, day=1, days=-1)
    ).strftime(START_DATE_FORMAT)

    return billing_info_json


# 見積書一覧を元にテンプレの行を生成
def generate_billing_pdf(mfci_session, billing_info: BillingInfo) -> Path:
    """
    MFクラウド請求書APIを使って請求書を生成する。戻り値はpdfファイルのパス
    """

    # 空の請求書作成
    billing_res = create_invoice_template_billing(
        mfci_session, generate_billing_info_json(billing_info)
    )
    # 品目を作成
    billing_item_data = generate_json_mfci_billing_item(billing_info)
    billing_item_res = create_item(mfci_session, billing_item_data)
    # 請求書へ品目を追加する
    attach_billingitem_into_billing(
        mfci_session, billing_res["id"], billing_item_res["id"]
    )

    print("請求書に変換しました")

    # pdfファイルのバイナリをgetする
    download_billing_pdf(mfci_session, billing_res["pdf_url"], BILLING_PDFFILEPATH)

    return BILLING_PDFFILEPATH


def set_border_style(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    num_rows: int,
    column_nums: list[int],
    start_num_row: int = 1,
):
    """
    ワークシートに罫線を入れる。
    """

    # 罫線の設定
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # B6からD6までのセルに上下左右に罫線をいれる
    for row in range(start_num_row, num_rows + start_num_row):
        for col in column_nums:
            cell = ws.cell(row=row, column=col)
            cell.border = border
            # D6の表示形式を日本円の通貨表記にする
            if col == 4:
                cell.number_format = "[$¥-ja-JP]#,##0;-[$¥-ja-JP]#,##0"


def generate_billing_list_excel(
    billing_target_quotes: list[QuoteData],
) -> Path:
    """
    請求対象一覧をxlsxファイルに出力する
    """
    # テンプレートの形式に沿った数字の設定
    row_start = 6
    # BillingTargetQuoteの構造に従ったマップ
    column_fieldmap = {
        "only_katasiki": 2,
        "durarion": 3,
        "price": 4,
    }

    wb = openpyxl.load_workbook(
        str((ROOTDIR / "itemparser/billing_list_template.xlsx"))
    )
    ws = wb.active

    # A1セルに"2023年**月請求一覧"と入れる。**は今月の月
    ws["B1"] = f"{today_datetime:%Y年%m月}請求一覧"
    # D1セルには今日の日付を入れる。フォーマットは"2023年4月26日" f-stringで入れる。
    ws["D1"] = f"{today_datetime:%Y年%m月%d日}"

    # 6行目からデータを入れる。
    for row, quote in enumerate(billing_target_quotes, start=row_start):
        for key, col in column_fieldmap.items():
            ws.cell(row=row, column=col).value = getattr(quote, key)

    # 罫線を入れる。set_border_styleを使う
    set_border_style(
        ws, len(billing_target_quotes), column_fieldmap.values(), row_start
    )

    # ファイルを保存する
    # TODO:2023-08-29 ここのファイル名の戻り値って意味ある？
    wb.save(BILLING_LIST_EXCELPATH)
    return BILLING_LIST_EXCELPATH


def str_to_datetime_with_dateutil(date_str):
    """
    日付の文字列 月/日 をdatetimeオブジェクトに変換する。
    現在の年を考慮して、今年の12月に1月または2月が入っていた場合、来年の1月または2月にする。

    注意:この考慮は少し曖昧で、12月だから1月2月が来年というよりは、現在の月より前の月だから来年という考え方が正しいはず。
    12月頃に通常発生する自然な表現として採用している。11月に1月2月が入っていたら、といったたらればは今は考えない。

    args:
        date_str: 日付の文字列 月/日
    return:
        datetimeオブジェクト
    """
    # 現在の年を取得
    current_year = datetime.now().year

    # 現在の月を取得
    current_month = datetime.now().month

    # dateutilを使用して日付文字列からdatetimeオブジェクトを生成
    # date = datetime.strptime(date_str, "%m/%d")
    date = parser.parse(date_str)

    # 現在の月が12月で、生成された日付の月が1月または2月の場合、年を+1する
    if current_month == 12 and (date.month == 1 or date.month == 2):
        date = date.replace(year=current_year + 1)

    return date


def is_target_date(t_date: datetime, start_date: datetime, end_date: datetime) -> bool:
    """
    指定した日付が、指定した日付の範囲内にあるかどうかを判定する
    args:
        t_date: 判定したい日付
        start_date: 範囲の開始日:end_dateより手前の日付を検証する
        end_date: 範囲の終了日
    return:
        t_dateがstart_dateとend_dateの間にあるかどうかの真偽値
    """
    # start_dateがend_dateより手前か確認する。違ったらエラー:valueerrorを出す
    if start_date > end_date:
        raise ValueError("start_dateはend_dateより前にしてください")

    return start_date.date() <= t_date.date() <= end_date.date()


# 請求書の金額合計のデータを生成する
# TODO:2023-10-16 count_quoteitemは使わないので、削除する
def generate_billing_data(
    quote_checked_list: list[QuoteData],
) -> BillingInfo:
    """
    請求書にする見積書の金額を合計する。
    表示用に件数とBillingInfoのタプルを返す

    args:
        quote_checked_list: 請求書にする見積書のリスト
    return:
        件数とBillingInfoのタプル

    """
    return BillingInfo(
        sum((i.price for i in quote_checked_list)),
        "ガススプリング配管図作製費",
        f"{today_datetime:%Y年%m月}請求分",
    )


# メール下書きを作成する
def set_draft_mail(attchment_filepaths: list[Path]) -> dict:
    """
    タイトルと本文を入力してメール下書きを作成する
    タイトルの例 "2023年03月請求書送付について"
    """
    # タイトルは日付が入ったもの。例:2023年03月請求書送付について
    mailtitle = SCRIPT_CONFIG.get("mail_template_title").replace(
        "{{datetime}}", f"{today_datetime:%Y年%m月}"
    )

    mailbody = SCRIPT_CONFIG.get("mail_template_body")
    mailto = SCRIPT_CONFIG.get("mail_to")
    mailcc = SCRIPT_CONFIG.get("mail_cc", "")

    # メール下書きを作成する
    return googleapi.append_draft(
        gmail_service, mailto, mailcc, mailtitle, mailbody, attchment_filepaths
    )


class PrepareTask(BaseTask):
    def execute_task(self) -> list[tuple[QuoteData, bool]]:
        # 見積書一覧を取得
        quote_result = get_quotes(
            mfcl_session,
            QUOTE_PER_PAGE,
        )

        # TODO:2023-10-16 [リファクタリング]日付のフィルターと、デフォルトチェックを同時に行うといいかも。
        # is_target_dateをis_date_rangeに変更する
        # is_date_rangeのstartとendを単体で取れるようにして、単体の時にはそれ以下、それ以上の判定を行うようにする

        # 取引先、実行時から40日前まででフィルター
        from_date = datetime.now(ZoneInfo("Asia/Tokyo")) + timedelta(days=-40)
        filtered_date_by_quote_result = (
            i
            for i in quote_result["data"]
            if i["department_id"] == MISUMI_TORIHIKISAKI_ID
            and datetime.fromisoformat(i["created_at"]) > from_date
        )

        # 見積一覧から必要情報を収集
        quote_list = [
            # 品目の最初の1行を使う
            QuoteData(
                durarion_src=quote["items"][0]["detail"],
                price=float(quote["subtotal_price"]),
                hinmoku_title=quote["items"][0]["name"],
            )
            for quote in filtered_date_by_quote_result
        ]

        # questionay.Choiceを使って見積書を選択する。
        # 納期（duration）を使って複数選択のデフォルト選択をマーク
        # 期日設定の 毎月26日から1か月前の日付をマーク
        # 例: 9/26締め切りの場合、8/26をマーク
        # 見積書の一覧を表示して、選択させる

        return [
            (
                quotedata,
                is_target_date(
                    str_to_datetime_with_dateutil(quotedata.durarion),
                    datetime(today_datetime.year, today_datetime.month, 26)
                    - relativedelta(months=1),
                    datetime(today_datetime.year, today_datetime.month, 27),
                ),
            )
            for quotedata in quote_list
        ]

    def execute_task_by_chat(self):
        result = self.execute_task()
        # 選択時のチェックボックス
        quotelist_checkbox = chat.card.genwidget_checkboxlist(
            "請求書にする見積書を選択してください。",
            "quoteitems",
            [
                chat.card.SelectionInputItem(
                    f"{quotedata.only_katasiki} | {quotedata.price} |{quotedata.durarion}",
                    json.dumps(asdict(quotedata)),
                    selected=checkbool,
                )
                for quotedata, checkbool in result
            ],
        )
        # 設定カードを生成
        config_body = chat.card.create_card(
            "config_card__get_billing",
            header=bot_header,
            widgets=[
                quotelist_checkbox,
                # ボタンを追加
                chat.card.genwidget_buttonlist(
                    [
                        chat.card.gencomponent_button(
                            "タスク実行確認", "confirm__get_billing"
                        ),
                        chat.card.gencomponent_button("キャンセル", "cancel_task"),
                    ]
                ),
            ],
        )
        return googleapi.create_chat_message(chat_service, spacename, config_body)


class MainTask(BaseTask):
    def execute_task(self, process_data: ProcessData | None = None) -> dict:
        ask_choiced_quote_list = process_data["task_data"].get("choiced_quote_list")
        billing_data = generate_billing_data(ask_choiced_quote_list)

        export_xlsx_path = generate_billing_list_excel(ask_choiced_quote_list)
        billing_pdf_path = generate_billing_pdf(mfcl_session, billing_data)
        print("一覧と請求書生成しました")
        print(f"一覧xlsxファイルパス:{export_xlsx_path}\n請求書pdf:{billing_pdf_path}")

        return set_draft_mail([export_xlsx_path, billing_pdf_path])

    # チャット用のタスクメソッド
    def execute_task_by_chat(
        self, process_data: ProcessData | None = None
    ) -> dict | str:
        result = self.execute_task(process_data)
        # チャット用のメッセージを作成する
        send_message_body = chat.card.create_card(
            "result_card__get_billing",
            header=bot_header,
            widgets=[
                chat.card.genwidget_textparagraph(
                    f"請求書と請求一覧を作成しました。: {result.get('id')}"
                ),
            ],
        )
        # send_message_body.update({"actionResponse": {"type": "NEW_MESSAGE"}})
        return googleapi.create_chat_message(chat_service, spacename, send_message_body)
