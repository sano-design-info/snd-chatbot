import base64
import html
import io
import itertools
import os
import os.path
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from pprint import pprint
import re

import copier
import dateutil.parser
import dateutil.tz
import questionary
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from jinja2 import Environment, FileSystemLoader
from dateutil.relativedelta import relativedelta
import openpyxl

from helper import google_api_helper

# load config
load_dotenv()
estimate_template_gsheet_id = os.environ.get("ESTIMATE_TEMPLATE_GSHEET_ID")
schedule_sheet_id = os.environ.get("SCHEDULE_SHEET_ID")
msm_gas_boilerplate_url = os.environ.get("MSM_GAS_BOILERPLATE_URL")
gsheet_tmmp_dir_ids = os.environ.get("GSHEET_TMP_DIR_IDS")

table_search_range = os.environ.get("TABLE_SEARCH_RANGE")

# Path
parent_dirpath = Path(__file__).parents[1]


def decode_base64url(s) -> bytes:
    return base64.urlsafe_b64decode(s) + b"=" * (4 - (len(s) % 4))


def encode_base64url(bytes_data) -> str:
    return base64.urlsafe_b64encode(bytes_data).rstrip(b"=")


def convert_gmail_datetimestr(gmail_datetimeformat: str) -> datetime:
    persed_time = dateutil.parser.parse(gmail_datetimeformat)
    return persed_time.astimezone(dateutil.tz.gettz("Asia/Tokyo"))


def pick_msm_katasiki_by_renrakukoumoku_filename(filepath: Path) -> str:
    # 配管連絡項目から必要な情報を取り出して、スケジュール表を更新
    # TODO:2022-12-14 ボイラープレートはExcelファイルのファイルパスから型式を取り出す

    # ボイラープレートと見積書作成時に必要になるミスミ型式番号を取得
    # 取得ができない場合は0000を用意
    msm_katasiki_num = "0000"
    if katasiki_matcher := re.match(
        r"MA-(\d{1,4}|\d{1,4}-\d{1})_",
        str(filepath.name),
    ):
        msm_katasiki_num = katasiki_matcher.group(1)

    return msm_katasiki_num


@dataclass
class ExpandedMessageItem:
    "メッセージ一覧の選択や、メール回りで使うときに利用する"
    gmail_message: dict
    payload: dict = field(init=False)
    headers: dict = field(init=False)

    id: str = field(init=False)
    title: str = field(init=False)
    from_addresss: str = field(init=False)
    cc_address: str = field(init=False)
    datetime_: datetime = field(init=False)
    body_related: dict = field(init=False)
    body_parts: dict = field(init=False)

    def __post_init__(self):
        self.payload = self.gmail_message.get("payload")
        self.headers = self.payload.get("headers")

        self.id = self.gmail_message.get("id")
        self.title = next((i for i in self.headers if i.get("name") == "Subject")).get(
            "value"
        )

        self.from_addresss = next(
            (i for i in self.headers if i.get("name") == "From")
        ).get("value")

        self.cc_address = next((i for i in self.headers if i.get("name") == "CC")).get(
            "value"
        )

        self.datetime_ = convert_gmail_datetimestr(
            next((i for i in self.headers if i.get("name") == "Date")).get("value")
        )

        # メールのmimeマルチパートを考慮して、構造が違うモノに対応する
        # relative> altanative の手順で掘る。mixedは一番上で考慮しているのでここでは行わない
        self.body_related = {}
        self.body_parts = next(
            (
                i
                for i in self.payload.get("parts")
                if i.get("mimeType") == "multipart/alternative"
            ),
            {},
        ).get("parts", [])
        if not self.body_parts:
            self.body_related = next(
                (
                    i
                    for i in self.payload.get("parts")
                    if i.get("mimeType") == "multipart/related"
                )
            )
            self.body_parts = next(
                (
                    i
                    for i in self.body_related.get("parts")
                    if i.get("mimeType") == "multipart/alternative"
                )
            ).get("parts")


def generate_mail_printhtml(
    messageitem: ExpandedMessageItem, export_dirpath: Path
) -> None:
    # メール印刷用HTML生成
    # TODO:2022-12-14 ここではメールのheader, payloadがあればよさそう。ExpandedMessageItemが引っ張れればよさそうかな
    messages_text_parts = [
        i for i in messageitem.body_parts if "text" in i.get("mimeType")
    ]

    b64dec_msg_byte = decode_base64url(messages_text_parts[1].get("body").get("data"))

    # imgタグを除去する
    mail_html_bs4 = BeautifulSoup(b64dec_msg_byte, "html.parser")
    only_body_tags = mail_html_bs4.body
    for t in only_body_tags.find_all("img"):
        t.decompose()

    # jinja2埋込
    # テンプレート読み込み
    env = Environment(
        loader=FileSystemLoader(str((parent_dirpath / "prepare")), encoding="utf8")
    )
    tmpl = env.get_template("export.html.jinja")

    # 設定ファイル読み込み
    params = {
        "export_html": only_body_tags,
        "message_title": html.escape(messageitem.title),
        "message_from_addresss": html.escape(messageitem.from_addresss),
        "message_cc_address": html.escape(messageitem.cc_address),
        "message_datetime": messageitem.datetime_,
    }
    # レンダリングして出力
    with (export_dirpath / Path("./export_mail.html")).open(
        "w", encoding="utf8"
    ) as exp_mail_hmtl:
        exp_mail_hmtl.write(tmpl.render(params))

    pass


def generate_pdf_by_renrakukoumoku_excel(
    attachment_dirpath: Path,
    export_dirpath: Path,
    google_creds: Credentials,
) -> None:
    # 連絡項目の印刷用PDFファイル生成
    mimetype_xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    target_filepath = next(attachment_dirpath.glob("*MA-*.xlsx"))

    if not target_filepath:
        print("cant generate_pdf_byrenrakuexcel")
        return None

    # ExcelファイルをPDFに変換する
    media = MediaFileUpload(
        target_filepath,
        mimetype=mimetype_xlsx,
        resumable=True,
    )

    file_metadata = {
        "name": target_filepath.name,
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents": gsheet_tmmp_dir_ids,
    }

    try:
        drive_service = build("drive", "v3", credentials=google_creds)

        # Excelファイルのアップロードを行って、そのアップロードファイルをPDFで保存できるかチェック
        upload_results = (
            drive_service.files().create(body=file_metadata, media_body=media).execute()
        )

        print(upload_results)
        # print(upload_results.get("id"))

        # pdfファイルを取りに行ってみる
        dl_request = drive_service.files().export_media(
            fileId=upload_results.get("id"), mimeType="application/pdf"
        )
        file = io.BytesIO()
        downloader = MediaIoBaseDownload(file, dl_request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print(f"Download {int(status.progress() * 100)}.")

        with (export_dirpath / Path("./export_excel.pdf")).open(
            "wb"
        ) as export_exceltopdf:
            export_exceltopdf.write(file.getvalue())

        # print("[Post Process...]")
        # post-porcess: Googleドキュメントに一時保持した配管連絡項目を除去する
        delete_tmp_excel_result = (
            drive_service.files()
            .delete(fileId=upload_results.get("id"), fields="id")
            .execute()
        )
        print("deleted file: ", delete_tmp_excel_result)

    except HttpError as error:
        # TODO:2022-12-09 エラーハンドリングは基本行わずここで落とすこと
        print(f"An error occurred: {error}")


def generate_projectdir(attachment_dirpath: Path, export_dirpath: Path) -> None:

    export_project_dir = export_dirpath / "proj_dir"
    export_project_dir.mkdir(exist_ok=True)

    # ファイルパスから型式を取り出す
    target_filepath = next(attachment_dirpath.glob("*MA-*.xlsx"))
    if not target_filepath:
        print("cant create projectdir")
        return None

    msm_katasiki_num = pick_msm_katasiki_by_renrakukoumoku_filename(target_filepath)

    # ボイラープレートからディレクトリ生成
    boilerplate_config = {
        "project_name": msm_katasiki_num,
        "haikan_pattern": "type_s",
    }
    copier.run_copy(
        msm_gas_boilerplate_url,
        export_project_dir,
        data=boilerplate_config,
    )

    # TODO:2022-12-14 添付ファイルをコピーする


def add_schedule_spreadsheet(
    attachment_dirpath: Path, google_creds: Credentials
) -> None:
    target_filepath = next(attachment_dirpath.glob("*MA-*.xlsx"))
    if not target_filepath:
        print("cant add schedule")
        return None

    msm_katasiki_num = pick_msm_katasiki_by_renrakukoumoku_filename(target_filepath)

    # エンドユーザー: 今は列がないので登録しない
    renrakukoumoku_range_enduser = "D10"
    # 顧客
    renrakukoumoku_range_kokyaku = "D6"

    # openpyxlで必要な位置から値を取り出す
    renrakukoumoku_wb = openpyxl.load_workbook(target_filepath)
    renrakukoumoku_ws = renrakukoumoku_wb.active
    add_schedule_kokyaku = renrakukoumoku_ws[renrakukoumoku_range_kokyaku].value

    renrakukoumoku_wb.close()

    # 型式
    add_schedule_msm_katasiki = f"MA-{msm_katasiki_num}"
    # 開始日: 実行日でよし
    now_datetime = datetime.now()
    add_schedule_start_datetime = now_datetime.strftime("%Y/%m/%d")
    # 振込タイミング:開始日の来月を基本にする。ずれる場合は手動で修正する
    add_schedule_hurikomiduki = now_datetime.replace(day=1) + relativedelta(months=1)

    append_values = [
        [
            "=row() - 5",
            "ミスミ",
            add_schedule_msm_katasiki,
            "友希",
            "",
            add_schedule_start_datetime,
            "",
            "",
            "",
            f"{add_schedule_hurikomiduki.month}月末",
            "",
            "",
            add_schedule_kokyaku,
        ]
    ]

    sheet_service = build("sheets", "v4", credentials=google_creds)

    # print(schedule_sheet_id, " ", table_search_range, "", )
    # print(append_values)

    # スケジュール表の一番後ろの行へ追加する
    schedule_gsheet = (
        sheet_service.spreadsheets()
        .values()
        .append(
            spreadsheetId=schedule_sheet_id,
            range=table_search_range,
            body={"values": append_values},
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
        )
    ).execute()

    print(f"schedule updated -> {schedule_gsheet.get('updates')}")


def generate_estimate_calcsheet(
    attachment_dirpath: Path, google_creds: Credentials
) -> None:

    # 見積書生成機能をここで動かす。別のライブラリ化しておいて、それをここで呼び出すで良いと思う。
    # TODO:2022-12-15 ここは情緒なんだけど外に出す必要性があんまりないので今はそのまま。

    target_filepath = next(attachment_dirpath.glob("*MA-*.xlsx"))
    if not target_filepath:
        print("cant add schedule")
        return None

    msm_katasiki_num = pick_msm_katasiki_by_renrakukoumoku_filename(target_filepath)

    try:
        drive_service = build("drive", "v3", credentials=google_creds)

        copy_template_results = (
            drive_service.files()
            .copy(fileId=estimate_template_gsheet_id, fields="id,name")
            .execute()
        )

        # コピー先のファイル名を変更する
        template_suffix = "[ミスミ型番] のコピー"
        renamed_filename = copy_template_results.get("name").replace(
            template_suffix, msm_katasiki_num
        )
        rename_body = {"name": renamed_filename}
        rename_estimate_gsheet_result = (
            drive_service.files()
            .update(
                fileId=copy_template_results.get("id"),
                body=rename_body,
                fields="name",
            )
            .execute()
        )

        print(
            f"estimate copy and renamed -> {rename_estimate_gsheet_result.get('name')}"
        )
    except HttpError as error:
        # TODO:2022-12-09 エラーハンドリングは基本行わずここで落とすこと
        print(f"An error occurred: {error}")
