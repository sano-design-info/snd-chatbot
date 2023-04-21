import base64
import html
import io
import re
import shutil
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import copier
import dateutil.parser
import dateutil.tz
import openpyxl
from bs4 import BeautifulSoup
from dateutil.relativedelta import relativedelta
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from jinja2 import Environment, FileSystemLoader

from helper import extract_zip, load_config

# load config

config = load_config.CONFIG
estimate_template_gsheet_id = config.get("google").get("ESTIMATE_TEMPLATE_GSHEET_ID")
schedule_sheet_id = config.get("google").get("SCHEDULE_SHEET_ID")
msm_gas_boilerplate_url = config.get("other").get("MSM_GAS_BOILERPLATE_URL")
gsheet_tmmp_dir_ids = config.get("google").get("GSHEET_TMP_DIR_IDS")
table_search_range = config.get("google").get("TABLE_SEARCH_RANGE")
copy_project_dir_dest_path = Path(config.get("other").get("COPY_PROJECT_DIR_DEST_PATH"))

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
    subject: str = field(init=False)
    from_address: str = field(init=False)
    to_address: str = field(init=False)
    cc_address: str = field(init=False)
    datetime_: datetime = field(init=False)
    body_related: dict = field(init=False)
    body_parts: dict = field(init=False)
    body: str = field(init=False)

    def __post_init__(self):
        self.payload = self.gmail_message.get("payload")
        self.headers = self.payload.get("headers")

        self.id = self.gmail_message.get("id")
        self.title = next((i for i in self.headers if i.get("name") == "Subject")).get(
            "value"
        )
        self.subject = self.title

        self.from_address = next(
            (i for i in self.headers if i.get("name") == "From")
        ).get("value")

        self.to_address = next((i for i in self.headers if i.get("name") == "To")).get(
            "value"
        )

        # CCのアドレスがある場合は、CCのアドレスを取得する。無い場合は空文字を入れる
        self.cc_address = next(
            (i for i in self.headers if i.get("name") == "CC"), {}
        ).get("value", "")

        self.datetime_ = convert_gmail_datetimestr(
            next((i for i in self.headers if i.get("name") == "Date")).get("value")
        )

        # メールのmimeマルチパートを考慮して、構造が違うモノに対応する
        # メールがリッチテキストかつimgファイルがある場合は、multipart/relatedとなり、body_relatedを入れるとimgファイル収集も可能なので、別で用意している
        self.body_related = {}

        # ここはtext plane or multipart/altanative or multipart/related >  multipart/altanative の構造になってるらしいので、分離した処理に切り替えないといけない

        # partsがない場合 = シンプルなテキストベースの場合
        if not self.payload.get("parts"):
            self.body_parts = [self.payload]
        else:
            # リッチテキスト系の場合
            mail_part_mimetype = next(
                i.get("mimeType")
                for i in self.payload.get("parts")
                if i.get("partId") in ("0")
            )

            match mail_part_mimetype:
                case "text/plain":
                    self.body_parts = self.payload.get("parts")
                case "multipart/alternative":
                    self.body_parts = next(
                        (
                            i
                            for i in self.payload.get("parts")
                            if i.get("mimeType") == "multipart/alternative"
                        ),
                        {},
                    ).get("parts")
                case "multipart/related":
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
                case _:
                    pass

        # body_partsからbodyを取得する
        mailbody = next(
            (
                i["body"]["data"]
                for i in self.body_parts
                if "text/plain" in i.get("mimeType")
            )
        )
        self.body = decode_base64url(mailbody).decode("utf8")


def generate_mail_printhtml(
    messageitem: ExpandedMessageItem, attachment_dirpath: Path
) -> None:
    # メール印刷用HTML生成

    # TODO:2023-04-07 この部分はExpandMessageItemへ移動する
    # TODO:2022-12-27 個々の実装はhtmlかplaneで最初から分けたほうがいいかも
    messages_text_parts = next(
        (i for i in messageitem.body_parts if "text/html" in i.get("mimeType")), None
    )
    if not messages_text_parts:
        messages_text_parts = next(
            (i for i in messageitem.body_parts if "text/plain" in i.get("mimeType"))
        )

    # imgタグを除去する
    mail_body = ""
    b64decoded_mail_body = decode_base64url(messages_text_parts.get("body").get("data"))

    mail_html_bs4 = BeautifulSoup(b64decoded_mail_body, "html.parser")

    if mail_html_bs4.body:
        mail_body = mail_html_bs4.body
        for t in mail_body.find_all("img"):
            t.decompose()
    else:
        # decodeしないで、改行タグをhtmlの<br>へ置き換え
        mail_body = b64decoded_mail_body.decode("utf8").replace("\r\n", "<br>")

    # jinja2埋込
    # テンプレート読み込み
    env = Environment(
        loader=FileSystemLoader(str((parent_dirpath / "prepare")), encoding="utf8")
    )
    tmpl = env.get_template("export.html.jinja")

    # 設定ファイル読み込み
    params = {
        "export_html": mail_body,
        "message_title": html.escape(messageitem.title),
        "message_from_address": html.escape(messageitem.from_address),
        "message_cc_address": html.escape(messageitem.cc_address),
        "message_datetime": messageitem.datetime_,
    }
    # レンダリングして出力
    with (attachment_dirpath / Path("./メール本文印刷用ファイル.html")).open(
        "w", encoding="utf8"
    ) as exp_mail_hmtl:
        exp_mail_hmtl.write(tmpl.render(params))


def generate_pdf_by_renrakukoumoku_excel(
    attachment_dirpath: Path,
    export_dirpath: Path,
    google_cred: Credentials,
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
        drive_service = build("drive", "v3", credentials=google_cred)

        # Excelファイルのアップロードを行って、そのアップロードファイルをPDFで保存できるかチェック
        upload_results = (
            drive_service.files().create(body=file_metadata, media_body=media).execute()
        )
        print(upload_results)

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

        with (attachment_dirpath / Path("./連絡項目印刷用ファイル.pdf")).open(
            "wb"
        ) as export_exceltopdf:
            export_exceltopdf.write(file.getvalue())

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
    generated_copier_worker = copier.run_copy(
        msm_gas_boilerplate_url,
        export_project_dir,
        data=boilerplate_config,
    )

    # 添付ファイルを解凍する
    for attachment_zfile in attachment_dirpath.glob("*.zip"):
        print(attachment_zfile)
        extract_zip.extract_zipfile(attachment_zfile, attachment_dirpath)

    # 添付ファイルをコピーする
    # TODO:2022-12-16 プロジェクトフォルダの名称は環境変数化したほうがいいかも？
    project_dir = generated_copier_worker.dst_path / f"ミスミ配管図MA-{msm_katasiki_num}納期 -"
    shutil.copytree(attachment_dirpath, (project_dir / "資料"), dirs_exist_ok=True)


def add_schedule_spreadsheet(
    attachment_dirpath: Path, google_cred: Credentials, nyukin_nextmonth: bool = False
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
    # 振込タイミング
    nyukin_month = 1
    if nyukin_nextmonth:
        nyukin_month = 2
    add_schedule_hurikomiduki = now_datetime.replace(day=1) + relativedelta(
        months=nyukin_month
    )

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

    sheet_service = build("sheets", "v4", credentials=google_cred)

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
    attachment_dirpath: Path, google_cred: Credentials
) -> None:

    # 見積書生成機能をここで動かす。別のライブラリ化しておいて、それをここで呼び出すで良いと思う。
    # TODO:2022-12-15 ここは情緒なんだけど外に出す必要性があんまりないので今はそのまま。

    target_filepath = next(attachment_dirpath.glob("*MA-*.xlsx"))
    if not target_filepath:
        print("cant add schedule")
        return None

    msm_katasiki_num = pick_msm_katasiki_by_renrakukoumoku_filename(target_filepath)

    try:
        drive_service = build("drive", "v3", credentials=google_cred)

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


def copy_projectdir(export_path: Path) -> None:
    # プロジェクトフォルダのパスを用意する
    export_prij_path = next((export_path / "proj_dir").glob("ミスミ配管図*"))
    print(export_prij_path)
    print(copy_project_dir_dest_path)
    shutil.copytree(
        export_prij_path,
        copy_project_dir_dest_path / export_prij_path.name,
        dirs_exist_ok=True,
    )
    pass
