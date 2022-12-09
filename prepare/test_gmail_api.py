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
import questionary
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from jinja2 import Environment, FileSystemLoader

# load config
load_dotenv()

cred_filepath = os.environ.get("CRED_FILEPATH")
target_userid = os.environ.get("GMAIL_USER_ID")

msm_gas_boilerplate_url = os.environ.get("MSM_GAS_BOILERPLATE_URL")

mimetype_gsheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# If modifying these scopes, delete the file token.json.
SCOPES = [
    "https://mail.google.com/",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.readonly",
    # "https://www.googleapis.com/auth/gmail.metadata",
    "https://www.googleapis.com/auth/drive.metadata.readonly",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive.appdata",
]

# generate Path
parent_dirpath = Path(__file__).parents[1]
export_dirpath = parent_dirpath / "export_files"
export_dirpath.mkdir(exist_ok=True)

token_save_path = parent_dirpath / "token.json"
cred_json = parent_dirpath / cred_filepath


def decode_base64url(s) -> bytes:
    return base64.urlsafe_b64decode(s) + b"=" * (4 - (len(s) % 4))


def encode_base64url(bytes_data) -> str:
    return base64.urlsafe_b64encode(bytes_data).rstrip(b"=")


def convert_gmail_datetimestr(gmail_datetimeformat: str) -> datetime:
    return dateutil.parser.parse(gmail_datetimeformat)


def save_attachment_file(
    service, filename: str, message_id: str, attachment_id: str
) -> None:
    attachfile_data = (
        service.users()
        .messages()
        .attachments()
        .get(
            userId=target_userid,
            messageId=message_id,
            id=attachment_id,
        )
        .execute()
    )
    # print(attachfile_data)
    print((export_dirpath / Path(filename)))
    with (export_dirpath / Path(filename)).open("wb") as msg_img_file:
        msg_img_file.write(base64.urlsafe_b64decode(attachfile_data.get("data")))


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


def main() -> None:
    print("[Start Process...]")
    creds = None
    if token_save_path:
        creds = Credentials.from_authorized_user_file(token_save_path, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(cred_json, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with token_save_path.open("w") as token:
            token.write(creds.to_json())

    messages: list[ExpandedMessageItem] = []
    try:
        # Call the Gmail API
        service = build("gmail", "v1", credentials=creds)

        # スレッド検索
        thread_results = (
            service.users()
            .threads()
            .list(userId=target_userid, q="label:snd-ミスミ subject:(*MA-*)")
            .execute()
        )
        threads = thread_results.get("threads", [])

        # pprint(threads)

        # 上位10件のスレッド -> メッセージを取得。
        # スレッドに紐づきが2件ぐらいのメッセージの部分でのもので十分かな
        if threads:
            top_threads = list(itertools.islice(threads, 0, 10))
            for thread in top_threads:

                # threadsのid = threadsの一番最初のmessage>idなので、そのまま使う
                message_id = thread.get("id", "")
                # print(message_id)

                # スレッドの数が2以上 = すでに納品済みと思われるので削る。
                # TODO:2022-12-09 ここは2件以上でもまだやり取り中だったりする場合もあるので悩ましい
                # （数見るだけでもいいかもしれない
                thread_result = (
                    service.users()
                    .threads()
                    .get(userId=target_userid, id=message_id, fields="messages")
                ).execute()

                message_result = (
                    service.users().messages().get(userId=target_userid, id=message_id)
                ).execute()

                if len(thread_result.get("messages")) <= 2:
                    messages.append(ExpandedMessageItem(gmail_message=message_result))

    except HttpError as error:
        # TODO:2022-12-09 エラーハンドリングは基本行わずここで落とすこと
        print(f"An error occurred: {error}")
        exit()

    if not messages:
        print(f"cant find messages...")
        exit()

    # 取得が面倒なので最初から必要な値を取り出す

    message_item_and_labels = []
    for message in messages:
        # 送信日, タイトル
        choice_label = [
            ("class:text", f"{message.datetime_}"),
            ("class:highlighted", f"{message.title}"),
        ]
        message_item_and_labels.append(
            questionary.Choice(title=choice_label, value=message)
        )

    # 上位10のスレッドから > メッセージの最初取り出して、その中から選ぶ

    print("[Select Mail...]")

    selected_message: ExpandedMessageItem = questionary.select(
        "select msm gas mail message", choices=message_item_and_labels
    ).ask()

    # 選択後、処理開始していいか問い合わせして実行
    comefirm_check = questionary.confirm("run Process?", False).ask()

    if not comefirm_check:
        print("[Cancell Process...")
        exit()

    print("[Generate mail Printable PDF]")

    # メールのmimeマルチパートを考慮して、構造が違うモノに対応する
    # relative> altanative の手順で掘る。mixedは一番上で考慮しているのでここでは行わない
    message_body_related = {}
    message_body_parts = next(
        (
            i
            for i in selected_message.payload.get("parts")
            if i.get("mimeType") == "multipart/alternative"
        ),
        {},
    ).get("parts", [])
    if not message_body_parts:
        message_body_related = next(
            (
                i
                for i in selected_message.payload.get("parts")
                if i.get("mimeType") == "multipart/related"
            )
        )
        message_body_parts = next(
            (
                i
                for i in message_body_related.get("parts")
                if i.get("mimeType") == "multipart/alternative"
            )
        ).get("parts")

    messages_text_parts = [i for i in message_body_parts if "text" in i.get("mimeType")]

    b64dec_msg_byte = decode_base64url(messages_text_parts[1].get("body").get("data"))

    # imgタグを除去する
    mail_html_bs4 = BeautifulSoup(b64dec_msg_byte, "html.parser")
    only_body_tags = mail_html_bs4.body
    for t in only_body_tags.find_all("img"):
        t.decompose()

    # jinja2埋込
    # テンプレート読み込み
    env = Environment(
        loader=FileSystemLoader(str(Path(__file__).absolute().parent), encoding="utf8")
    )
    tmpl = env.get_template("export.html.jinja")

    # 設定ファイル読み込み
    params = {
        "export_html": only_body_tags,
        "message_title": html.escape(selected_message.title),
        "message_from_addresss": html.escape(selected_message.from_addresss),
        "message_cc_address": html.escape(selected_message.cc_address),
        "message_datetime": selected_message.datetime_,
    }
    # レンダリングして出力
    rendered_html = tmpl.render(params)
    # print(rendered_html)

    with (export_dirpath / Path("./export_mail.html")).open(
        "w", encoding="utf8"
    ) as exp_mail_hmtl:
        exp_mail_hmtl.write(rendered_html)

    print("[Save Attachment file and mail image]")
    # メール本文にimgファイルがある場合はそれを取り出す
    # multipart/relatedの時にあるので、それを狙い撃ちで取る

    if message_body_related:
        message_imgs = [
            i for i in message_body_related.get("parts") if "image" in i.get("mimeType")
        ]
        # print(message_imgs)
        for msg_img in message_imgs:
            # print(msg_img.get("filename"))
            save_attachment_file(
                service,
                msg_img.get("filename"),
                selected_message.id,
                msg_img.get("body").get("attachmentId"),
            )
    # 添付ファイルの保持
    message_attachmentfiles = [
        i
        for i in selected_message.payload.get("parts")
        if "application" in i.get("mimeType")
    ]

    # print(message_attachmentfiles)

    for msg_attach in message_attachmentfiles:
        save_attachment_file(
            service,
            msg_attach.get("filename"),
            selected_message.id,
            msg_attach.get("body").get("attachmentId"),
        )

    print("[Generate Excel Printable PDF]")
    # 連絡項目の印刷用PDFファイル生成
    renrakukoumoku_excel_attachmenfile = next(
        (
            i
            for i in selected_message.payload.get("parts")
            if mimetype_gsheet in i.get("mimeType")
        )
    )

    # print(renrakukoumoku_excel_attachmenfile)

    # exit()
    if renrakukoumoku_excel_attachmenfile:
        # ExcelファイルをPDFに変換する
        target_filepath = export_dirpath / renrakukoumoku_excel_attachmenfile.get(
            "filename"
        )
        upload_mimetype = mimetype_gsheet

        media = MediaFileUpload(
            target_filepath,
            mimetype=upload_mimetype,
            resumable=True,
        )

        file_metadata = {
            "name": target_filepath.name,
            "mimeType": "application/vnd.google-apps.spreadsheet",
            "parents": ["1f0efE1nKIodvUBQ_rB5GjZlyqwDlglI_"],
        }

        try:
            service = build("drive", "v3", credentials=creds)

            # Excelファイルのアップロードを行って、そのアップロードファイルをPDFで保存できるかチェック
            upload_results = (
                service.files()
                .create(body=file_metadata, media_body=media, fields="id")
                .execute()
            )

            print(upload_results.get("id"))

            # pdfファイルを取りに行ってみる
            dl_request = service.files().export_media(
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

            # できたら最後にファイルを除去する
            delete_tmp_excel_result = (
                service.files()
                .delete(fileId=upload_results.get("id"), fields="id")
                .execute()
            )
            print("delete file: ", delete_tmp_excel_result)

        except HttpError as error:
            # TODO:2022-12-09 エラーハンドリングは基本行わずここで落とすこと
            print(f"An error occurred: {error}")

        # ボイラープレートからディレクトリ生成

        print("[Generate template dirs]")
        # detault設定: 正規表現で取れなかった場合はとりあえず作ってしまう
        boilerplate_config = {"project_name": "0000", "haikan_pattern": "type_s"}

        boilerplate_config["project_name"] = re.match(
            r"MA-(\d{1,4}|\d{1,4}-\d{1})_",
            renrakukoumoku_excel_attachmenfile.get("filename"),
        ).group(1)
        copier.run_copy(
            msm_gas_boilerplate_url,
            parent_dirpath / "export_files",
            data=boilerplate_config,
        )

    exit()


if __name__ == "__main__":
    main()
