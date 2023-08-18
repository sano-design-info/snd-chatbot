import base64
import io
import mimetypes
from email.message import EmailMessage
from pathlib import Path

import requests
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import Resource, build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

# load config
from helper import EXPORTDIR_PATH, load_config
from itemparser import ExpandedMessageItem

config = load_config.CONFIG
cred_filepath = config.get("google").get("CRED_FILEPATH")

# generate Path
EXPORTDIR_PATH.mkdir(parents=True, exist_ok=True)
token_save_path = EXPORTDIR_PATH / "token.json"
cred_json = EXPORTDIR_PATH / cred_filepath

API_SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://mail.google.com/",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.appdata",
    "https://www.googleapis.com/auth/drive.metadata",
]


def get_cledential(scopes: list[str]) -> Credentials:
    """
    Google APIの認証情報を取得します。
    既に認証情報があればそれを返します。
    なければ認証情報を取得し、token.jsonに保存します。
    args:
        scopes: 認証情報を取得する際に必要なスコープ
    return:
        認証情報
    """
    creds = None
    if token_save_path.exists():
        creds = Credentials.from_authorized_user_file(token_save_path, scopes)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(cred_json, scopes)
            creds = flow.run_local_server(bind_addr="0.0.0.0", port=18081)
        # Save the credentials for the next run
        with token_save_path.open("w") as token:
            token.write(creds.to_json())
    return creds


# [gmail api]
def search_threads(gmail_service: Resource, query) -> list[dict]:
    """
    Google Gmail APIを使用して、スレッドを検索します。
    args:
        service: Gmail APIのサービス
        query: 検索クエリ
    return:
        スレッドのリスト
    """
    try:
        response = (
            gmail_service.users()
            .threads()
            .list(userId="me", q=query, includeSpamTrash=False)
            .execute()
        )
        threads = []
        if "threads" in response:
            threads.extend(response["threads"])

        while "nextPageToken" in response:
            page_token = response["nextPageToken"]
            response = (
                gmail_service.users()
                .threads()
                .list(
                    userId="me",
                    q=query,
                    pageToken=page_token,
                    includeSpamTrash=False,
                )
                .execute()
            )
            threads.extend(response["threads"])

        return threads
    except HttpError as error:
        print(f"An error occurred: {error}")
        return []


def get_messages_by_threadid(gmail_service: Resource, thread_id) -> list[dict]:
    """
    Gmail APIを使用して、スレッドIDからメッセージを取得します。
    args:
        service: Gmail APIのサービス
        thread_id: スレッドID
    return:
        メッセージのリスト
    """
    try:
        response = (
            gmail_service.users()
            .threads()
            .get(
                userId="me",
                id=thread_id,
            )
            .execute()
        )
        messages = []
        if "messages" in response:
            messages.extend(response["messages"])

        while "nextPageToken" in response:
            page_token = response["nextPageToken"]
            response = (
                gmail_service.users()
                .threads()
                .get(
                    userId="me",
                    id=thread_id,
                    pageToken=page_token,
                )
                .execute()
            )
            messages.extend(response["messages"])

        return messages
    except HttpError as error:
        print(f"An error occurred: {error}")
        return []


def create_messagedata(
    sender: str,
    to: str,
    cc: str,
    reply_to: str,
    subject: str,
    message_text: str,
    attachment_files: list[Path] | None = None,
) -> dict:
    """
    Gmail APIで利用するメッセージを作成します。
    args:
        sender: 送信者
        to: 宛先
        subject: 件名
        message_text: 本文 text/plain想定で作ってます
        attachment_files: 添付ファイルのリスト Path型 Noneの場合は添付ファイルなし
    return:
        メッセージのdict
    """
    message = EmailMessage()
    message["to"] = to
    message["from"] = sender
    message["cc"] = cc
    message["reply-to"] = reply_to
    message["subject"] = subject
    message["references"] = reply_to

    # ここではtext/plainのみ対応
    message.set_content(message_text)

    # 添付ファイルの追加
    if attachment_files:
        for attachment_file in attachment_files:
            # guessing the MIME type
            type_subtype, _ = mimetypes.guess_type(attachment_file)
            maintype, subtype = type_subtype.split("/")
            with attachment_file.open("rb") as f:
                file_data = f.read()
            message.add_attachment(
                file_data,
                maintype=maintype,
                subtype=subtype,
                filename=attachment_file.name,
            )

    return base64.urlsafe_b64encode(message.as_bytes()).decode()


# 返信用のメッセージ生成
def create_reply_gmail_messagedata(
    gmail_service: Resource,
    body: str,
    message_id: str,
    attachment_files: list[Path] | None = None,
    thread_id: str | None = None,
) -> dict:
    """
    Gmail APIを使用して、返信用のメッセージを生成します。メッセージはテキストモードで生成されます
    返信先のメッセージを引用符で囲み、本文を追加します。
    args:
        service: Gmail APIのサービス
        body: メッセージ本文
        attachment_files: 添付ファイルのリスト
        message_id: メッセージID
        thread_id: スレッドID
    return:
        返信メッセージ用のデータ。Gmail API向けrawメッセージ
    """
    # 元のメッセージからヘッダー情報を取得
    org_message = (
        gmail_service.users().messages().get(userId="me", id=message_id).execute()
    )
    org_messageitem = ExpandedMessageItem(org_message)

    # 元メッセージの日付と差出人情報を載せる
    reply_info = f"--- {org_messageitem.datetime_:%Y/%m/%d %H:%M:%S} {org_messageitem.from_address} wrote ---"

    # 元のメッセージのボディを引用符で囲む
    replace_body = org_messageitem.body.replace("\\", "\\\\").replace("\n", "\n> ")
    quoted_body = f"> {replace_body}"

    reply_body = f"{body}\n{reply_info}\n{quoted_body}"

    # 返信メッセージを作成
    mime_message = create_messagedata(
        org_messageitem.to_address,
        org_messageitem.from_address,
        org_messageitem.cc_address,
        message_id,
        org_messageitem.subject,
        reply_body,
        attachment_files,
    )
    # 返信用にスレッドIDを指定
    if thread_id:
        return {"raw": mime_message, "threadId": thread_id}
    else:
        return {"raw": mime_message}


# 新規のGmail APIの rawメッセージを作成
def create_blank_gmail_messagedata(
    gmail_service: Resource,
    to: str,
    cc: str,
    subject: str,
    body: str,
    attachment_files: list[Path] | None = None,
) -> dict:
    """
    Gmail APIを使用して、新規のGmail APIの rawメッセージを作成します。
    args:
        service: Gmail APIのサービス
        to: 宛先
        cc: CC
        subject: 件名
        body: 本文
        attachment_files: 添付ファイルのリスト
    return:
        新規メッセージ用のデータ。Gmail API向けrawメッセージ
    """
    try:
        # 送信者のメールアドレスを取得
        sender = gmail_service.users().getProfile(userId="me").execute()["emailAddress"]
    except HttpError as error:
        print(f"An error occurred: {error}")
        return {}

    mime_message = create_messagedata(
        sender,
        to,
        cc,
        sender,
        subject,
        body,
        attachment_files,
    )
    return {"raw": mime_message}


# Gmail APIで下書きメールを作成してスレッドにつける
def append_draft_in_thread(
    gmail_service: Resource,
    body: str,
    attachment_files: list[Path],
    message_id: str,
    thread_id: str | None = None,
) -> dict | None:
    """
    Gmail APIを使用して、下書きメールを作成してスレッドにつけます。
    args:
        service: Gmail APIのサービス
        body: メッセージ本文
        message_id: メッセージID
        thread_id: スレッドID
    return:
        作成した下書きメールの情報
    """
    try:
        draft = (
            gmail_service.users()
            .drafts()
            .create(
                userId="me",
                body={
                    "message": create_reply_gmail_messagedata(
                        gmail_service, body, message_id, attachment_files, thread_id
                    )
                },
            )
            .execute()
        )

        print(f'Saved draft: Draft Id: {draft["id"]}')

    except HttpError as error:
        print(f"An error occurred: {error}")
        draft = None

    return draft


# メールの下書き新規作成用
def append_draft(
    gmail_service: Resource,
    to: str,
    cc: str,
    title: str,
    body: str,
    attachment_files: list[Path],
) -> dict | None:
    """
    Gmail APIを使用して、下書きメールを作成します。
    args:
        service: Gmail APIのサービス
        to: 宛先
        cc: CC
        title: 件名
        body: 本文
        attachment_files: 添付ファイルのリスト
    return:
        作成した下書きメールの情報
    """
    try:
        draft = (
            gmail_service.users()
            .drafts()
            .create(
                userId="me",
                body={
                    "message": create_blank_gmail_messagedata(
                        gmail_service,
                        to,
                        cc,
                        title,
                        body,
                        attachment_files,
                    )
                },
            )
            .execute()
        )
        print(f'Saved draft: Draft Id: {draft["id"]}')

    except HttpError as error:
        print(f"An error occurred: {error}")
        draft = None

    return draft


#


# Google Drive API
def dupulicate_file(
    drive_service: Resource, file_id: str, file_name: str | None = None
) -> str:
    """
    Google Drive APIを使用して、シートを複製します。
    args:
        drive_service: Drive APIのサービス
        file_id: ファイルID
        file_name: 複製したファイルの名前。Noneの場合は複製元のファイル名をそのまま使う
    return:
        複製したファイルのID
    """

    # テンプレートを複製します
    dupulicated_file_id = ""
    try:
        response = drive_service.files().copy(fileId=file_id).execute()

        # 複製したファイルのIDを取得
        dupulicated_file_id = response["id"]
    except HttpError as error:
        print(f"An error occurred: {error}")
        return ""

    # file_nameがある場合は、ファイル名を変更する
    if file_name:
        try:
            _ = (
                drive_service.files()
                .update(
                    fileId=dupulicated_file_id,
                    body={"name": file_name},
                )
                .execute()
            )
        except HttpError as error:
            print(f"An error occurred: {error}")
            return ""

    # 複製したファイルのIDを返す
    return dupulicated_file_id


##


def download_pdf_as_file(
    drive_service: Resource, file_id: str, pdffilepath: Path
) -> None:
    """
    Google Drive APIを使用して、file_idからPDFファイルをダウンロードします。
    ダウンロードに失敗した場合は例外を出します。
    args:
        drive_service: Google Drive APIのサービス
        file_id: ファイルID
        pdffilepath: PDFファイルの保存先
    return:
        なし
    """

    try:
        request = drive_service.files().export_media(
            fileId=file_id, mimeType="application/pdf"
        )
        fh = io.FileIO(pdffilepath, "wb")
        downloader = MediaIoBaseDownload(fh, request)

        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print("ダウンロード進捗: %d%%" % int(status.progress() * 100))
        fh.close()
    except Exception as e:
        print(f"PDFファイルのダウンロードに失敗しました: {e}")
        raise


# GoogleドライブのエクスポートURLを元にしたファイルダウンロード
def export_pdf_by_driveexporturl(token: str, file_id: str, save_path: Path) -> None:
    """
    GoogleドライブのエクスポートURLを元に、PDFファイルをダウンロードします。
    googleapiclientモジュールは利用せず、requestsモジュールを使用しています。そのため、google_auth_oauthlibモジュールで取得したトークンをtoken引数に渡してください。

    ダウンロードに失敗した場合は、requests.exceptions.HTTPError例外を出します。

    現時点で、特定のワークロードでのみ動作確認しています。エクスポートURLのクエリパラメータは以下の通りです。

    * format=pdf
    * portrait=false : 横向きにする

    TODO:2023-06-05 クエリパラメータの指定を引数で行えるようにする

    args:
        token: Google APIのトークン
        file_id: ファイルID
        save_path: PDFファイルの保存先
    return:
        なし

    """
    # エクスポートするURLを生成
    export_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=pdf&portrait=false"

    # requestsでダウンロードする。stream指定でチャンクサイズは1MBでダウンロードする
    params = {"access_token": token}

    with requests.get(export_url, params=params, stream=True) as r:
        r.raise_for_status()
        with save_path.open("wb") as f:
            for chunk in r.iter_content(chunk_size=1024 * 1024):
                f.write(chunk)


# TODO: 2023-03-30 以降にGoogle APIを操作するヘルパー関数を引越してまとめておく
