from __future__ import print_function

import base64
import os
import os.path
from datetime import datetime
from pathlib import Path

import dateutil.parser
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from jinja2 import Environment, FileSystemLoader, Template

import html

load_dotenv()


cred_filepath = os.environ.get("CRED_FILEPATH")
target_userid = os.environ.get("GMAIL_USER_ID")

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

cred_json = Path(__file__).parents[1] / cred_filepath


def decode_base64url(s):
    return base64.urlsafe_b64decode(s) + b"=" * (4 - (len(s) % 4))


def encode_base64url(bytes_data):
    return base64.urlsafe_b64encode(bytes_data).rstrip(b"=")


def convert_gmail_datetimestr(gmail_datetimeformat: str) -> datetime:
    return dateutil.parser.parse(gmail_datetimeformat)


def main():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(cred_json, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

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
        # 一番上のものでいいから表示
        threads = thread_results.get("threads", [])
        from pprint import pprint

        pprint(threads)

        # 上位10件のスレッド -> メッセージを取得。見てもスレッド数が2件ぐらいのもので十分かな
        # 上位10のスレッドから > メッセージの最初取り出して、その中から選ぶ
        # 選択後のメッセージを元に処理開始

        top_message_id = threads[2].get("id", "")
        print(top_message_id)

        message_result = (
            service.users().messages().get(userId=target_userid, id=top_message_id)
        ).execute()

        message_payload = message_result.get("payload")
        message_headers = message_payload.get("headers")

    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f"An error occurred: {error}")

    # print("message_payload...")
    # print(message_payload)
    # print(message_headers)

    message_title = html.escape(
        next((i for i in message_headers if i.get("name") == "Subject")).get("value")
    )

    message_from_addresss = html.escape(
        next((i for i in message_headers if i.get("name") == "From")).get("value")
    )

    message_cc_address = html.escape(
        next((i for i in message_headers if i.get("name") == "CC")).get("value")
    )
    message_datetime_str = next(
        (i for i in message_headers if i.get("name") == "Date")
    ).get("value")

    message_datetime = convert_gmail_datetimestr(message_datetime_str)

    print(
        f"title:{message_title}\nfrom:{message_from_addresss}\ncc:{message_cc_address}\n{message_datetime}"
    )

    # メールのmimeマルチパートを考慮して、構造が違うモノに対応する
    # relative> altanative の手順で掘る。mixedは一番上で考慮しているのでここでは行わない
    message_body_related = {}
    message_body_parts = next(
        (
            i
            for i in message_payload.get("parts")
            if i.get("mimeType") == "multipart/alternative"
        ),
        {},
    ).get("parts", [])
    if not message_body_parts:
        message_body_related = next(
            (
                i
                for i in message_payload.get("parts")
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
        "message_title": message_title,
        "message_from_addresss": message_from_addresss,
        "message_cc_address": message_cc_address,
        "message_datetime": message_datetime,
    }
    # レンダリングして出力
    rendered_html = tmpl.render(params)
    # print(rendered_html)

    with Path("./export_mail.html").open("w", encoding="utf8") as exp_mail_hmtl:
        exp_mail_hmtl.write(rendered_html)

    # メール本文にimgファイルがある場合はそれを取り出す
    # multipart/relatedの時にあるので、それを狙い撃ちで取る

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
        with Path(filename).open("wb") as msg_img_file:
            msg_img_file.write(decode_base64url(attachfile_data.get("data")))

    if message_body_related:
        message_imgs = [
            i for i in message_body_related.get("parts") if "image" in i.get("mimeType")
        ]
        print(message_imgs)
        for msg_img in message_imgs:
            # print(msg_img.get("filename"))
            save_attachment_file(
                service,
                msg_img.get("filename"),
                top_message_id,
                msg_img.get("body").get("attachmentId"),
            )
    # 添付ファイルの保持

    message_attachmentfiles = [
        i for i in message_payload.get("parts") if "application" in i.get("mimeType")
    ]
    print(message_attachmentfiles)
    for msg_attach in message_attachmentfiles:
        save_attachment_file(
            service,
            msg_attach.get("filename"),
            top_message_id,
            msg_attach.get("body").get("attachmentId"),
        )
    exit()


if __name__ == "__main__":
    main()
