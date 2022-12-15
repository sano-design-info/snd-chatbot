import base64
import itertools
import os
from pathlib import Path

import questionary
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from helper import google_api_helper
from prepare import (
    ExpandedMessageItem,
    add_schedule_spreadsheet,
    generate_estimate_calcsheet,
    generate_mail_printhtml,
    generate_pdf_by_renrakukoumoku_excel,
    generate_projectdir,
)

# load config
load_dotenv()

target_userid = os.environ.get("GMAIL_USER_ID")

# generate Path
parent_dirpath = Path(__file__).parents[0]
export_dirpath = parent_dirpath / "export_files"
export_dirpath.mkdir(exist_ok=True)

attachment_files_dirpath = export_dirpath / "attachments"
attachment_files_dirpath.mkdir(exist_ok=True)

# If modifying these scopes, delete the file token.json.
GOOGLE_API_SCOPES = [
    "https://mail.google.com/",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.readonly",
    # "https://www.googleapis.com/auth/gmail.metadata",
    "https://www.googleapis.com/auth/drive.metadata.readonly",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive.appdata",
]


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
    print((attachment_files_dirpath / Path(filename)))
    with (attachment_files_dirpath / Path(filename)).open("wb") as attachmentfiile:
        attachmentfiile.write(base64.urlsafe_b64decode(attachfile_data.get("data")))


def main() -> None:
    print("[Start Process...]")

    google_creds: Credentials = google_api_helper.get_cledential(GOOGLE_API_SCOPES)

    messages: list[ExpandedMessageItem] = []
    try:
        # Call the Gmail API
        service = build("gmail", "v1", credentials=google_creds)

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
        print("[Cancell Process...]")
        exit()

    print("[Save Attachment file and mail image]")
    # メール本文にimgファイルがある場合はそれを取り出す
    # multipart/relatedの時にあるので、それを狙い撃ちで取る

    if selected_message.body_related:
        message_imgs = [
            i
            for i in selected_message.body_related.get("parts")
            if "image" in i.get("mimeType")
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

    # 各種機能を呼び出す

    print("[Generate Mail Printable PDF]")
    generate_mail_printhtml(selected_message, export_dirpath)

    print("[Generate Excel Printable PDF]")
    generate_pdf_by_renrakukoumoku_excel(
        attachment_files_dirpath, export_dirpath, google_creds
    )

    print("[Generate template dirs]")
    generate_projectdir(attachment_files_dirpath, export_dirpath)

    print("[append schedule]")
    add_schedule_spreadsheet(attachment_files_dirpath, google_creds)

    print("[add estimate calcsheet]")
    generate_estimate_calcsheet(attachment_files_dirpath, google_creds)

    print("[End Process...]")
    exit()


if __name__ == "__main__":
    main()
