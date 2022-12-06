from __future__ import print_function

import os.path
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from dotenv import load_dotenv

import base64
import quopri
import os

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
        # 並びは最新？
        # その中からメッセージを取り出してメッセージの情報を収集
        # 一番上のものでいいから表示
        threads = thread_results.get("threads", [])

        # print(threads)

        top_message_id = threads[0].get("id", "")
        message_result = (
            service.users().messages().get(userId=target_userid, id=top_message_id)
        ).execute()

        message_payload = message_result.get("payload", {})
        message_headers = message_payload.get("headers")

        # print("message_payload...")
        # print(message_payload)

        message_title = next(
            (i for i in message_headers if i.get("name") == "Subject")
        ).get("value")
        message_from_addresss = next(
            (i for i in message_headers if i.get("name") == "From")
        ).get("value")
        message_cc_addresss = next(
            (i for i in message_headers if i.get("name") == "CC")
        ).get("value")

        print(
            f"title:{message_title}\nfrom:{message_from_addresss}\ncc:{message_cc_addresss}"
        )

        message_body_parts = next(
            (
                i
                for i in message_payload.get("parts")
                if i.get("mimeType") == "multipart/alternative"
            )
        ).get("parts")

        messages_text_parts = [
            i for i in message_body_parts if "text" in i.get("mimeType")
        ]

        b64dec_msg_byte = decode_base64url(
            messages_text_parts[1].get("body").get("data")
        )

        print("decode msg")
        print(b64dec_msg_byte)

        with open("./export.html", "wb") as exp_html:
            exp_html.write(b64dec_msg_byte)

        # print(
        #     [
        #         base64.b64decode(i.get("body").get("data")).decode("")
        #         for i in messages_text_parts
        #     ]
        # )

        # message_attachedfiles = [
        #     i for i in message_body_parts if "application" in i.get("mimeType")
        # ]
        # print(message_attachedfiles)

        exit()

    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f"An error occurred: {error}")


if __name__ == "__main__":
    main()
