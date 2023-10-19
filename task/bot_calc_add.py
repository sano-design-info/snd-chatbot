from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

from api import googleapi
import chat.card
from helper import load_config, chatcard

# 認証情報をtomlファイルから読み込む
config = load_config.CONFIG


google_cred: Credentials = googleapi.get_cledential_by_serviceaccount(
    googleapi.CHAT_API_SCOPES
)
chat_service = build("chat", "v1", credentials=google_cred)

spacename = config.get("google").get("CHAT_SPACENAME")
bot_header = chatcard.bot_header


def bot_calc_add(a: int, b: int) -> dict:
    # チャットで非同期に結果を返す

    print("taskstart")

    import time

    time.sleep(2)

    print(f"bot_calc_add: {a} + {b} = {a+b}")

    send_message_body = chat.card.create_card(
        "result_card__calc_add",
        header=bot_header,
        widgets=[
            chat.card.genwidget_textparagraph(
                f"計算結果: {a} + {b} = {a+b} ",
            ),
        ],
    )
    return googleapi.create_chat_message(chat_service, spacename, send_message_body)
