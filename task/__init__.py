# 各タスクの基底クラス
from typing import Protocol, TypedDict


# TODO:2023-10-09　やり取りに使うデータ構造の定義はtypeddictではなくてdataclassを使う
class ProcessData(TypedDict):
    task_data: dict


class MessageData(TypedDict):
    """
    google chatにポストするメッセージのデータ構造
    """

    # TODO:2023-09-29 データ構造をつくって、型指定する
    # 必要なもの: メッセージ本文 or 設定カードのjson

    text: str
    # name: str
    # icon_url: str
    # link: str


class DialogMessage(TypedDict):
    """
    google chatの設定カードメッセージのデータ構造
    """

    # TODO:2023-09-29 google chatのドキュメントを見て、必要なものを追加する

    text: str
    dialog_json: dict


class BaseTask(Protocol):
    def execute_task(self, process_data: ProcessData | None = None) -> dict | str:
        ...

    # google chatにタスク結果をポストする
    # TODO:2023-09-28 データ構造をつくって、型指定する
    def send_message(self, message_data: MessageData) -> str:
        # TODO: ここでgoogle chatにポストする処理を書く

        print(message_data)
        ...
