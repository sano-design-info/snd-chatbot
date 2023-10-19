from dataclasses import dataclass, asdict


# アクションレスポンス生成
def genactionresponse_dialog(action_status: str = "OK") -> dict:
    """
    Google Chatのアクションレスポンスを生成する。結果はjson構造として辞書型で返す。
    OKの場合はOKのみを入れる。デフォルト引数に指定済み。
    エラーの場合はエラーメッセージとして文字列を入れる。
    args:
        action_status: OK or error message by string

    return: アクションレスポンス
    """
    return {
        "actionResponse": {
            "type": "DIALOG",
            "dialogAction": {"actionStatus": {"statusCode": action_status}},
        }
    }


def genheader(title: str, subtitle: str, imageurl: str) -> dict:
    return {
        "title": title,
        "subtitle": subtitle,
        "imageUrl": imageurl,
        "imageType": "CIRCLE",
    }


def genwidget_textinput_singleline(label: str, name: str) -> dict:
    return {
        "textInput": {
            "label": label,
            "type": "SINGLE_LINE",
            "name": name,
        }
    }


def genwidget_textparagraph(text: str) -> dict:
    return {"textParagraph": {"text": text}}


# ボタンリストのボタン生成用
def gencomponent_button(
    text: str, function_name: str, parameters: list = None, alttext: str = ""
) -> dict:
    if parameters is None:
        parameters = []
    return {
        "text": text,
        "onClick": {
            "action": {
                "function": function_name,
                "parameters": parameters,
            }
        },
        "altText": alttext,
        "disabled": False,
    }


# ボタンリストの全体生成用
def genwidget_buttonlist(buttons: list[dict]) -> dict:
    return {"buttonList": {"buttons": buttons}}


@dataclass
class SelectionInputItem:
    """
    選択肢の型
    """

    text: str
    value: str
    selected: bool = False


def genwidget_radiobuttonlist(
    label: str, name: str, items: list[SelectionInputItem]
) -> dict:
    """
    単一選択式:ラジオボタンのウィジェットを生成する
    args:
        label: ラベル
        name: ラジオボタンのname
        items: ラジオボタンの選択肢
    return: ウィジェット
    """
    return {
        "selectionInput": {
            "type": "RADIO_BUTTON",
            "label": label,
            "name": name,
            "items": [asdict(item) for item in items],
        }
    }


def genwidget_checkboxlist(
    label: str, name: str, items: list[SelectionInputItem]
) -> dict:
    """
    複数選択式:チェックボックスのウィジェットを生成する
    args:
        label: ラベル
        name: チェックボックスのname
        items: チェックボックスの選択肢
    return: ウィジェット
    """

    return {
        "selectionInput": {
            "type": "CHECK_BOX",
            "label": label,
            "name": name,
            "items": [asdict(item) for item in items],
        }
    }


def genwidget_switchlist(
    label: str, name: str, items: list[SelectionInputItem]
) -> dict:
    """
    スイッチ方式: スイッチのウィジェットを生成する
    args:
        label: ラベル
        name: スイッチのname
        items: スイッチの選択肢
    return: ウィジェット
    """

    return {
        "selectionInput": {
            "type": "SWITCH",
            "label": label,
            "name": name,
            "items": [asdict(item) for item in items],
        }
    }


def create_card(cardid: str, header: dict, widgets: list[dict] | None = None) -> dict:
    """Generate a card with a header and single widget.

    Args:
        message_text: The text of the message to be placed in the card body.

    """
    return {
        "cardsV2": [
            {
                "cardId": cardid,
                "card": {
                    "header": header,
                    "sections": [{"widgets": widgets}],
                },
            }
        ]
    }


# よくあるカードの生成
# 応答確認
# 総合機能のメニュー表示


# カードを使った短文の送信
def create_card_text(cardid: str, header: dict, text: str) -> dict:
    textparagraph = genwidget_textparagraph(text)

    return create_card(cardid, header, [textparagraph])
