import json
import logging
import os
from pprint import pprint
from dataclasses import dataclass

import redis
from flask import Flask, render_template, request
from flask import json as f_json
from rq import Queue

import chat.card
import chat.session
from helper import chatcard, convert_dataclass_to_jsonhash_str
from task import (
    bot_calc_add,
    generate_quotes,
    get_billing,
    run_mail_action,
    # send_sagakusyoumei,
)

# Flaskアプリケーションの初期化
app = Flask(__name__)

# RQのキューを初期化
queue = Queue(connection=redis.from_url(os.environ.get("RQ_REDIS_URL")))

# チャットボットの設定
bot_header = chatcard.bot_header

# アクションレスポンスの定義
# https://developers.google.com/hangouts/chat/reference/message-formats/cards#action_response
ACTION_RESPONSE_OK_JSON = chat.card.genactionresponse_dialog()


# TODO:2023-10-19 cardIdとinvokedFunctionの名前についてメモを残す。
# いずれは
# [基本ルール]
# 最後に、機能の名称は盛り込む [なにか]__[関係する各タスクの名称]

# ##invokedFunction名
# * run_prepare: PrepareTaskを実行する
# * comfirm: 確認用カードを出す
# * run_task: MainTaskを実行する
# * cancell: キャンセル処理を行う

# ##カード名
# * config_card: タスク実行のための設定を行うカード
# * comfirm_card: タスク実行の確認用カード
# * result_card: タスク実行の結果で伝えるときに使う
# * nortify_task_card: 実行中などに起こる色々な状態を伝えるときに使う


# [generate_quotes]: 設定カード→メインタスク実行
def run_preparetask_generate_quotes() -> dict:
    # prepareタスク実行後に設定カードを開く
    script_task = generate_quotes.PrepareTask()
    job = queue.enqueue(script_task.execute_task_by_chat)

    print(f"job.id:{job.id}")
    return ACTION_RESPONSE_OK_JSON


def run_task_generate_quotes(session_data: dict) -> dict:
    task_data = {
        "task_data": {
            # selected_estimate_calcsheetsリスト中身はjson文字列なのでloadする
            "selected_estimate_calcsheets": [
                json.loads(estimate_calcsheet)
                for estimate_calcsheet in session_data.get(
                    "selected_estimate_calcsheets"
                )
            ]
        }
    }
    script_task = generate_quotes.MainTask()
    job = queue.enqueue(script_task.execute_task_by_chat, args=(task_data,))

    print(f"job.id:{job.id}")
    return ACTION_RESPONSE_OK_JSON


# [get_billing]: 設定カード→内容確認→メインタスク実行
def run_preparetask_get_billing() -> dict:
    # prepareタスク実行後に設定カードを開く
    prepare_task = get_billing.PrepareTask()
    job = queue.enqueue(prepare_task.execute_task_by_chat)

    print(f"job.id:{job.id}")
    return ACTION_RESPONSE_OK_JSON


def confirm_get_billing(session_data: dict) -> dict:
    # 請求書の計算を行う: セッションデータにはjson文字列が入っているので、dataclassへ変換する
    ask_choiced_quote_list = [
        convert_dataclass_to_jsonhash_str(choiced_quote_jsonstr, get_billing.QuoteData)
        for choiced_quote_jsonstr in session_data.get("choiced_quote_list")
    ]

    # 見積書の情報を元に、金額の合計を出す
    billing_data = get_billing.generate_invoice_data(ask_choiced_quote_list)

    # ここで請求書情報を出して、こちらの検証と正しいか確認
    calc_result = f"""
    [請求情報]
    件数: {len(ask_choiced_quote_list)}
    合計金額:{billing_data.price}
    """

    cardbody = chat.card.create_card(
        "confirm__get_billing",
        header=bot_header,
        widgets=[
            chat.card.genwidget_textparagraph(
                "請求書作成を実行しますか？",
            ),
            chat.card.genwidget_textparagraph(calc_result),
            chat.card.genwidget_buttonlist(
                [
                    chat.card.gencomponent_button("実行", "run_task__get_billing"),
                    chat.card.gencomponent_button("キャンセル", "cancell_task"),
                ]
            ),
        ],
    )

    return cardbody


def runtask_get_billing(session_data: dict) -> dict:
    # メインタスク向けに、セッションデータを整形する
    task_data = {
        "task_data": {
            # ask_dataの中身はjson文字列のリストなので、dataclassに変換する
            "choiced_quote_list": [
                convert_dataclass_to_jsonhash_str(
                    choiced_quote_jsonstr, get_billing.QuoteData
                )
                for choiced_quote_jsonstr in session_data.get("choiced_quote_list")
            ]
        }
    }
    script_task = get_billing.MainTask()
    job = queue.enqueue(script_task.execute_task_by_chat, args=(task_data,))

    print(f"job.id:{job.id}")
    return ACTION_RESPONSE_OK_JSON


# [run_mail_action]: 設定カード→メインタスク実行
def run_preparetask_run_mail_action() -> dict:
    # prepareタスク実行後に、設定カードを開く。設定カードはREST API経由で開く
    script_task = run_mail_action.PrepareTask()
    job = queue.enqueue(script_task.execute_task_by_chat)

    print(f"job.id:{job.id}")
    return ACTION_RESPONSE_OK_JSON


def runtask_run_mail_action(session_data: dict) -> dict:
    task_data = {
        "task_data": {
            "selected_message_id": session_data.get("selected_message_id"),
            "ask_generate_projectfile": session_data.get("ask_generate_projectfile"),
            "ask_add_schedule_and_generate_estimate_calcsheet": session_data.get(
                "ask_add_schedule_and_generate_estimate_calcsheet"
            ),
            "ask_add_schedule_nextmonth": session_data.get(
                "ask_add_schedule_nextmonth"
            ),
        }
    }
    script_task = run_mail_action.MainTask()
    job = queue.enqueue(script_task.execute_task_by_chat, args=(task_data,))

    print(f"job.id:{job.id}")
    return ACTION_RESPONSE_OK_JSON


# # [send_sagakusyoumei]: タスク実行のみ
# def runtask_send_sagakusyoumei() -> dict:
#     script_task = send_sagakusyoumei.ScriptTask()
#     job = queue.enqueue(script_task.execute_task_by_chat)

#     print(f"job.id:{job.id}")
#     # タスク実行がされたことを示すカードメッセージを送付する
#     return chat.card.create_card_text(
#         "nortify_task_card__send_sagakusyoumei",
#         bot_header,
#         "send_sagakusyoumei のタスクを実行します...",
#     )


# [calc_add]: 設定カード→内容確認カード→メインタスク実行
def open_cofig_card_calc_add() -> dict:
    first_num = chat.card.genwidget_textinput_singleline("first_num", "first_num")
    second_num = chat.card.genwidget_textinput_singleline("second_num", "second_num")
    buttonlist = chat.card.genwidget_buttonlist(
        [chat.card.gencomponent_button("確認", "confirm__calc_add")]
    )
    config_body = chat.card.create_card(
        "config_card__calc_add", bot_header, [first_num, second_num, buttonlist]
    )
    return config_body


def confirm_calc_add(session_data) -> dict:
    first_num = int(session_data["first_num"])
    second_num = int(session_data["second_num"])

    cardbody = chat.card.create_card(
        "confirm_card__run_calc",
        header=bot_header,
        widgets=[
            chat.card.genwidget_textparagraph(
                f"計算を実行しますか？: {first_num} + {second_num} ",
            ),
            chat.card.genwidget_buttonlist(
                [
                    chat.card.gencomponent_button("実行", "run_task__calc_add"),
                    chat.card.gencomponent_button("キャンセル", "cancell_task"),
                ]
            ),
        ],
    )
    return cardbody


def runtask_calc_add(session_data):
    job = queue.enqueue(
        bot_calc_add.bot_calc_add,
        args=(int(session_data["first_num"]), int(session_data["second_num"])),
    )

    print(f"job.id:{job.id}")
    return ACTION_RESPONSE_OK_JSON


# メインタスク実行
def response_generator(event):
    """Determine what response to provide based upon event data.

    Args:
        event: A dictionary with the event data.

    """

    text = ""

    # event shotcut values
    event_type = event["type"]
    event_common = event["common"]
    # dialog_event_type = event.get("dialogEventType", None)
    invoked_function = event_common.get("invokedFunction")
    slash_command = event["message"].get("slashCommand", dict())

    # TODO:2023-10-18 debugはlogging化とする
    # debug
    print(f"type:{event['type']}")
    print(f"common: {event_common}")
    # pprint(f"dialog_event_type: {dedialog_event_type}")
    print(f"slachCommand: {slash_command}")
    print(f"invoked_function:{invoked_function}")

    # セッションマネージャーを用意して、セッション管理に必要な情報を取得する
    session_manager = chat.session.SessionManager()
    user_id = event["user"]["name"]

    # TODO: 2023-10-13 カードメニューから各種タスク実行をする機能も入れる

    # CARD_CLICKEDイベントの処理
    if event_type == "CARD_CLICKED":
        event_forminputs = event_common.get("formInputs", dict())
        match invoked_function:
            # [各種機能のグループ分け]
            # generate_quote
            case "run_task__generate_quotes":
                print("run_task__generate_quotes")

                # 設定カードの結果はjson文字列をそのまま入れる。タスク側で変換する
                session_data = {
                    "selected_estimate_calcsheets": event_forminputs[
                        "estimate_list_checkbox"
                    ]["stringInputs"]["value"]
                }

                # タスクを実行する-> アクションレスポンスを返す
                return run_task_generate_quotes(session_data)

            # get_billing
            case "confirm__get_billing":
                print("confirm__get_billing")

                choiced_quote_list = event_forminputs["quoteitems"]["stringInputs"][
                    "value"
                ]

                session_manager.update_session(
                    user_id,
                    "get_billing",
                    chat.session.TaskState.RUNNING,
                    data={"choiced_quote_list": choiced_quote_list},
                )
                # 確認用のカードを表示する
                return confirm_get_billing(
                    session_manager.get_session(user_id, "get_billing")["data"]
                )

            case "run_task__get_billing":
                print("run_task__get_billing")
                now_session = session_manager.get_session(user_id, "get_billing")
                session_data = now_session.get("data", "{}")
                session_manager.update_session(
                    user_id, "get_billing", chat.session.TaskState.COMPLETED
                )

                # タスクを実行する-> アクションレスポンスを返す
                return runtask_get_billing(session_data)

            # run_mail_action
            case "run_task__run_mail_action":
                print("run_task__run_mail_action")

                session_data = {
                    "selected_message_id": event_forminputs["selected_message_id"][
                        "stringInputs"
                    ]["value"][0],
                    "ask_generate_projectfile": "ask_generate_projectfile"
                    in event_forminputs["run_mail_action_settings"]["stringInputs"][
                        "value"
                    ],
                    "ask_add_schedule_and_generate_estimate_calcsheet": "ask_add_schedule_and_generate_estimate_calcsheet"
                    in event_forminputs["run_mail_action_settings"]["stringInputs"][
                        "value"
                    ],
                    "ask_add_schedule_nextmonth": "ask_add_schedule_nextmonth"
                    in event_forminputs["run_mail_action_settings"]["stringInputs"][
                        "value"
                    ],
                }

                # タスクを実行する-> アクションレスポンスを返す
                return runtask_run_mail_action(session_data)

            # calc_add:動作確認用
            case "confirm__calc_add":
                print("confirm_calc_add")
                session_manager.update_session(
                    user_id,
                    "run_calc_add",
                    chat.session.TaskState.RUNNING,
                    data={
                        "first_num": int(
                            event_forminputs["first_num"]["stringInputs"]["value"][0]
                        ),
                        "second_num": int(
                            event_forminputs["second_num"]["stringInputs"]["value"][0]
                        ),
                    },
                )
                confirm_card = confirm_calc_add(
                    session_manager.get_session(user_id, "run_calc_add").get(
                        "data", "{}"
                    )
                )

                return confirm_card

            case "run_task__calc_add":
                print("runtask__calc_add")
                now_session = session_manager.get_session(user_id, "run_calc_add")
                session_data = now_session.get("data", "{}")
                session_manager.update_session(
                    user_id, "run_calc_add", chat.session.TaskState.COMPLETED
                )

                # タスクを実行する-> アクションレスポンスを返す
                return runtask_calc_add(session_data)

            # 各タスクキャンセル処理
            # TODO: 2023-10-13 このキャンセル処理は各タスクに紐づいてない。
            # セッションも除去できないので、実質動かない。タスクアプリごとにキャンセル処理を実装するか、名前をもとに解決できる手段を考える
            # チャットのレスポンスの内部にある何らかの名前が使えないかな？
            case "cancell_task":
                print("cancell")
                # セッションを終了したとする
                session_manager.update_session(
                    user_id, "run_calc_add", chat.session.TaskState.CANCELLED
                )
                return chat.card.create_card_text(
                    "cancell_task", bot_header, "キャンセルしました"
                )

            # カードクリック時に判別できなかったがあればこちらが呼ばれる
            case _:
                print("default")
                return chat.card.create_card_text(
                    "default", bot_header, "判断できなかった処理がありました。"
                )

    # 2. スラッシュコマンドの処理
    if slash_command := event["message"].get("slashCommand"):
        command_id = slash_command.get("commandId", 0)
        # TODO:2023-10-10 スラッシュコマンドでそれぞれの機能を呼び出すカードを出す機能も実装する
        match command_id:
            case "1":
                print("slash command 1: run_calc")
                # セッション初期化
                session_manager.initialize_session(
                    user_id, "run_calc", initial_data=None
                )
                return open_cofig_card_calc_add()

            # case "101":
            #     print("slash command 101: send_sagakusyoumei")
            #     print("run_task__send_sagakusyoumei")
            #     return runtask_send_sagakusyoumei()

            case "102":
                print("slash command 102: generate_quotes")

                return run_preparetask_generate_quotes()
            case "103":
                print("slash command 103: get_billing")

                return run_preparetask_get_billing()
            case "104":
                print("slash command 104: run_mail_action")

                return run_preparetask_run_mail_action()

    # 3. チャットボットのオンボーディング処理
    # 4. 通常のメッセージに対してのレスポンス
    # Case 1: The app was added to a room
    if event_type == "ADDED_TO_SPACE" and event["space"]["type"] == "ROOM":
        text = f'「{event["space"]["displayName"]}」に追加してくれてありがとう！'

    # Case 2: The app was added to a DM
    elif event_type == "ADDED_TO_SPACE" and event["space"]["type"] == "DM":
        text = f"DMに追加してくれてありがとう、{event['user']['displayName']}!"

    elif event_type == "MESSAGE":
        text = f'メッセージ: "{event["message"]["text"]}"'

    result = chat.card.create_card_text("other_event", bot_header, text)
    # pprint(result)
    return result


# Flask Routing
@app.route("/", methods=["GET"])
def home_get():
    """Respond to GET requests to this endpoint."""

    return render_template("home.html")


@app.route("/", methods=["POST"])
def home_post():
    """Respond to POST requests to this endpoint.

    All requests sent to this endpoint from Google Chat are POST
    requests.
    """

    data = request.get_json()

    resp = "Remove Space"

    if data["type"] == "REMOVED_FROM_SPACE":
        logging.info("App removed from a space")

    else:
        resp_dict = response_generator(data)
        resp = f_json.jsonify(resp_dict)

    return resp


# [END basic-bot]

if __name__ == "__main__":
    # This is used when running locally. Gunicorn is used to run the
    # application on Google App Engine. See entrypoint in app.yaml.
    app.run(host="0.0.0.0", port=8080, debug=True)
