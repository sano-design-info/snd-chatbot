import os
import time

import questionary
import redis
from rq import Queue
from itemparser import ExpandedMessageItem

from task import run_mail_action


def main() -> None:
    print("[Start Process...]")

    queue = Queue(connection=redis.from_url(os.environ.get("RQ_REDIS_URL")))

    prepare_task = run_mail_action.PrepareTask()
    prepare_job = queue.enqueue(prepare_task.execute_task)

    # 値が返ってくるまで and stateがfinishedになるまで待機

    while True:
        if prepare_job.result:
            print(f"task result:{prepare_job.result}")
            break
        time.sleep(1)

    messages = prepare_job.result

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

    # TODO:2023-09-14 タスク実行のために、ExmandedMessageItemを渡さないようにする。
    selected_message: ExpandedMessageItem = questionary.select(
        "メールの選択をしてください", choices=message_item_and_labels
    ).ask()

    # このタイミングでメール選択がされていなければ終了
    if not selected_message:
        print("[Cancell Process...]")
        exit()

    # その他質問を確認
    ask_generate_projectfile = questionary.confirm(
        "プロジェクトファイルを生成しますか？(修正案件の場合は作成しないこと 例: MA-0000-1)", True
    ).ask()

    if not ask_generate_projectfile:
        print("[Cancell Process...]")
        exit()

    ask_add_schedule_and_generate_estimate_calcsheet = questionary.confirm(
        "スケジュール表追加と見積計算表の作成を行いますか？（プロジェクトファイル再作成時はFalseで）", True
    ).ask()

    # スキップする場合、↑の質問がFalseになる場合
    ask_add_schedule_nextmonth = (
        questionary.confirm("スケジュール表追加時に入金日を予定月の来月にしますか？", False)
        .skip_if(
            ask_add_schedule_and_generate_estimate_calcsheet is False, default=False
        )
        .ask()
    )

    # 選択後、処理開始していいか問い合わせして実行
    comefirm_check = questionary.confirm("run Process?", False).ask()

    if not comefirm_check:
        print("[Cancell Process...]")
        exit()

    # TODO:2023-09-28 [task start] 選択肢, メールのID
    # TODO:2023-10-09  今はseleceted_messageにしているが、ここはメールのIDを渡すようにする方がいいかな。
    # redis側が保持してくれるなら気にしなくてもいいかもだけど
    task_data = {
        "selected_message": selected_message,
        "ask_generate_projectfile": ask_generate_projectfile,
        "ask_add_schedule_and_generate_estimate_calcsheet": ask_add_schedule_and_generate_estimate_calcsheet,
        "ask_add_schedule_nextmonth": ask_add_schedule_nextmonth,
    }

    main_task = run_mail_action.MainTask()

    main_job = queue.enqueue(main_task.execute_task, {"task_data": task_data})

    while True:
        if main_job.result:
            print(f"task result:{main_job.result}")
            break
        time.sleep(1)

    print(main_job.result)

    print("[End Process...]")
    exit()


if __name__ == "__main__":
    main()
