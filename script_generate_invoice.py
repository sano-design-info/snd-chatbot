from datetime import datetime
import os
import time

import questionary
import redis
from rq import Queue

from task import generate_invoice


# 本日の日付を全体で使うためにここで宣言
today_datetime = datetime.now()


def main():
    queue = Queue(connection=redis.from_url(os.environ.get("RQ_REDIS_URL")))

    # 見積一覧から必要情報を収集
    prepare_task = generate_invoice.PrepareTask()
    prepare_job = queue.enqueue(prepare_task.execute_task)

    # 値が返ってくるまで and stateがfinishedになるまで待機
    while True:
        if prepare_job.result:
            print(f"task result:{prepare_job.result}")
            break
        time.sleep(1)

    quote_checked_list = prepare_job.result
    choice_list = [
        questionary.Choice(
            title=f"{quotedata.only_katasiki} |{quotedata.price} | {quotedata.durarion}",
            value=quotedata,
            checked=checkbool,
        )
        for quotedata, checkbool in quote_checked_list
    ]

    ask_choiced_quote_list: list[generate_invoice.QuoteData] = questionary.checkbox(
        "請求書にする見積書を選択してください。", choices=choice_list
    ).ask()

    # 見積書の情報を元に、金額の合計を出す
    invoice_data = generate_invoice.generate_invoice_data(ask_choiced_quote_list)

    # ここで請求書情報を出して、こちらの検証と正しいか確認
    print(
        f"""
    [請求情報]
    件数: {len(ask_choiced_quote_list)}
    合計金額:{invoice_data.price}
    """
    )
    ask_runtask = questionary.confirm("請求書送付メールを作成しますか？").ask()

    # キャンセルの場合は再度実行とする
    if not ask_runtask:
        exit("キャンセルしました。再度実行しなおしてください")

    task_data = {"task_data": {"choiced_quote_list": ask_choiced_quote_list}}

    main_task = generate_invoice.MainTask()
    main_job = queue.enqueue(main_task.execute_task, task_data)

    while True:
        if main_job.result:
            print(f"task result:{main_job.result}")
            break
        time.sleep(1)

    print(f"見積書を作成しました。: {main_job.result}")


if __name__ == "__main__":
    main()
