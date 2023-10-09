from datetime import datetime
import os
import time

import questionary
import redis
from rq import Queue

from task import get_billing


# 本日の日付を全体で使うためにここで宣言
today_datetime = datetime.now()


def main():
    queue = Queue(connection=redis.from_url(os.environ.get("RQ_REDIS_URL")))

    # 見積一覧から必要情報を収集
    prepare_task = get_billing.PrepareTask()
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

    ask_choiced_quote_list: list[get_billing.QuoteData] = questionary.checkbox(
        "請求書にする見積書を選択してください。", choices=choice_list
    ).ask()

    # TODO:2023-10-09 [chat end]: 渡すデータ: 見積もりの取得ID一覧（taskでidでg
    # etしてくる）, billing_dataのdictフォーマット
    # これもタスクにするか、そのままこれで動かすかは決める。
    # タスクのほうがchat側へ結果を渡しやすいからいいかも

    # 見積書の情報を元に、金額の合計を出す
    billing_data = get_billing.BillingInfo(
        sum((i.price for i in ask_choiced_quote_list)),
        "ガススプリング配管図作製費",
        f"{today_datetime:%Y年%m月}請求分",
    )

    # ここで請求書情報を出して、こちらの検証と正しいか確認
    print(
        f"""
    [請求情報]
    件数: {len(ask_choiced_quote_list)}
    合計金額:{billing_data.price}
    """
    )
    ask_runtask = questionary.confirm("請求書送付メールを作成しますか？").ask()

    if not ask_runtask:
        # TODO:2023-10-03 ここは再度実行のために、再びprepareを呼び出す実装にする。再帰関数？
        exit("キャンセルしました。再度実行しなおしてください")

    # TODO:2023-09-28 [chat end]: 渡すデータ:
    # 見積もりの取得ID一覧（taskでidでgetしてくる->QuoteDataにする？）
    # -> 今はこれはやらないで、QuoteDataが入ることを想定で良し。
    # billing_dataのdictフォーマット

    task_data = {
        "choiced_quote_id": ask_choiced_quote_list,
        "billing_data": billing_data,
    }

    main_task = get_billing.MainTask()
    main_job = queue.enqueue(main_task.execute_task, {"task_data": task_data})

    while True:
        if main_job.result:
            print(f"task result:{main_job.result}")
            break
        time.sleep(1)

    print(main_job.result)


if __name__ == "__main__":
    main()
