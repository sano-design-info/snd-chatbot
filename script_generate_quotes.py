import os
import sys
import time
from pprint import pprint

import click
import questionary
import redis
from rq import Queue

from task import generate_quotes


def dry_run(anken_quotes: list[generate_quotes.AnkenQuote]):
    """
    dry run用の関数
    args:
        anken_quotes (list[AnkenQuote]): 見積もり情報
    return:
        None
    """

    print("[dry run]anken_quotes dump:")
    # anken_number, duration, price
    pprint(
        [
            (
                anken_quote.anken_number,
                anken_quote.duration.date(),
                anken_quote.price,
            )
            for anken_quote in anken_quotes
        ]
    )


@click.command()
@click.option("--dry-run", is_flag=True, help="Dry Run Flag")
def main(dry_run):
    queue = Queue(connection=redis.from_url(os.environ.get("RQ_REDIS_URL")))

    prepare_task = generate_quotes.PrepareTask()
    prepare_job = queue.enqueue(prepare_task.execute_task)

    # 値が返ってくるまで and stateがfinishedになるまで待機
    while True:
        if prepare_job.result:
            print(f"task result:{prepare_job.result}")
            break
        time.sleep(1)

    estimate_calcsheet_list = prepare_job.result

    # 一覧から該当する見積もり計算表を取得
    selected_estimate_calcsheets = questionary.checkbox(
        "見積もりを作成する見積もり計算表を選択してください。",
        choices=[
            questionary.Choice(title=estimate_gsheet.get("name"), value=estimate_gsheet)
            for estimate_gsheet in estimate_calcsheet_list
        ],
    ).ask()

    # キャンセル処理を入れる
    if not selected_estimate_calcsheets:
        print("操作をキャンセルしました。終了します。")
        sys.exit(0)

    task_data = {"task_data": {"estimate_calcsheet_list": selected_estimate_calcsheets}}

    main_task = generate_quotes.MainTask()
    main_job = queue.enqueue(main_task.execute_task, task_data)

    while True:
        if main_job.result:
            print(f"task result:{main_job.result}")
            break
        time.sleep(1)

    print(main_job.result)


if __name__ == "__main__":
    main()
