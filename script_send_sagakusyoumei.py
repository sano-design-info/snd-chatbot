import os
import time

import redis
from rq import Queue

from task import send_sagakusyoumei


def main():
    queue = Queue(connection=redis.from_url(os.environ.get("RQ_REDIS_URL")))
    print("[task start]")

    script_task = send_sagakusyoumei.ScriptTask()

    task = queue.enqueue(script_task.execute_task)

    # タスクが終わるまで待機
    # 実行結果を待って表示する
    while True:
        if task.result:
            print(f"task result:{task.result}")
            break
        time.sleep(1)

    print("[task end]")


if __name__ == "__main__":
    main()
