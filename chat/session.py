from dataclasses import dataclass
import redis
import json
import os
from enum import Enum


# タスク状態をEnumで定義
class TaskState(Enum):
    INITIAL = "initial"
    RUNNING = "running"
    COMPLETED = "completed"
    CANCELLED = "cancelled"
    ERROR = "error"


@dataclass
class SessionData:
    current_task: str
    state: TaskState
    data: dict


class SessionManager:
    """
    チャットで利用するセッションマネージャーです。
    初期化を行い、更新、取得を行うことができます。
    タスク状態であるTaskStateはEnumで定義しています。

    # TODO:2023-10-10
    # 状態管理はタスク状態で行うようにする。
    # COMPLETED, CANCELLED, ERRORの場合は、redis上から削除するような仕様にする（できる？

    # session_manager = SessionManager()
    # data = {"key1": "value1", "key2": "value2"}
    # 初期化...
    # session_manager.initialize_session("user_id", "task_name")
    # 更新...
    # session_manager.update_session("user_id", "task_name", TaskState.RUNNING, data)
    # 取得...
    # session_info = session_manager.get_session("user_id", "task_name")
    """

    def __init__(self):
        redis_url = os.environ.get("SESSION_REDIS_URL", "redis://localhost:6379/10")
        self.redis = redis.StrictRedis.from_url(redis_url, decode_responses=True)

    def _generate_session_key(self, user_id: str, task_name: str) -> str:
        return f"{user_id}:{task_name}"

    def initialize_session(self, user_id: str, task_name: str, initial_data=None):
        key = self._generate_session_key(user_id, task_name)
        session_data = {
            "current_task": task_name,
            "state": TaskState.INITIAL.value,
            # dataはjson文字列にして保存する
            "data": json.dumps(initial_data or {}),
        }
        self.redis.hmset(key, session_data)

    def get_session(self, user_id: str, task_name: str) -> dict:
        key = self._generate_session_key(user_id, task_name)

        # dataはjson文字列なので、loadする
        session_data = self.redis.hgetall(key)
        session_data["data"] = json.loads(session_data["data"])

        return session_data

    def update_session(self, user_id: str, task_name: str, state: TaskState, data=None):
        key = self._generate_session_key(user_id, task_name)
        session_data = {"state": state.value, "data": json.dumps(data or {})}
        self.redis.hmset(key, session_data)
