import base64
import mimetypes
from email.message import EmailMessage

import pytest

from api.googleapi import create_messagedata

# create_messagedataのテスト。正しいデータが形成できてるか確認する


def test_create_messagedata():
    # テストデータ
    testdata = {
        "sender": "z@example.com",
        "to": "a@example.com",
        "cc": "b@example.com",
        "reply_to": "a@example.com",
        "subject": "test",
        "message_text": "test",
        "attachment_files": None,
    }

    # テストデータを元にメッセージを作成
    result_b64safebinary = create_messagedata(**testdata)

    # メッセージをデコードして、メッセージオブジェクトにする
    result_message = EmailMessage.from_bytes(result_b64safebinary)

    # メッセージオブジェクトの検証
    assert result_message is not None
