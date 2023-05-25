import base64
import email

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

    # テストデータを元にメッセージを作成 (base64エンコードされたバイナリ)
    result_b64safebinary = create_messagedata(**testdata)

    # base64でエンコードされたメッセージをMessageオブジェクトに変換
    result_message = email.message_from_bytes(
        base64.urlsafe_b64decode(result_b64safebinary)
    )

    # メッセージオブジェクトの検証
    assert result_message is not None
    assert result_message["From"] == testdata["sender"]
    assert result_message["To"] == testdata["to"]
    assert result_message["Cc"] == testdata["cc"]
    assert result_message["Reply-To"] == testdata["reply_to"]
    assert result_message["Subject"] == testdata["subject"]
