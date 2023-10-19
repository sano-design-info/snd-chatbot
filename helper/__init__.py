import base64
import json
from pathlib import Path

# dirpath
ROOTDIR = Path(__file__).parents[1]
EXPORTDIR_PATH = ROOTDIR / "exportdir"
EXPORTDIR_PATH.mkdir(parents=True, exist_ok=True)


def decode_base64url(base64_s) -> bytes:
    """
    urlセーフなbase64文字列をデコードする
    ref: https://stackoverflow.com/questions/2941995/python-ignore-incorrect-padding-error-when-base64-decoding

    Args:
        base64_s (str): base64文字列

    Returns:
        bytes: デコード後のbytes
    """
    return base64.urlsafe_b64decode(base64_s) + b"=" * (4 - (len(base64_s) % 4))


def convert_dataclass_to_jsonhash_str(jsonhash_str: str, dataclass_):
    """
    設定カードから返されたjsonハッシュ文字列を、dataclassの形式に変換する

    文字列はjsonのハッシュを想定して変換する。
    dataclassはinit=Falseのフィールドは除外すること

    Args:
        dialog_str (str): 設定カードから返された文字列
        dataclass_ ([type]): 変換先のdataclass

    Returns:
        dataclass_: 変換後のdataclass

    Raises:
        ValueError: dialog_strがjson文字列ではない場合

    >>> @dataclass
    ... class AddressData:
    ...     name: str
    ...     age: int
    ...     address: str = field(init=False)

    >>> conv_addressdata = convert_dataclass_to_jsonhash_str('{"name": "test", "age": 20}', AddressData)
    >>> conv_addressdata.address = "Asia/Tokyo"
    >>> conv_addressdata.name
    'test name'

    """

    # json文字列をdictに変換する。
    converted_jsonobject = json.loads(jsonhash_str)

    if not isinstance(converted_jsonobject, dict):
        raise ValueError("dialog_str is not json object string")

    # dataclassのフィールド名を取得する。init=Trueのフィールドのみを取得する
    dataclass_field_names = (
        dataclass_fieldname
        for dataclass_fieldname, dataclass_field in dataclass_.__dataclass_fields__.items()
        if dataclass_field.init is True
    )

    return dataclass_(
        **{
            dataclass_field_name: converted_jsonobject.get(dataclass_field_name)
            for dataclass_field_name in dataclass_field_names
        }
    )
