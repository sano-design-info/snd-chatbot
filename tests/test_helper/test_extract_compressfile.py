import pytest
from pathlib import Path

from helper.extract_compressfile import extract_file

testdata_path = Path(__file__).parents[1] / "testdata"
testdata_extracted_path = testdata_path / "extracted"


# teardown: 解凍したファイルやファイルを入れたフォルダを削除する
@pytest.fixture(scope="function")
def remove_extracted_file():
    yield "run test_extract_file()"
    for file in testdata_extracted_path.glob("*"):
        file.unlink()
    testdata_extracted_path.rmdir()


def test_extract_file(remove_extracted_file):
    # testdata内のzip, lzhファイルを解凍する。それぞれファイルが1つづつ入ってる
    extract_file(testdata_path / "blank_into_zip.zip", testdata_extracted_path)
    extract_file(testdata_path / "test_lzh_data__G01075.lzh", testdata_extracted_path)

    # 検証: ファイルは2つ出てきたか
    assert len(list(testdata_extracted_path.glob("*"))) == 2
