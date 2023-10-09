import pytest
from pathlib import Path

from task.run_mail_action import filter_msm_katasiki_by_filename


# pick_msm_katasiki_by_renrakukoumoku_filename のテスト
# ファイルパスと抽出した案件番号（例:0000）のパターンを用意してテストする。ファイルパスは架空のもの
@pytest.mark.parametrize(
    ("target_path", "expected"),
    [
        (Path("./tests/MA-1007_標準ガス配管図連絡項目_118052ユニプレス_岩下様_DB00371.xlsx"), "1007"),
        (
            Path(
                "./tests/MA-1008 標準ガス配管図連絡項目_118052ユニプレス_叶様_DB00241_DB00248.xlsx",
            ),
            "1008",
        ),
    ],
)
def test_pick_anken_number(target_path, expected):
    # ファイルパスを適当に用意
    assert filter_msm_katasiki_by_filename(target_path) == expected
