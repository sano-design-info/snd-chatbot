import re

# 正規表現のパターン一覧

# ミスミ型式のパターン
MSM_ANKEN_NUMBER = re.compile(r"(?P<basepartnumber>MA-(?P<onlynumber>\d{4})).*")
