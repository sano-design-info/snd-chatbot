import re

# 正規表現のパターン一覧

# ミスミ型式のパターン
# * ベース: MA-0000
# * グループ上下: MA-0000-UPPER
# * グループ上下: MA-0000-LOWER
# * グループ右左: MA-0000-RH
# * グループ右左: MA-0000-LH
# * グループ上下右左: MA-0000-UPPER-RH
# * グループ上下右左: MA-0000-UPPER-LH
# * グループ上下右左: MA-0000-LOWER-RH
# * グループ上下右左: MA-0000-LOWER-LH
# * グループベース修正: MA-0000-1
# * グループ上下右左で修正: MA-0000-UPPER-RH-1
MSM_ANKEN_NUMBER = re.compile(
    r"(?P<basepartnumber>MA-(?P<onlynumber>\d{4}))-?((UPPER|LOWER)?-?(RH|LH)?-?(\d{1})?)?"
)

# generate_invoiceのパターン
INVOICE_DURARION = re.compile(r"納期\s*(?P<durarion>\d{1,2}/\d{1,2})")

# itemparserのパターン

# "sheetname!A1"の形式でマッチする。
RANGE_ADDR_PATTERN = re.compile(
    r"^(?P<sheetname>.*)!(?P<firstcolumn>[A-Z]+)(?P<firstrow>\d+)"
)
