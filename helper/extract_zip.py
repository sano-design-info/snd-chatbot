import os
import zipfile
from pathlib import Path


def extract_zipfile(zipfile_path: Path, outdir: Path) -> None:
    def _rename(info: zipfile.ZipInfo) -> None:
        """ヘルパー: `ZipInfo` のファイル名を SJIS でデコードし直す"""
        LANG_ENC_FLAG = 0x800
        encoding = "utf-8" if info.flag_bits & LANG_ENC_FLAG else "cp437"

        # cp932でdecodeすると、それ以外のエンコーディングで失敗するので、エンコーディング判断が必要
        print(info.orig_filename.encode(encoding))

        # TODO:2022-12-16 chardetライブラリを使って判断してみたけど出来なさそうだったので、最初cp932でデコードして、unicodeerrorになったらutf8にする処理にする
        try:
            info.filename = info.orig_filename.encode(encoding).decode("cp932")
        except UnicodeDecodeError:
            info.filename = info.orig_filename.encode(encoding).decode("utf-8")

    with zipfile.ZipFile(zipfile_path) as zfile:
        for info in zfile.infolist():
            # cop932で固定するとまずいので、
            _rename(info)
            # info.filename = info.orig_filename.encode("cp437").decode("cp932")
            if os.sep != "/" and os.sep in info.filename:
                info.filename = info.filename.replace(os.sep, "/")
        zfile.extract(info, outdir)
