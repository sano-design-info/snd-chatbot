import os
import zipfile
from pathlib import Path


# TODO:2022-12-23 zipファイル解凍が失敗したらエラーとして解凍しないで進める
def extract_zipfile(zipfile_path: Path, outdir: Path) -> None:
    """zipファイルのパスを渡して、outのdirパスへ解凍する"""

    def _rename(info: zipfile.ZipInfo) -> None:
        """ヘルパー: `ZipInfo` のファイル名を SJIS でデコードし直す"""
        LANG_ENC_FLAG = 0x800
        encoding = "utf-8" if info.flag_bits & LANG_ENC_FLAG else "cp437"

        print(info.orig_filename.encode(encoding))

        # 最初cp932でデコードして、unicodeerrorになったらutf8にする処理にする
        try:
            info.filename = info.orig_filename.encode(encoding).decode("cp932")
        except UnicodeDecodeError:
            info.filename = info.orig_filename.encode(encoding).decode("utf-8")

    try:
        with zipfile.ZipFile(zipfile_path) as zfile:
            # zipアーカイブのファイル単位でループ
            for info in zfile.infolist():

                # ファイル名のデコード処理
                _rename(info)
                if os.sep != "/" and os.sep in info.filename:
                    info.filename = info.filename.replace(os.sep, "/")

                # ファイル単位で解凍する
                zfile.extract(info, outdir)
    except zipfile.BadZipFile:
        print(f"Zipファイルの解凍に失敗しました。不正なZipファイルの可能性があります :{zipfile_path}")
