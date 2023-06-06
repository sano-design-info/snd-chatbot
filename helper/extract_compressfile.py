import subprocess
from pathlib import Path


def extract_file(file_path: Path, output_folder: Path, password: str = None) -> bool:
    # 7-Zipがインストールされていてコマンドラインからアクセス可能か確認
    try:
        subprocess.run(["7z"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except FileNotFoundError:
        raise EnvironmentError("7-Zipがインストールされていないか、コマンドラインからアクセスできません。")

    # 入力ファイルが存在することを確認
    if not file_path.exists():
        raise FileNotFoundError(f"入力ファイル {str(file_path)} が存在しません。")

    # 出力ディレクトリが存在するか確認、なければ作成
    output_folder.mkdir(parents=True, exist_ok=True)

    # コマンドを準備: 上書きモードで解凍
    command = ["7z", "x", str(file_path), "-o" + str(output_folder), "-y"]
    if password is not None:
        command.append("-p" + password)

    # 7-Zipコマンドを実行してファイルを解凍
    result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # パスワードが間違っているか確認
    if "Wrong password?" in result.stderr.decode():
        raise ValueError(f"ファイル {str(file_path)} のパスワードが間違っています。")

    # コマンドが成功したかどうかを返す
    return result.returncode == 0
