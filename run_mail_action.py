import html
import itertools
import shutil
from datetime import datetime
from pathlib import Path

import copier
import openpyxl
import questionary
from bs4 import BeautifulSoup
from dateutil.relativedelta import relativedelta
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from jinja2 import Environment, FileSystemLoader

from api import googleapi
from helper import EXPORTDIR_PATH, decode_base64url, load_config, extract_compressfile
from helper.regexpatterns import MSM_ANKEN_NUMBER

from itemparser import ExpandedMessageItem


# generate Path
parent_dirpath = Path(__file__).parents[0]
exportfiles_dirpath = (
    EXPORTDIR_PATH
    / "export_files"
    / f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
)
attachment_dirpath = exportfiles_dirpath / "attachments"

GOOGLE_API_SCOPES = googleapi.API_SCOPES

# load config
config = load_config.CONFIG
gsheet_tmp_dir_ids = config.get("google").get("GSHEET_TMP_DIR_IDS")
msm_gas_boilerplate_path = config.get("other").get("MSM_GAS_BOILERPLATE_PATH")
schedule_sheet_id = config.get("google").get("SCHEDULE_SHEET_ID")
sheet_name = config.get("google").get("TABLE_SEARCH_RANGE")
estimate_template_gsheet_id = config.get("google").get("ESTIMATE_TEMPLATE_GSHEET_ID")
copy_project_dir_dest_path = Path(config.get("other").get("COPY_PROJECT_DIR_DEST_PATH"))

target_userid = config.get("google").get("GMAIL_USER_ID")

# google api service
google_cred: Credentials = googleapi.get_cledential(GOOGLE_API_SCOPES)
drive_service = build("drive", "v3", credentials=google_cred)
sheet_service = build("sheets", "v4", credentials=google_cred)
gmail_service = build("gmail", "v1", credentials=google_cred)


def generate_dirs() -> None:
    exportfiles_dirpath.mkdir(exist_ok=True, parents=True)
    attachment_dirpath.mkdir(exist_ok=True)


def generate_mail_printhtml(
    messageitem: ExpandedMessageItem, attachment_dirpath: Path
) -> None:
    # メール印刷用HTML生成

    # TODO:2023-04-07 この部分はExpandMessageItemへ移動する。今後の課題

    # mimetypeがplaneかhtmlで分ける
    messages_text_parts = next(
        (i for i in messageitem.body_parts if "text/html" in i.get("mimeType")), None
    )
    if not messages_text_parts:
        messages_text_parts = next(
            (i for i in messageitem.body_parts if "text/plain" in i.get("mimeType"))
        )

    mail_body = ""
    b64decoded_mail_body = decode_base64url(messages_text_parts.get("body").get("data"))

    mail_html_bs4 = BeautifulSoup(b64decoded_mail_body, "html.parser")

    # htmlの場合、imgタグを除去する
    if mail_body := mail_html_bs4.body:
        for t in mail_body.find_all("img"):
            t.decompose()
    # textの場合、改行タグをhtmlの<br>へ置き換え
    else:
        # decodeしないで、改行タグをhtmlの<br>へ置き換え
        mail_body = b64decoded_mail_body.decode("utf8").replace("\r\n", "<br>")

    # jinja2埋込
    # テンプレート読み込み
    env = Environment(
        loader=FileSystemLoader(str((parent_dirpath / "itemparser")), encoding="utf8")
    )
    tmpl = env.get_template("export.html.jinja")

    # 設定ファイル読み込み
    params = {
        "export_html": mail_body,
        "message_title": html.escape(messageitem.title),
        "message_from_address": html.escape(messageitem.from_address),
        "message_cc_address": html.escape(messageitem.cc_address),
        "message_datetime": messageitem.datetime_,
    }
    # レンダリングして出力
    with (attachment_dirpath / Path("./メール本文印刷用ファイル.html")).open(
        "w", encoding="utf8"
    ) as exp_mail_hmtl:
        exp_mail_hmtl.write(tmpl.render(params))


def generate_pdf_by_renrakukoumoku_excel(
    attachment_dirpath: Path,
) -> None:
    # 連絡項目の印刷用PDFファイル生成
    target_filepath = next(attachment_dirpath.glob("*MA-*.xlsx"))
    if not target_filepath:
        print("cant generate_pdf_byrenrakuexcel")
        return None

    try:
        upload_results = googleapi.upload_file(
            drive_service,
            target_filepath,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/vnd.google-apps.spreadsheet",
            gsheet_tmp_dir_ids,
        )
        print(f"upload_results: {upload_results}")

        googleapi.save_gdrive_file(
            drive_service,
            upload_results.get("id"),
            "application/pdf",
            attachment_dirpath / Path("./連絡項目印刷用ファイル.pdf"),
        )

        # post-porcess: Googleドキュメントに一時保持した配管連絡項目を除去する
        delete_tmp_excel_result = googleapi.delete_file(
            drive_service, file_id=upload_results.get("id"), fields="id"
        )
        print(f"deleted file: {delete_tmp_excel_result}")

    except HttpError as error:
        # TODO:2022-12-09 エラーハンドリングは基本行わずここで落とすこと
        print(f"An error occurred: {error}")


def filter_msm_katasiki_by_filename(filepath: Path) -> str:
    """
    ボイラープレートと見積書作成時に必要になるミスミ型式番号を取得
    取得ができない場合は0000を用意

    TODO:2023-09-14 この注記はdocstringへ書く
    ここでは、追加案件: MA-0000-1 のような 表記は対応していない。
    このスクリプト上で追加案件を対応することは無いと思われる。
    （追加案件を対応する場合は、泥臭いけど全体をmatchさせてから、
    groupsの最後で -1のような追加表現があるかを確認してあれば、
    抽出したファイル名にして返せばよいと思う。
    今後発生したら実装しよう。）

    args:
        filepath: ファイルパス
    return:
        msm_katasiki_num: 型式番号
    """
    msm_katasiki_num = "0000"

    if katasiki_matcher := MSM_ANKEN_NUMBER.match(str(filepath.name)):
        msm_katasiki_num = katasiki_matcher.group("onlynumber")

    return msm_katasiki_num


def generate_projectdir(attachment_dirpath: Path, export_dirpath: Path) -> None:
    """
    プロジェクトフォルダを生成します。
    ボイラープレートをコピーして、添付ファイルをコピーします。

    手順:
    * ボイラープレートで利用する型式番号を取得
    * ボイラープレートをコピー
    * 添付ファイルがあれば解凍してコピー

    args:
        attachment_dirpath: 添付ファイルのパス
        export_dirpath: プロジェクトフォルダの出力先パス
    return:
        None
    """
    export_project_dir = export_dirpath / "proj_dir"
    export_project_dir.mkdir(exist_ok=True)

    # ファイルパスから型式を取り出す
    target_filepath = next(attachment_dirpath.glob("*MA-*.xlsx"))
    if not target_filepath:
        print("cant create projectdir")
        return None

    msm_katasiki_num = filter_msm_katasiki_by_filename(target_filepath)

    # ボイラープレートからディレクトリ生成
    boilerplate_config = {
        "project_name": msm_katasiki_num,
    }
    generated_copier_worker = copier.run_copy(
        msm_gas_boilerplate_path,
        export_project_dir,
        data=boilerplate_config,
    )

    # 添付ファイルを解凍する
    for attachment_zfile in itertools.chain(
        attachment_dirpath.glob("*.zip"), attachment_dirpath.glob("*.lzh")
    ):
        print(attachment_zfile)
        try:
            _ = extract_compressfile.extract_file(attachment_zfile, attachment_dirpath)
        except ValueError as e:
            # パスワード入れずに処理 or 間違っている場合はエラー。その場合はスキップする
            print(f"パスワードあり圧縮ファイルなのでスキップしています -> {e}")

    # 添付ファイルをコピーする
    # TODO:2022-12-16 プロジェクトフォルダの名称は環境変数化したほうがいいかも？
    project_dir = generated_copier_worker.dst_path / f"ミスミ配管図MA-{msm_katasiki_num}納期 -"
    shutil.copytree(attachment_dirpath, (project_dir / "資料"), dirs_exist_ok=True)


# TODO:2023-05-24 この関数はスケジュール表更新でも使うので、このスクリプトから独立させる予定
def add_schedule_spreadsheet(
    attachment_dirpath: Path, nyukin_nextmonth: bool = False
) -> None:
    """
    GoogleSheetのミスミスケジュール表にスケジュールを追加します。

    args:
        attachment_dirpath: 添付ファイルのパス
        nyukin_nextmonth: 入金日を予定月の来月にするかどうか
    return:
        None
    """
    target_filepath = next(attachment_dirpath.glob("*MA-*.xlsx"))
    if not target_filepath:
        print("cant add schedule")
        return None

    msm_katasiki_num = filter_msm_katasiki_by_filename(target_filepath)

    # エンドユーザー: 今は列がないので登録しない
    # renrakukoumoku_range_enduser = "D10"
    # 顧客
    renrakukoumoku_range_kokyaku = "D6"

    # openpyxlで必要な位置から値を取り出す
    renrakukoumoku_wb = openpyxl.load_workbook(target_filepath)
    renrakukoumoku_ws = renrakukoumoku_wb.active
    add_schedule_kokyaku = renrakukoumoku_ws[renrakukoumoku_range_kokyaku].value

    renrakukoumoku_wb.close()

    # 型式
    add_schedule_msm_katasiki = f"MA-{msm_katasiki_num}"
    # 開始日: 実行日でよし
    now_datetime = datetime.now()
    add_schedule_start_datetime = now_datetime.strftime("%Y/%m/%d")
    # 振込タイミング
    nyukin_month = 1
    if nyukin_nextmonth:
        nyukin_month = 2
    add_schedule_hurikomiduki = now_datetime.replace(day=1) + relativedelta(
        months=nyukin_month
    )

    append_values = [
        [
            "=row() - 5",
            "ミスミ",
            add_schedule_msm_katasiki,
            "友希",
            "",
            add_schedule_start_datetime,
            "",
            "",
            "",
            f"{add_schedule_hurikomiduki.month}月末",
            "",
            "",
            add_schedule_kokyaku,
        ]
    ]

    # スケジュール表の一番後ろの行へ追加する]
    print(f"append_values: {append_values}")

    schedule_gsheet = googleapi.append_sheet(
        sheet_service,
        schedule_sheet_id,
        sheet_name,
        append_values,
        "USER_ENTERED",
        "INSERT_ROWS",
    )

    print(f"schedule updated -> {schedule_gsheet.get('updates')}")


def generate_estimate_calcsheet(attachment_dirpath: Path) -> None:
    """
    見積もり計算表を作成します。

    手順:
    * 添付ファイル内にある連絡項目のファイル名から型式番号を取得
    * 見積もり計算表のテンプレートをコピー

    args:
        attachment_dirpath: 添付ファイルのパス
    return:
        None
    """

    target_filepath = next(attachment_dirpath.glob("*MA-*.xlsx"))
    if not target_filepath:
        print("cant add schedule")
        return None

    msm_katasiki_num = filter_msm_katasiki_by_filename(target_filepath)

    try:
        # 見積書のコピーを作成する。
        copy_template_results = googleapi.copy_file(
            drive_service,
            estimate_template_gsheet_id,
            "id,name",
        )

        # コピー先のファイル名を変更する。suffixの文字を型式番号に置き換える
        template_suffix = "[ミスミ型番] のコピー"
        rename_body = {
            "name": copy_template_results.get("name").replace(
                template_suffix, msm_katasiki_num
            )
        }
        rename_estimate_gsheet_result = googleapi.update_file(
            drive_service,
            copy_template_results.get("id"),
            rename_body,
            "name",
        )

        print(
            f"estimate calcsheet copy and renamed -> {rename_estimate_gsheet_result.get('name')}"
        )
    except HttpError as error:
        # TODO:2022-12-09 エラーハンドリングは基本行わずここで落とすこと
        print(f"An error occurred: {error}")


def copy_projectdir(export_path: Path) -> None:
    """
    プロジェクトファイルを所定の場所へコピーします。

    args:
        export_path: 出力先パス
    return:
        None
    """
    # プロジェクトフォルダのパスを用意する
    export_prij_path = next((export_path / "proj_dir").glob("ミスミ配管図*"))

    print(f"プロジェクトファイルの現在の位置: {export_prij_path}")
    print(f"プロジェクトファイルをコピーするパス: {copy_project_dir_dest_path}")

    shutil.copytree(
        export_prij_path,
        copy_project_dir_dest_path / export_prij_path.name,
        dirs_exist_ok=True,
    )


def main() -> None:
    print("[Start Process...]")

    messages: list[ExpandedMessageItem] = []
    try:
        # Call the Gmail API

        # 該当メールのスレッド検索
        threads = googleapi.search_threads(gmail_service, "label:snd-ミスミ (*MA-*)")

        # 上位10件のスレッド -> メッセージを取得。
        # スレッドに紐づきが2件ぐらいのメッセージの部分でのもので十分かな
        if threads:
            top_threads = list(itertools.islice(threads, 0, 10))
            for thread in top_threads:
                # threadsのid = threadsの一番最初のmessage>idなので、そのまま使う
                message_id = thread.get("id", "")

                # スレッドの数が2以上 = すでに納品済みと思われるので削る。
                # TODO:2022-12-09 ここは2件以上でもまだやり取り中だったりする場合もあるので悩ましい
                # （数見るだけでもいいかもしれない
                thread_result = googleapi.get_thread_by_message_id(
                    gmail_service, message_id, user_id=target_userid, fields="messages"
                )

                if len(thread_result.get("messages")) <= 2:
                    # スレッドの一番先頭にあるメッセージを取得する
                    messages.append(
                        ExpandedMessageItem(
                            gmail_message=thread_result.get("messages")[0]
                        )
                    )

    except HttpError as error:
        # TODO:2022-12-09 エラーハンドリングは基本行わずここで落とすこと
        print(f"[search thread] An error occurred: {error}")
        exit()

    if not messages:
        print("cant find messages...")
        exit()

    # 取得が面倒なので最初から必要な値を取り出す
    message_item_and_labels = []
    for message in messages:
        # 送信日, タイトル
        choice_label = [
            ("class:text", f"{message.datetime_}"),
            ("class:highlighted", f"{message.title}"),
        ]
        message_item_and_labels.append(
            questionary.Choice(title=choice_label, value=message)
        )

    # 上位10のスレッドから > メッセージの最初取り出して、その中から選ぶ
    print("[Select Mail...]")

    # TODO:2023-09-14 タスク実行のために、ExmandedMessageItemを渡さないようにする。

    selected_message: ExpandedMessageItem = questionary.select(
        "メールの選択をしてください", choices=message_item_and_labels
    ).ask()

    # このタイミングでメール選択がされていなければ終了
    if not selected_message:
        print("[Cancell Process...]")
        exit()

    # その他質問を確認
    ask_generate_projectfile = questionary.confirm(
        "プロジェクトファイルを生成しますか？(修正案件の場合は作成しないこと 例: MA-0000-1)", True
    ).ask()

    if not ask_generate_projectfile:
        print("[Cancell Process...]")
        exit()

    ask_add_schedule_and_generate_estimate_calcsheet = questionary.confirm(
        "スケジュール表追加と見積計算表の作成を行いますか？（プロジェクトファイル再作成時はFalseで）", True
    ).ask()

    # スキップする場合、↑の質問がFalseになる場合
    ask_add_schedule_nextmonth = (
        questionary.confirm("スケジュール表追加時に入金日を予定月の来月にしますか？", False)
        .skip_if(
            ask_add_schedule_and_generate_estimate_calcsheet is False, default=False
        )
        .ask()
    )

    # 選択後、処理開始していいか問い合わせして実行
    comefirm_check = questionary.confirm("run Process?", False).ask()

    if not comefirm_check:
        print("[Cancell Process...]")
        exit()

    print("[Generate Dirs...]")
    generate_dirs()

    print("[Save Attachment file and mail image]")

    # TODO:2023-09-14 ここはExpandedMessageItemへ移動する。
    # メール本文にimgファイルがある場合はそれを取り出す
    # multipart/relatedの時にあるので、それを狙い撃ちで取る

    if selected_message.body_related:
        message_imgs = [
            i
            for i in selected_message.body_related.get("parts")
            if "image" in i.get("mimeType")
        ]
        for msg_img in message_imgs:
            googleapi.save_attachment_file(
                gmail_service,
                selected_message.id,
                msg_img.get("body").get("attachmentId"),
                attachment_dirpath / msg_img.get("filename"),
            )
    # 添付ファイルの保持
    message_attachmentfiles = [
        i
        for i in selected_message.payload.get("parts")
        if "application" in i.get("mimeType")
    ]

    for msg_attach in message_attachmentfiles:
        googleapi.save_attachment_file(
            gmail_service,
            selected_message.id,
            msg_attach.get("body").get("attachmentId"),
            attachment_dirpath / msg_attach.get("filename"),
        )

    # 各種機能を呼び出す

    print("[Generate Mail Printable PDF]")
    generate_mail_printhtml(selected_message, attachment_dirpath)

    print("[Generate Excel Printable PDF]")
    generate_pdf_by_renrakukoumoku_excel(attachment_dirpath)

    if ask_generate_projectfile:
        print("[Generate template dirs]")
        generate_projectdir(attachment_dirpath, exportfiles_dirpath)

        print("[copy project dir]")
        copy_projectdir(exportfiles_dirpath)
    else:
        print("[Not Generate template dirs]")

    if ask_add_schedule_and_generate_estimate_calcsheet:
        print("[append schedule]")
        add_schedule_spreadsheet(attachment_dirpath, ask_add_schedule_nextmonth)

        print("[add estimate calcsheet]")
        generate_estimate_calcsheet(attachment_dirpath)
    else:
        print("[Not Add Scuedule, Generate estimate calcsheet]")

    print("[End Process...]")
    exit()


if __name__ == "__main__":
    main()
