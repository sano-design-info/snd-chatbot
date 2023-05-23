import base64
import html
import io
import itertools
from datetime import datetime
from pathlib import Path
import shutil
from bs4 import BeautifulSoup
import copier
from dateutil.relativedelta import relativedelta
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from jinja2 import Environment, FileSystemLoader
import openpyxl

import questionary
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from api import googleapi
from helper import extract_zip
from itemparser import (
    ExpandedMessageItem,
    copy_project_dir_dest_path,
    decode_base64url,
    estimate_template_gsheet_id,
    gsheet_tmmp_dir_ids,
    msm_gas_boilerplate_url,
    parent_dirpath,
    pick_msm_katasiki_by_renrakukoumoku_filename,
    schedule_sheet_id,
    table_search_range,
)

target_userid = "me"

# generate Path
parent_dirpath = Path(__file__).parents[0]
export_dirpath = (
    parent_dirpath
    / "export_files"
    / f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
)
attachment_dirpath = export_dirpath / "attachments"

GOOGLE_API_SCOPES = googleapi.API_SCOPES


def save_attachment_file(
    service, filename: str, message_id: str, attachment_id: str
) -> None:
    attachfile_data = (
        service.users()
        .messages()
        .attachments()
        .get(
            userId=target_userid,
            messageId=message_id,
            id=attachment_id,
        )
        .execute()
    )
    print((attachment_dirpath / Path(filename)))
    with (attachment_dirpath / Path(filename)).open("wb") as attachmentfiile:
        attachmentfiile.write(base64.urlsafe_b64decode(attachfile_data.get("data")))


def generate_dirs() -> None:
    export_dirpath.mkdir(exist_ok=True, parents=True)
    attachment_dirpath.mkdir(exist_ok=True)


def generate_mail_printhtml(
    messageitem: ExpandedMessageItem, attachment_dirpath: Path
) -> None:
    # メール印刷用HTML生成

    # TODO:2023-04-07 この部分はExpandMessageItemへ移動する
    # TODO:2022-12-27 個々の実装はhtmlかplaneで最初から分けたほうがいいかも
    messages_text_parts = next(
        (i for i in messageitem.body_parts if "text/html" in i.get("mimeType")), None
    )
    if not messages_text_parts:
        messages_text_parts = next(
            (i for i in messageitem.body_parts if "text/plain" in i.get("mimeType"))
        )

    # imgタグを除去する
    mail_body = ""
    b64decoded_mail_body = decode_base64url(messages_text_parts.get("body").get("data"))

    mail_html_bs4 = BeautifulSoup(b64decoded_mail_body, "html.parser")

    if mail_html_bs4.body:
        mail_body = mail_html_bs4.body
        for t in mail_body.find_all("img"):
            t.decompose()
    else:
        # decodeしないで、改行タグをhtmlの<br>へ置き換え
        mail_body = b64decoded_mail_body.decode("utf8").replace("\r\n", "<br>")

    # jinja2埋込
    # テンプレート読み込み
    env = Environment(
        loader=FileSystemLoader(str((parent_dirpath / "prepare")), encoding="utf8")
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
    export_dirpath: Path,
    google_cred: Credentials,
) -> None:
    # 連絡項目の印刷用PDFファイル生成
    mimetype_xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    target_filepath = next(attachment_dirpath.glob("*MA-*.xlsx"))

    if not target_filepath:
        print("cant generate_pdf_byrenrakuexcel")
        return None

    # ExcelファイルをPDFに変換する
    media = MediaFileUpload(
        target_filepath,
        mimetype=mimetype_xlsx,
        resumable=True,
    )

    file_metadata = {
        "name": target_filepath.name,
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents": gsheet_tmmp_dir_ids,
    }

    try:
        drive_service = build("drive", "v3", credentials=google_cred)

        # Excelファイルのアップロードを行って、そのアップロードファイルをPDFで保存できるかチェック
        upload_results = (
            drive_service.files().create(body=file_metadata, media_body=media).execute()
        )
        print(upload_results)

        # pdfファイルを取りに行ってみる
        dl_request = drive_service.files().export_media(
            fileId=upload_results.get("id"), mimeType="application/pdf"
        )
        file = io.BytesIO()
        downloader = MediaIoBaseDownload(file, dl_request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print(f"Download {int(status.progress() * 100)}.")

        with (attachment_dirpath / Path("./連絡項目印刷用ファイル.pdf")).open(
            "wb"
        ) as export_exceltopdf:
            export_exceltopdf.write(file.getvalue())

        # post-porcess: Googleドキュメントに一時保持した配管連絡項目を除去する
        delete_tmp_excel_result = (
            drive_service.files()
            .delete(fileId=upload_results.get("id"), fields="id")
            .execute()
        )
        print("deleted file: ", delete_tmp_excel_result)

    except HttpError as error:
        # TODO:2022-12-09 エラーハンドリングは基本行わずここで落とすこと
        print(f"An error occurred: {error}")


def generate_projectdir(attachment_dirpath: Path, export_dirpath: Path) -> None:
    export_project_dir = export_dirpath / "proj_dir"
    export_project_dir.mkdir(exist_ok=True)

    # ファイルパスから型式を取り出す
    target_filepath = next(attachment_dirpath.glob("*MA-*.xlsx"))
    if not target_filepath:
        print("cant create projectdir")
        return None

    msm_katasiki_num = pick_msm_katasiki_by_renrakukoumoku_filename(target_filepath)

    # ボイラープレートからディレクトリ生成
    boilerplate_config = {
        "project_name": msm_katasiki_num,
    }
    generated_copier_worker = copier.run_copy(
        msm_gas_boilerplate_url,
        export_project_dir,
        data=boilerplate_config,
    )

    # 添付ファイルを解凍する
    for attachment_zfile in attachment_dirpath.glob("*.zip"):
        print(attachment_zfile)
        extract_zip.extract_zipfile(attachment_zfile, attachment_dirpath)

    # 添付ファイルをコピーする
    # TODO:2022-12-16 プロジェクトフォルダの名称は環境変数化したほうがいいかも？
    project_dir = generated_copier_worker.dst_path / f"ミスミ配管図MA-{msm_katasiki_num}納期 -"
    shutil.copytree(attachment_dirpath, (project_dir / "資料"), dirs_exist_ok=True)


def add_schedule_spreadsheet(
    attachment_dirpath: Path, google_cred: Credentials, nyukin_nextmonth: bool = False
) -> None:
    target_filepath = next(attachment_dirpath.glob("*MA-*.xlsx"))
    if not target_filepath:
        print("cant add schedule")
        return None

    msm_katasiki_num = pick_msm_katasiki_by_renrakukoumoku_filename(target_filepath)

    # エンドユーザー: 今は列がないので登録しない
    renrakukoumoku_range_enduser = "D10"
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

    sheet_service = build("sheets", "v4", credentials=google_cred)

    # スケジュール表の一番後ろの行へ追加する
    schedule_gsheet = (
        sheet_service.spreadsheets()
        .values()
        .append(
            spreadsheetId=schedule_sheet_id,
            range=table_search_range,
            body={"values": append_values},
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
        )
    ).execute()

    print(f"schedule updated -> {schedule_gsheet.get('updates')}")


def generate_estimate_calcsheet(
    attachment_dirpath: Path, google_cred: Credentials
) -> None:
    # 見積書生成機能をここで動かす。別のライブラリ化しておいて、それをここで呼び出すで良いと思う。
    # TODO:2022-12-15 ここは情緒なんだけど外に出す必要性があんまりないので今はそのまま。

    target_filepath = next(attachment_dirpath.glob("*MA-*.xlsx"))
    if not target_filepath:
        print("cant add schedule")
        return None

    msm_katasiki_num = pick_msm_katasiki_by_renrakukoumoku_filename(target_filepath)

    try:
        drive_service = build("drive", "v3", credentials=google_cred)

        copy_template_results = (
            drive_service.files()
            .copy(fileId=estimate_template_gsheet_id, fields="id,name")
            .execute()
        )

        # コピー先のファイル名を変更する
        template_suffix = "[ミスミ型番] のコピー"
        renamed_filename = copy_template_results.get("name").replace(
            template_suffix, msm_katasiki_num
        )
        rename_body = {"name": renamed_filename}
        rename_estimate_gsheet_result = (
            drive_service.files()
            .update(
                fileId=copy_template_results.get("id"),
                body=rename_body,
                fields="name",
            )
            .execute()
        )

        print(
            f"estimate copy and renamed -> {rename_estimate_gsheet_result.get('name')}"
        )
    except HttpError as error:
        # TODO:2022-12-09 エラーハンドリングは基本行わずここで落とすこと
        print(f"An error occurred: {error}")


def copy_projectdir(export_path: Path) -> None:
    # プロジェクトフォルダのパスを用意する
    export_prij_path = next((export_path / "proj_dir").glob("ミスミ配管図*"))
    print(export_prij_path)
    print(copy_project_dir_dest_path)
    shutil.copytree(
        export_prij_path,
        copy_project_dir_dest_path / export_prij_path.name,
        dirs_exist_ok=True,
    )
    pass


def main() -> None:
    print("[Start Process...]")

    google_cred: Credentials = googleapi.get_cledential(GOOGLE_API_SCOPES)

    messages: list[ExpandedMessageItem] = []
    try:
        # Call the Gmail API
        service = build("gmail", "v1", credentials=google_cred)

        # スレッド検索
        thread_results = (
            service.users()
            .threads()
            .list(userId=target_userid, q="label:snd-ミスミ (*MA-*)")
            .execute()
        )
        threads = thread_results.get("threads", [])

        # 上位10件のスレッド -> メッセージを取得。
        # スレッドに紐づきが2件ぐらいのメッセージの部分でのもので十分かな
        if threads:
            top_threads = list(itertools.islice(threads, 0, 10))
            for thread in top_threads:
                # threadsのid = threadsの一番最初のmessage>idなので、そのまま使う
                message_id = thread.get("id", "")
                # print(message_id)

                # スレッドの数が2以上 = すでに納品済みと思われるので削る。
                # TODO:2022-12-09 ここは2件以上でもまだやり取り中だったりする場合もあるので悩ましい
                # （数見るだけでもいいかもしれない
                thread_result = (
                    service.users()
                    .threads()
                    .get(userId=target_userid, id=message_id, fields="messages")
                ).execute()

                if len(thread_result.get("messages")) <= 2:
                    # スレッドの一番先頭にあるメッセージを取得する
                    messages.append(
                        ExpandedMessageItem(
                            gmail_message=thread_result.get("messages")[0]
                        )
                    )

    except HttpError as error:
        # TODO:2022-12-09 エラーハンドリングは基本行わずここで落とすこと
        print(f"An error occurred: {error}")
        exit()

    if not messages:
        print(f"cant find messages...")
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

    selected_message: ExpandedMessageItem = questionary.select(
        "メールの選択をしてください", choices=message_item_and_labels
    ).ask()

    # その他質問を確認
    ask_generate_projectfile = questionary.confirm(
        "プロジェクトファイルを生成しますか？(修正案件の場合は作成しないこと 例: MA-0000-1)", True
    ).ask()

    ask_add_schedule_and_generate_estimate_calcsheet = questionary.confirm(
        "スケジュール表追加と見積計算表の作成を行いますか？（プロジェクトファイル再作成時の対応に利用します）", True
    ).ask()

    # スキップする場合、↑の質問がFalseになる場合
    ask_add_schedule_nextmonth = (
        questionary.confirm("スケジュール表追加時に入金日を予定月の来月にしますか？", False)
        .skip_if(
            ask_add_schedule_and_generate_estimate_calcsheet == False, default=False
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
    # メール本文にimgファイルがある場合はそれを取り出す
    # multipart/relatedの時にあるので、それを狙い撃ちで取る

    if selected_message.body_related:
        message_imgs = [
            i
            for i in selected_message.body_related.get("parts")
            if "image" in i.get("mimeType")
        ]
        for msg_img in message_imgs:
            save_attachment_file(
                service,
                msg_img.get("filename"),
                selected_message.id,
                msg_img.get("body").get("attachmentId"),
            )
    # 添付ファイルの保持
    message_attachmentfiles = [
        i
        for i in selected_message.payload.get("parts")
        if "application" in i.get("mimeType")
    ]

    for msg_attach in message_attachmentfiles:
        save_attachment_file(
            service,
            msg_attach.get("filename"),
            selected_message.id,
            msg_attach.get("body").get("attachmentId"),
        )

    # 各種機能を呼び出す

    print("[Generate Mail Printable PDF]")
    generate_mail_printhtml(selected_message, attachment_dirpath)

    print("[Generate Excel Printable PDF]")
    generate_pdf_by_renrakukoumoku_excel(
        attachment_dirpath, export_dirpath, google_cred
    )

    if ask_generate_projectfile:
        print("[Generate template dirs]")
        generate_projectdir(attachment_dirpath, export_dirpath)
        print("[copy project dir]")
        copy_projectdir(export_dirpath)
    else:
        print("[Not Generate template dirs]")

    # TODO:2022-12-23 この部分をオフにする質問を追加する
    if ask_add_schedule_and_generate_estimate_calcsheet:
        print("[append schedule]")
        add_schedule_spreadsheet(
            attachment_dirpath, google_cred, ask_add_schedule_nextmonth
        )

        print("[add estimate calcsheet]")
        generate_estimate_calcsheet(attachment_dirpath, google_cred)
    else:
        print("[Not Add Scuedule, Generate estimate calcsheet]")

    print("[End Process...]")
    exit()


if __name__ == "__main__":
    main()
