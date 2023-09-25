# 必要なライブラリをインポート
import re
import sys
import time
from datetime import datetime

from dateutil import relativedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

from api import googleapi, mfcloud_api
from helper import EXPORTDIR_PATH, extract_compressfile, load_config

# 認証情報をtomlファイルから読み込む
config = load_config.CONFIG
MSM_TORIHIKISAKI_ID = config.get("mfci").get("TORIHIKISAKI_ID")
MSM_PROSUGATE_ID = config["send_sagakusyoumei"]["MSM_PROSUGATE_ID"]
MSM_PROSUGATE_PASS = config["send_sagakusyoumei"]["MSM_PROSUGATE_PASS"]
TEMPLATE_SHEET_ID = config["send_sagakusyoumei"]["TEMPLATE_SHEET_ID"]
MAIL_TO = config["send_sagakusyoumei"]["MAIL_TO"]
MAIL_CC = config["send_sagakusyoumei"]["MAIL_CC"]
MAIL_TEMPLATE_TITLE = config["send_sagakusyoumei"]["MAIL_TEMPLATE_TITLE"]
MAIL_TEMPLATE_BODY = config["send_sagakusyoumei"]["MAIL_TEMPLATE_BODY"]


# 時刻の設定
today_datetime = datetime.today()
today_dt_minus_one_month = today_datetime - relativedelta.relativedelta(months=1)
delivery_month = today_dt_minus_one_month.strftime("%Y/%m")  # 例: "2023/06"

today_dt_minus_one_month__year = today_dt_minus_one_month.strftime("%Y")
today_dt_minus_one_month__month = today_dt_minus_one_month.strftime("%m")

# ダウンロード先フォルダを指定
DOWNLOAD_DIR = EXPORTDIR_PATH / "Downloads" / "misumi_prosgate_purchase_list"
EXPORT_DIR = EXPORTDIR_PATH / "billing"

DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
EXPORT_DIR.mkdir(parents=True, exist_ok=True)

# 差額証明のファイル名を定義
sagakusyoumei_new_name = (
    f"ミスミ売掛金勘定差額明細書_{today_dt_minus_one_month__year}{today_dt_minus_one_month__month}"
)


def extract_total_amount(text: str) -> int | None:
    """
    購入リストのテキストファイルから、買掛金金額を抽出する。結果はint

    Args:
        text (str): 購入リストのテキストファイル
    Returns:
        int | None: 買掛金金額
    """
    match = re.search("TOTAL AMOUNT\t(\d+)", text)
    if match:
        return int(match.group(1))
    else:
        return None


def main():
    # 各APIの認証
    mfci_session = mfcloud_api.MFCIClient().get_session()

    google_cred: Credentials = googleapi.get_cledential(googleapi.API_SCOPES)
    sheet_service = build("sheets", "v4", credentials=google_cred)
    drive_service = build("drive", "v3", credentials=google_cred)
    gmail_service = build("gmail", "v1", credentials=google_cred)

    # TODO:2023-09-14 関数にする
    # 1. 買掛金案内書から、msm側の購入リストの金額を取得
    # headless modeを設定
    webdriver_options = Options()
    webdriver_options.add_argument("--headless=new")
    webdriver_options.add_argument("--no-sandbox")
    webdriver_options.add_argument("--disable-dev-shm-usage")
    # ダウンロード先フォルダを指定
    webdriver_options.add_experimental_option(
        "prefs", {"download.default_directory": str(DOWNLOAD_DIR)}
    )

    # ChromeのWebDriverオブジェクトを作成
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=webdriver_options
    )
    driver.implicitly_wait(5)

    # seleniumでブラウザ操作。ログイン -> 買掛金案内書生成 -> ファイルダウンロード
    try:
        # https://prosugate.misumi-ec.com/ にアクセス
        driver.get("https://prosugate.misumi-ec.com/")

        # IDとパスワードを入力してログイン
        driver.find_element(By.XPATH, '//*[@id="userNameInput"]').send_keys(
            MSM_PROSUGATE_ID
        )
        driver.find_element(By.XPATH, '//*[@id="passwordInput"]').send_keys(
            MSM_PROSUGATE_PASS
        )
        driver.find_element(By.XPATH, '//*[@id="submitButton"]').click()

        # ログイン後のリダイレクトを検証
        WebDriverWait(driver, 10).until(
            EC.url_to_be("https://prosugate.misumi-ec.com/#/prosu/top/management")
        )

        redirected_element = driver.find_element(
            By.XPATH, '//*[@id="topAnchor"]/div/div[1]/div[1]'
        )
        print("Found redirected element: ", redirected_element)

        # 検収紹介にアクセス
        kaikake_url = "https://prosugate.misumi-ec.com/#/prosu/check/search"
        driver.get(kaikake_url)

        WebDriverWait(driver, 10).until(EC.url_to_be(kaikake_url))
        print("Redirected successfully!: " + kaikake_url)

        # TODO:2023-05-29 ここはクリックができるかの検証を入れたいが、入れても動かないので、time.sleep()で対応
        # 買掛案内書生成に切り替え
        time.sleep(5)
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/app-component/pronogate-component/fc0417sc01-component/main/div[2]/div[2]/div[2]/div[3]/button",
                )
            )
        ).click()

        # フォームに納品月を入力
        time.sleep(5)
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="misumiKensyuGetuId"]/div/div/input')
            )
        ).send_keys(delivery_month)

        # ファイル作成ボタンを押す
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "/html/body/app-component/pronogate-component/fc0417sc01-component/main/div[2]/div[2]/div[2]/button",
                )
            )
        ).click()
        print("ファイルダウンロード開始")

        time.sleep(5)

        # エラー処理
        notice_component_xpath = "/html/body/app-component/pronogate-component/fc0417sc01-component/noticearea-component"
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, notice_component_xpath))
        )
        notice_component = driver.find_element(By.XPATH, notice_component_xpath)
        if (
            'ng-reflect-info-message="買掛金御案内書を印刷/ダウンロードしました。"'
            in notice_component.get_attribute("outerHTML")
        ):
            # ダウンロードしたファイルのパスを見つけて出力
            latest_file = max(DOWNLOAD_DIR.glob("*"), key=lambda x: x.stat().st_mtime)
            print("ファイルダウンロード成功しました！: ", latest_file)

        elif 'ng-reflect-errors-list="出力結果が0件です。"' in notice_component.get_attribute(
            "outerHTML"
        ):
            print("ファイルダウンロードができませんでした。終了します。")
            sys.exit(1)
    except Exception as e:
        print(f"例外発生で終了します:{e}")
        sys.exit(1)
    finally:
        # ブラウザを閉じる
        driver.quit()

    # 生成されたファイルから、zip解凍 -> ダウンロードで展開されたフォルダからtextファイル取得
    extract_compressfile.extract_file(latest_file, DOWNLOAD_DIR)
    purchaselist_textfilepath = max(
        DOWNLOAD_DIR.glob("*.txt"), key=lambda x: x.stat().st_mtime
    )
    print("解凍できた購入リストのファイルパス: ", purchaselist_textfilepath)

    # 購入リストファイルから、買掛金金額を取得
    with purchaselist_textfilepath.open(encoding="utf_16_le") as purchaselist_text:
        msm_bougth_total_amount = extract_total_amount(purchaselist_text.read())

    # 2.請求書の金額を取得

    # 請求書一覧取得をして、取得するフィルターを掛けたうえで一か月前の請求書を取得、請求書情報から 請求金額（税抜き）を取得する
    billing_list = mfcloud_api.get_billings(mfci_session, 100)
    latest_billing = next(
        (
            billing
            for billing in billing_list["data"]
            if (create_at_dt := datetime.fromisoformat(billing["created_at"])).year
            == today_dt_minus_one_month.year
            and create_at_dt.month == today_dt_minus_one_month.month
            and billing["department_id"] == MSM_TORIHIKISAKI_ID
        )
    )
    latest_billing_price = int(float(latest_billing["subtotal_price"]))

    # 比較して同値ではないなら終了、同値なら3へすすむ
    if latest_billing_price != msm_bougth_total_amount:
        print("請求書の金額と購入リストの金額が一致しません。終了します。")
        sys.exit(1)

    # 3.売掛金勘定差額明細書を作成
    cellmap = {
        "F6:H7": latest_billing_price,
        "I6:J7": msm_bougth_total_amount,
        "M1:N1": today_dt_minus_one_month__year,
        "P1:Q1": today_dt_minus_one_month__month,
        "Y6:AA7": today_datetime.strftime("%Y/%m/%d"),
    }

    sheet_id = googleapi.dupulicate_file(
        drive_service, TEMPLATE_SHEET_ID, sagakusyoumei_new_name
    )
    new_sheet_name = "売掛金勘定差額明細書"

    # 値を記入します
    try:
        for cell, value in cellmap.items():
            sheet_service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=new_sheet_name + "!" + cell,
                valueInputOption="USER_ENTERED",
                body={"values": [[value]]},
            ).execute()
    except Exception as e:
        print(f"例外発生で終了します:{e}")
        sys.exit(1)

    # シートをPDFとしてエクスポートします
    googleapi.export_pdf_by_driveexporturl(
        google_cred.token, sheet_id, EXPORT_DIR / f"{sagakusyoumei_new_name}.pdf"
    )

    # 4.メールの下書きを作成
    mail_body = MAIL_TEMPLATE_BODY.replace(
        "{{billing_year}}", today_dt_minus_one_month__year
    ).replace("{{billing_month}}", today_dt_minus_one_month__month)
    googleapi.append_draft(
        gmail_service,
        MAIL_TO,
        MAIL_CC,
        MAIL_TEMPLATE_TITLE,
        mail_body,
        [EXPORT_DIR / f"{sagakusyoumei_new_name}.pdf"],
    )
    print("メールの下書きを作成しました")


if __name__ == "__main__":
    main()
