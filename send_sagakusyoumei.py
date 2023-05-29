# 必要なライブラリをインポート
import time
import toml
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from pathlib import Path
from datetime import datetime

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# 認証情報をtomlファイルから読み込む
config = toml.load("config.toml")
MSM_PROSUGATE_ID = config["send_sagakusyoumei"]["MSM_PROSUGATE_ID"]
MSM_PROSUGATE_PASS = config["send_sagakusyoumei"]["MSM_PROSUGATE_PASS"]

# 納品月を取得
delivery_month = datetime.now().strftime("%Y/%m")
# delivery_month = "2023/06"

# ダウンロード先フォルダを指定
DOWNLOAD_DIR = str(Path.home() / "Downloads" / "misumi")


def main():
    # headless modeを設定
    webdriver_options = Options()
    webdriver_options.add_argument("--headless=new")
    # ダウンロード先フォルダを指定
    webdriver_options.add_experimental_option(
        "prefs", {"download.default_directory": DOWNLOAD_DIR}
    )

    # ChromeのWebDriverオブジェクトを作成
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=webdriver_options
    )
    driver.implicitly_wait(5)

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
        print("ファイルダウンロード開始！")

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
            print("File downloaded successfully!")

            # ダウンロードしたファイルのパスを見つけて出力
            download_dir = Path(DOWNLOAD_DIR)
            latest_file = max(download_dir.glob("*"), key=lambda x: x.stat().st_mtime)
            print("Downloaded file path: ", latest_file)

        elif 'ng-reflect-errors-list="出力結果が0件です。"' in notice_component.get_attribute(
            "outerHTML"
        ):
            print("File download failed!")
    except Exception as e:
        print(e)
    finally:
        # ブラウザを閉じる
        driver.quit()

    # 2.差額証明書を作成
    # MFクラウドの請求書情報を取得して、Google Sheet APIを使って、差額証明書のテンプレートに必要な情報を入れて、PDFファイルとしてダウンロードする
    TEMPLATE_SHEET_ID = "TEMPLATE_SHEET_ID"

    # 情報記入先: セルの指定
    # 買掛金金額: F6:H7

    # 3.メールの下書きを作成
