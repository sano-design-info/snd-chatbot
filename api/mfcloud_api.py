import json
import re
import wsgiref.simple_server
import wsgiref.util
from dataclasses import dataclass
from pathlib import Path

from authlib.integrations.requests_client import OAuth2Session

from helper import EXPORTDIR_PATH, load_config

# load config, credential
config = load_config.CONFIG

# v3の各URLはv2から変更されているので置き換えている
AUTHORZATION_BASE_URL = "https://api.biz.moneyforward.com/authorize"
TOKEN_URL = "https://api.biz.moneyforward.com/token"
API_ENDPOINT = "https://invoice.moneyforward.com/api/v3/"

REDIRECT_URI = "http://localhost:8080"
SCOPE = "mfc/invoice/data.write"
SAVE_TOKEN_PATH = EXPORTDIR_PATH / "mfcloud_access_token.json"

# postで使うヘッダー
post_headers = {"content-type": "application/json", "accept": "application/json"}


# get redirect code local server wsgi app
# ref: https://github.com/googleapis/google-auth-library-python-oauthlib/blob/3c829e87dd7720ddc1e70431927072e612170c98/google_auth_oauthlib/flow.py
# Licence: Apache v2
class RedirectWsgiApp(object):
    """WSGI app to handle the authorization redirect.
    Stores the request URI and displays the given success message.
    """

    def __init__(self, success_message):
        """
        Args:
            success_message (str): The message to display in the web browser
                the authorization flow is complete.
        """
        self.last_request_uri = None
        self._success_message = success_message

    def __call__(self, environ, start_response):
        """WSGI Callable.
        Args:
            environ (Mapping[str, Any]): The WSGI environment.
            start_response (Callable[str, list]): The WSGI start_response
                callable.
        Returns:
            Iterable[bytes]: The response body.
        """
        start_response("200 OK", [("Content-type", "text/plain; charset=utf-8")])
        self.last_request_uri = wsgiref.util.request_uri(environ)
        return [self._success_message.encode("utf-8")]


# local wsgi server -> request redirect url
def get_auth_response_by_localserver(
    success_message: str,
    bind_addr=None,
    host="localhost",
    port=8080,
    redirect_uri_trailing_slash=True,
    timeout_seconds=None,
) -> str:
    """認証時にlocalhostを起動して、その後、
    リダイレクトURLのレスポンスを得たらコードを取り出し終了する"""

    redirect_uri = ""
    wsgi_app = RedirectWsgiApp(success_message)

    wsgiref.simple_server.WSGIServer.allow_reuse_address = False
    local_server = wsgiref.simple_server.make_server(bind_addr or host, port, wsgi_app)

    redirect_uri_format = (
        "http://{}:{}/" if redirect_uri_trailing_slash else "http://{}:{}"
    )
    redirect_uri = redirect_uri_format.format(host, local_server.server_port)

    print(f"set redirect uri -> :{redirect_uri}")
    print("wait response...")

    local_server.timeout = timeout_seconds
    local_server.handle_request()

    # Note: using https here because oauthlib is very picky that
    # OAuth 2.0 should only occur over https.
    authorization_response = wsgi_app.last_request_uri.replace("http", "https")

    # This closes the socket
    local_server.server_close()

    return authorization_response


@dataclass
class MFCIClient:
    client_id: str = config.get("mfci").get("MFCLOUD_CLIENT_ID")
    client_secret: str = config.get("mfci").get("MFCLOUD_CLIENT_SECRET")
    token = None

    def _load_token(self) -> dict:
        """jsonで保存されているアクセストークンを取得する。"""
        try:
            with SAVE_TOKEN_PATH.open("r") as access_token_cache:
                token = json.load(access_token_cache)
        except IOError:
            return None
        return token

    def _save_token(self, token: dict, **kwargs):
        """アクセストークンをjsonで保存する。"""

        # tokenの中身をそのまま保存する
        # print(f"save taisyo token:{token}")
        with SAVE_TOKEN_PATH.open("w") as access_token_cache:
            json.dump(token, access_token_cache)

    def _get_new_oauth_session(self):
        """MFクラウドのoAuthのセッション取得処理"""

        session = OAuth2Session(
            client_id=self.client_id,
            client_secret=self.client_secret,
            scope=SCOPE,
            redirect_uri=REDIRECT_URI,
        )
        authorization_url, _ = session.create_authorization_url(
            AUTHORZATION_BASE_URL,
        )

        print("Please go here and authorize,", authorization_url)

        # ここでローカルサーバーを作って、リダイレクトされたUrl->認証レスポンスを取り出す
        # docker動作を考慮して、bind_addrを設定
        redirect_port = int(re.sub(r"https?://.*:", "", REDIRECT_URI))
        authorization_response = get_auth_response_by_localserver(
            success_message="get code!! close browser",
            bind_addr="0.0.0.0",
            port=redirect_port,
        )

        return session, authorization_response

    def get_session(self) -> OAuth2Session:
        self.token = self._load_token()

        # print(f"now token:{self.token}")
        if self.token:
            # 現在のトークンを使ってセッションを作成
            # 自動的にリフレッシュトークンも使いトークンの更新も行う
            session = OAuth2Session(
                self.client_id,
                self.client_secret,
                token=self.token,
                token_endpoint=TOKEN_URL,
                update_token=self._save_token,
            )

            # print(f"load/renew token:{session.token}")
        else:
            session, redirect_response = self._get_new_oauth_session()

            # アクセストークンを使ってfetchする
            self.token = session.fetch_token(
                url=TOKEN_URL,
                authorization_response=redirect_response,
            )

            # print(f"new token:{self.token}")
            # 取得したtokenをjsonで保存
            self._save_token(self.token)

        return session


## ここからはAPIのリクエストを行う関数群

### 品目操作用の関数群
# 見積書、請求書共に品目を使うので、共通で使う


def create_item(mfci_session: OAuth2Session, item_data: dict) -> dict:
    """品目を作成する"""
    created_item_res = mfci_session.post(
        f"{API_ENDPOINT}items",
        data=json.dumps(item_data),
        headers=post_headers,
    )

    return json.loads(created_item_res.content)


### 見積書操作用の関数群


def get_quotes(mfci_session: OAuth2Session, per_page_length: int) -> dict:
    """最近の見積一覧を取得する。"""
    quote_res = mfci_session.get(
        f"{API_ENDPOINT}quotes/?page=1&per_page={per_page_length}"
    )

    return json.loads(quote_res.content)


def create_quote(mfci_session: OAuth2Session, quote_data: dict) -> dict:
    """見積作成用のAPIリクエストを行いレスポンスを取得する
    MFCloudのoAuth2セッションとAPIリクエストに必要になるデータを受け取る。
    結果はAPIレスポンスで得られるjsonデータの辞書形式のダンプ"""

    created_quote_res = mfci_session.post(
        f"{API_ENDPOINT}quotes",
        data=json.dumps(quote_data),
        headers=post_headers,
    )

    return json.loads(created_quote_res.content)


def attach_item_into_quote(
    mfci_session: OAuth2Session, quote_id: str, item_id: str
) -> bool:
    """見積書に品目を追加する"""
    attach_item_res = mfci_session.post(
        f"{API_ENDPOINT}quotes/{quote_id}/items",
        data=json.dumps({"item_id": item_id}),
        headers=post_headers,
    )

    # TODO: 2023-08-24 ここはステータスコード複数があるので、それぞれ何を返すかを検討する
    if attach_item_res.status_code == 201:
        return True
    else:
        return False


# TODO: 2023-08-29 ここは見積書請求書どちらも同じものが使えるが便宜上分けている。後で統合する
def download_quote_pdf(
    mfci_session: OAuth2Session, quote_url: str, save_filepath: Path
) -> None:
    """
    quote pdfのURLのPDFデータをDLして、ファイル保存をする。
    """
    print("PDFのURL: ", quote_url)

    dl_quote_pdf_res = mfci_session.get(quote_url)

    with save_filepath.open("wb") as result_pdf:
        result_pdf.write(dl_quote_pdf_res.content)


### 請求書操作用の関数群


def get_billings(mfci_session: OAuth2Session, per_page_length: int) -> dict:
    """最近の見積一覧を取得する。"""
    quote_res = mfci_session.get(
        f"{API_ENDPOINT}billings/?page=1&per_page={per_page_length}"
    )

    return json.loads(quote_res.content)


def create_billing(mfci_session: OAuth2Session, billing_data: dict) -> dict:
    """請求書を生成する"""
    created_billing_res = mfci_session.post(
        f"{API_ENDPOINT}billings",
        data=json.dumps(billing_data),
        headers=post_headers,
    )

    return json.loads(created_billing_res.content)


def create_invoice_template_billing(
    mfci_session: OAuth2Session, billing_data: dict
) -> dict:
    """請求書を生成する。2023年10月から始まるインボイス制度向けの請求書作成用"""
    created_billing_res = mfci_session.post(
        f"{API_ENDPOINT}invoice_template_billings",
        data=json.dumps(billing_data),
        headers=post_headers,
    )

    return json.loads(created_billing_res.content)


def attach_billingitem_into_billing(
    mfci_session: OAuth2Session, quote_id: str, item_id: str
) -> bool:
    """請求書に品目を追加する"""
    attach_item_res = mfci_session.post(
        f"{API_ENDPOINT}billings/{quote_id}/items",
        data=json.dumps({"item_id": item_id}),
        headers=post_headers,
    )

    # TODO: 2023-08-24 ここはステータスコード複数があるので、それぞれ何を返すかを検討する
    if attach_item_res.status_code == 201:
        return True
    else:
        return False


# TODO: 2023-08-29 ここは見積書請求書どちらも同じものが使えるが便宜上分けている。後で統合する
def download_billing_pdf(
    mfci_session: OAuth2Session, billing_pdf_url: str, save_filepath: Path
) -> None:
    """
    請求書 pdfのURLのPDFデータをDLして、ファイル保存をする。
    """
    print("PDFのURL: ", billing_pdf_url)

    dl_quote_pdf_res = mfci_session.get(billing_pdf_url)

    with save_filepath.open("wb") as result_pdf:
        result_pdf.write(dl_quote_pdf_res.content)
