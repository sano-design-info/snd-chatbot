import json
import pickle
from dataclasses import dataclass
from pathlib import Path

from requests_oauthlib import OAuth2Session

from helper import EXPORTDIR_PATH, load_config

# load config, credential

config = load_config.CONFIG

AUTHORZATION_BASE_URL = "https://invoice.moneyforward.com/oauth/authorize"
TOKEN_URL = "https://invoice.moneyforward.com/oauth/token"
API_ENDPOINT = "https://invoice.moneyforward.com/api/v2"

# TODO:2022-04-05 token更新時にアクセスするローカルサーバーをつくってハンドリングできないかかんがえてみる
# TODO:2023-03-28 上のTodoはAPI v3から対応する。


@dataclass
class MFCICledential:
    client_id: str = config.get("mfci").get("MFCLOUD_CLIENT_ID")
    client_secret: str = config.get("mfci").get("MFCLOUD_CLIENT_SECRET")
    _token = None

    def _load_token(self):
        """トークンを取得する。"""
        try:
            with (EXPORTDIR_PATH / "access_token.dat").open("rb") as access_token_cache:
                token = pickle.load(access_token_cache)
        except IOError:
            return None
        return token

    def _get_oauth_session(self):
        """MFクラウドのoAuthのセッション取得処理を実行する"""
        oauth2_session = OAuth2Session(
            client_id=self.client_id,
            scope="write",
            redirect_uri="https://example.com",
            state=None,
        )
        authorization_url, state = oauth2_session.authorization_url(
            AUTHORZATION_BASE_URL
        )

        print("Please go here and authorize,", authorization_url)
        redirect_response = input("Paste the full redirect URL here:")

        return oauth2_session, redirect_response

    def _save_token(self, token):
        # 取得したtokenをpickleで直列化
        with (EXPORTDIR_PATH / "access_token.dat").open("wb") as access_token_cache:
            pickle.dump(token, access_token_cache)

    def get_session(self) -> OAuth2Session:
        self._token = self._load_token()

        if self._token:
            # リフレッシュトークンで自動的に更新しつつトークンを利用する
            extra = {"client_id": self.client_id, "client_secret": self.client_secret}
            mfcloud_invoice_session = OAuth2Session(
                self.client_id,
                token=self._token,
                auto_refresh_url=TOKEN_URL,
                auto_refresh_kwargs=extra,
                token_updater=self._save_token,
            )

        else:
            mfcloud_invoice_session, redirect_response = self._get_oauth_session()

            # アクセストークンを使ってfetchする
            self._token = mfcloud_invoice_session.fetch_token(
                TOKEN_URL,
                client_secret=self.client_secret,
                authorization_response=redirect_response,
            )
            print(self._token)
            # 取得したtokenをpickleで直列化
            self._save_token(self._token)

        return mfcloud_invoice_session


def generate_quote(mfcloud_invoice_session: OAuth2Session, quote_data: dict) -> dict:
    """見積作成用のAPIリクエストを行いレスポンスを取得する
    MFCloudのoAuth2セッションとAPIリクエストに必要になるデータを受け取る。
    結果はAPIレスポンスで得られるjsonデータの辞書形式のダンプ"""
    headers = {"content-type": "application/json", "accept": "application/json"}

    generated_quote_res = mfcloud_invoice_session.post(
        f"{API_ENDPOINT}/quotes?excise_type=boolean",
        data=json.dumps(quote_data),
        headers=headers,
    )

    return json.loads(generated_quote_res.content)


def download_quote_pdf(
    mfcloud_invoice_session: OAuth2Session, quote_url: str, save_filepath: Path
) -> None:
    """quote pdfのURLのPDFデータをDLして、ファイル保存をする。戻り値はファイル保存先のパス"""
    print("PDFのURL: ", quote_url)

    dl_quote_pdf_res = mfcloud_invoice_session.get(quote_url)

    with save_filepath.open("wb") as result_pdf:
        result_pdf.write(dl_quote_pdf_res.content)

    print(f"ファイルが生成されました。保存先:{save_filepath}")


def get_quote_list(
    mfcloud_invoice_session: OAuth2Session, per_page_length: int
) -> dict:
    """最近の見積一覧を取得する。"""
    quote_res = mfcloud_invoice_session.get(
        f"{API_ENDPOINT}/quotes/?page=1&per_page={per_page_length}&excise_type=boolean"
    )

    return json.loads(quote_res.content)


# 請求書一覧を取得する
def get_billing_list(
    mfcloud_invoice_session: OAuth2Session,
    per_page_length: int,
) -> dict:
    # TODO:2023-05-30 ここでは請求書一覧となるから、フィルターをかける必要がある。
    """最近の見積一覧を取得する。"""
    quote_res = mfcloud_invoice_session.get(
        f"{API_ENDPOINT}/billings/?page=1&per_page={per_page_length}&excise_type=boolean"
    )

    return json.loads(quote_res.content)


# TODO:2022-11-24 ここはAPIの扱いについて問い合わせ中なので、その結果を見て実装する
# TODO:2023-03-28 ここは残念ながら対応できないので利用しない。issueを作って除去する
def search_quote_list(
    mfcloud_invoice_session: OAuth2Session,
    per_page_length: int,
    range: dict,
    q: str = "",
) -> dict:
    """検索パラメータを使って見積一覧を検索し取得する。"""

    # TODO:2022-11-24 現在searchのapiでパラメータを渡してもうまく動いてくれないので、通常の検索結果 -> フィルターをしたバージョンにする

    quote_result = get_quote_list(mfcloud_invoice_session, per_page_length)

    # 取引先でフィルター
    result_data = (
        i for i in quote_result["data"] if i["attributes"]["department_id"] == q
    )

    # 日付でフィルター。create_atを見てfromより前になるまでページ番号をひたすら再起処理する

    return []

    # # パラメーターをgetメソッドに渡して渡せるようにする
    # param = {"page": 1, "per_page": {per_page_length}, "excise_type": "boolean"}
    # param |= range
    # # TODO:2022-11-24 APIのqの扱いがわかったらパラメーターに追加する

    # quote_res = mfcloud_invoice_session.get(
    #     f"{API_ENDPOINT}quotes/search/",
    #     params=param,
    # )
    # return json.loads(quote_res.content)
