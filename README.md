# snd-chatbot

佐野設計事務所のバックオフィスなどで使う機能を呼び出すチャットボット

TODO:2024-02-01 ここの内容は書き換えて、動作時に必要な説明を載せます。テストでも必要な操作を上げていきます

## prepare/準備

### config

config.tomlが本番用です。 dev-config.tomlは開発向けにしてます。 .env で`DEBUG=True`を入れると開発向けとして機能しています。
2023-04-17 実質設定値のみの変更しか考慮されていないです。

### .env
実質利用していませんが、DEBUGフラグがあります

### rcloneでボリュームマウント

docker composeのボリューム機能で rclone + Googleドライブ共有ドライブのマウントを行ってます。ミスミの作業用テンプレート（copierのボイラープレート）を呼び出して、プロジェクトファイルを展開したり。google api向けの認証ファイルや一時的なファイル保持の場所として使います

configはなにもない場合はホスト端末などrcloneが入っている環境で `rclone config` を行い、Googleドライブの共有ドライブを呼び出します。
compose.ymlにあるvolumeの意味は次の通り。

* snd-sync-dir
  * 社内の作業用共有フォルダと伝わる作業の同期用フォルダ。
  * Googleドライブで運用していて、こちらにあるボイラープレートを呼び出してプロジェクトファイル生成 -> 作業者へ展開を行う
* exportdir
  * chatbotの認証ファイルとexportdirとしてコード上で利用している一時作業用の環境
  * TODO: 2025-01-15 こちらは、認証ファイルとトークン以外はtmpフォルダ（メモリ上）での展開として、ここではつくらないような仕組みにしていきます。

## チャットボットの各種機能

TODO: 2025-01-15  各種機能の意味を乗せる。 run_mail_action, generaete_quote, generate_invoice
### run mail action/メールから各種ファイルとスプレッドシート登録スクリプト

やってくれること

* 印刷物用ファイルの生成: メール本文html, 配管連絡項目excel -> PDF化
* プロジェクトフォルダとファイルの生成。ボイラープレート経由で行います
  <!-- * ボイラープレート: <http://192.168.35.52:3000/snd-private/misumi-gas-boilerplate> -->
* 見積計算表の案件番号ファイル名でコピー生成
* スケジュール表へ登録

```cmd
> pipenv run python .\run_mail_action.py
```

## デプロイ方法

* docker compose build
* docker compose up -d
* google oauth認証を行う（TODO: 2025-01-15 現在のgoogle apiの認証だと、ブラウザを呼び出す方法なので、urlを表示するのみにする。）
  * python script_〇〇.py を実行して、googleのトークンを作成
* 動作確認は、google chatで calc_add を実行する。計算ができたらOK

## テスト方法

実際のメールやGoogleスプレッドシートなどでテストをする場合は以下の手順を踏みます。

TODO: 2025-01-15 現在こちらの方法ではテスト不可能なので、新たにテストを行う方法論を検討します。
TODO:2024-02-01 クリーンアップ用のスクリプトを用意することを検討すること。
TODO:2024-02-01 以下の手順を自動化する手段を検討すること。

* docker-compose -f ./compose.localdev.yml build -> docker-compose -f ./compose.localdev.yml up -d でサーバーを起動します
* メールはテスト用のメールを用意します
  * MA-9991というテスト用メールを作成します。リッチテキストとしって、配管連絡項目を用意します。
  * ファイルが古いと検索に出てこないので、あらかじめ作成しなおします
  * ラベル:`snd-ミスミ` を付ける必要があります（snd>ミスミという親子のラベル）
* 見積計算表を所定の位置に置きます。
  * 2024-02-01 時点で　`ミスミ配管図見積り計算表v3_MA-9991`という名称で用意済みです。
* script_**.pyの各種を実行してcli上で動作確認をします
* 実行結果は生成物や動作挙動を確認します。docker compose logs workerでログを確認します
* 最後にdocker-compose -f ./compose.localdev.yml downでサーバーを停止します
