# msm-gas-prepare

TODO:2024-02-01 ここの内容は書き換えて、動作時に必要な説明を載せます。テストでも必要な操作を上げていきます

## prepare

### config

config.tomlが本番用です。 dev-config.tomlは開発向けにしてます。 .env で`DEBUG=True`を入れると開発向けとして機能しています。
2023-04-17 実質設定値のみの変更しか考慮されていないです。

### delete old export file

プロジェクトフォルダのコピー元になるフォルダは実行時間事に生成して保持します（後で印刷し忘れとかを防ぐために）

必要なくなったらまとめて消去できます。以下のスクリプトファイルか`export_files`フォルダを消してください。

```cmd
> .\rmdir_export_files.ps1
```

## run mail action/メールから各種ファイルとスプレッドシート登録スクリプト

やってくれること

* 印刷物用ファイルの生成: メール本文html, 配管連絡項目excel -> PDF化
* プロジェクトフォルダとファイルの生成。ボイラープレート経由で行います
  * ボイラープレート: <http://192.168.35.52:3000/snd-private/misumi-gas-boilerplate>
* 見積計算表の案件番号ファイル名でコピー生成
* スケジュール表へ登録

```cmd
> pipenv run python .\run_mail_action.py
```

## テスト方法

実際のメールやGoogleスプレッドシートなどでテストをする場合は以下の手順を踏みます。

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
