# msm-gas-prepare

## prepare

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
