# harbors_gas_api

## Set up
`yarn global add @google/clasp`

`yarn install`

`touch .clasp.json`

and write this 
```
{"scriptId":"XXXXXXXXXXXXXXXXXXXX"}
```

`clasp login`

## Google Apps Script APIの設定がONになっていることを確認する

https://script.google.com/u/1/home/usersettings

## deploy
clasp push

## スクリプトプロパティ設定
スクリプトプロパティを設定しています。環境に合わせ設定が必要となります。
（オーナー権限が必要）

`WEBHOOK_URL`
 Slackの通知先

`CALENDAR_CONTACT_ID`
 予約を追加するカレンダーID

`RESERVE_CONFIRMATION_TEMPLATE`
 予約完了時のメール送信テンプレートID
 

