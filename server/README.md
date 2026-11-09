# 棚卸API サーバー(PostgreSQL保存)

スマホ/PCのアプリ → このAPIサーバー → PostgreSQL(172.18.61.232 / test_db)の流れで記録と写真を保存します。
DBのパスワードはこのサーバーの `.env` にだけ置き、GitHubには載せません。

## 起動手順(社内のPC/サーバー1台で。Node.js 18以上)
1. `server` フォルダで `npm install`
2. `.env.example` を `.env` にコピーし、`PGPASSWORD` と `API_KEY`(合言葉)を入力
3. `npm start`(初回に `records` と `photos` テーブルを自動作成します)
4. アプリの「設定・出力」→「保存先サーバー」に、サーバーのアドレスと合言葉を入力して「接続テスト」

## HTTPSについて(重要)
スマホのカメラと、GitHub PagesからのAPI呼び出しは、どちらもHTTPSが必要です。
- 社内の証明書(または自己署名証明書)を `.env` の `TLS_CERT` / `TLS_KEY` に指定してください。
- `SERVE_APP=1` にすると、このサーバーでアプリ画面も配信します。スマホでは `https://サーバー:8443/` を開いて使うのが一番簡単です(GitHub Pagesとの通信の制限を受けません)。
- 自己署名証明書は、各スマホに一度だけ証明書をインストール(信頼)する必要があります。

## 保存される内容
- `records`: 資産コード・現物・方法・日時・緯度経度・精度・写真名
- `photos`: 写真本体(bytea)。名前は `資産コード_機番.jpg` / `資産コード_連番.jpg`

## API
`X-API-Key` ヘッダーに合言葉が必要です。
`GET /api/health` / `GET /api/records` / `PUT|DELETE /api/records/:code` / `GET /api/photos` / `PUT|GET|DELETE /api/photos/:name`
