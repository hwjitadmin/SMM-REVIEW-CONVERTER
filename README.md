# smm-review-converter

SMM 4年プラン 住宅用 案件審査ツール（[smm-review-tool.web.app](https://smm-review-tool.web.app)、リポジトリ `SMM-REVIEW-TOOL`）の
「審査書類一式ファイル出力」で使う、見積ファイル（Excel）→ PDF の変換サーバーです。Cloud Run で動きます。

## 構成

```
app/main.py      API（Flask + gunicorn）。POST /api/convert-sheets、GET /healthz
app/convert.py   LibreOffice（UNO）での変換。パターンA/Bの判定・対象シートだけ出力・パターンBの印刷設定の補正
Dockerfile       Debian + LibreOffice Calc + 日本語フォント（IPA / Noto CJK）
fonts.conf       MS ゴシック・メイリオ等 → IPA / Noto フォントへの対応付け
.github/workflows/deploy.yml  main に push すると Cloud Run（hanwha-japan-app / asia-northeast1）へ自動デプロイ
```

## API

| | |
|---|---|
| `POST /api/convert-sheets` | multipart の `file`（.xls / .xlsx / .xlsm、30MB まで） |
| 成功 | `200 application/pdf`、ヘッダー `X-Pattern`（A / B）、`X-Conversion-Warnings`（URL エンコードした JSON 配列） |
| 失敗 | `4xx / 5xx` `{"error": "メッセージ"}` |

画面からは Firebase Hosting の rewrite（`/api/**` → このサービス）経由で、同じオリジンとして呼ばれます。
この形を変えるときは、画面側（`SMM-REVIEW-TOOL` の `src/app.js` の `exportPackagePdf()`）も合わせて直してください。

## 見積ファイルの2つのパターン

| パターン | 出力するシート | 印刷設定 |
|---|---|---|
| A | 御見積書／割付図・系統図／架台図／発電シミュレーション | ファイル自身の設定のまま |
| B（古い .xls） | 見積書／配置図（2ページ）／発電シミュレーション | A4・印刷範囲ごとに1ページへ補正（見積書は縦、他は横）。印刷範囲が無ければ既定範囲を設定 |

## 手元での確認（Docker）

```bash
docker build -t smm-review-converter .
docker run --rm -p 8090:8080 smm-review-converter
curl -F "file=@見積ファイル.xlsx" http://localhost:8090/api/convert-sheets -o out.pdf
```

サンプル見積は実在の顧客情報を含むため、`tests/fixtures/` に置き、Git には入れないでください（`.gitignore` 済み）。

## デプロイ

`main` に push すると GitHub Actions がデプロイします。必要な Secrets：

- `GCP_SA_KEY`：サービスアカウントの JSON 鍵（ロール：Cloud Run 管理者、Cloud Build 編集者、Artifact Registry 管理者、サービス アカウント ユーザー、Storage 管理者）
- `EMAIL_USERNAME` / `EMAIL_PASSWORD`：デプロイ結果のメール通知用

**初回は、画面側（SMM-REVIEW-TOOL）の rewrite 設定より先に、このサービスをデプロイしてください**（rewrite 先のサービスが無いと Firebase のデプロイが失敗します）。

## 注意

- 見積ファイルは一時フォルダーで変換後すぐに削除し、保存しません。ログにもファイル名・内容は出しません。
- 1台で同時に1件だけ変換します（`--concurrency 1`、最大3台）。使っていない間は停止するため、しばらくぶりの1回目は起動に時間がかかります。
- LibreOffice は Excel と別ソフトのため、フォントや改ページなど見た目が Excel の印刷結果と完全には一致しない場合があります。
