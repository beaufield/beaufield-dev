# AI発注補助・クラウド試作パッケージ

これは発注画面と共有する解釈契約・計算器と、合成データによるCloudタスク試験用コード。
アプリからCodex Cloudタスクを自動呼出しするAPI、HTTP公開サーバー、認証情報は含まない。

## クラウド環境での試験

Node.js 18以上。追加npmパッケージ・APIキー・DB接続は不要。
`manifest.json` のファイルだけを準備する。価格・顧客・実在庫・Dropbox設定は含めない。

1. `node ai-order/cloud-task.cjs prompt` で条件解釈用の入力とスキーマを表示する。
2. Codex Cloudのタスク自身が、その入力からJSONの条件を作る。別のCodexやResponses APIは起動しない。
3. JSONを `node ai-order/cloud-task.cjs evaluate` の標準入力へ渡す。
4. `response.status`、確認質問、ケース別合計、未対応条件、警告を確認する。

fixtureのデータは作ったもの。6ケースずつは1000が72本、3000が36本。
結果を本番カートに使わない。実データ・本番権限・公開接続は別段階。

## ファイルと責任

- `interpretation.js`: 依頼、質問への回答、既知仕入先をAIの条件JSONへ変換する契約。欠損、未対応、余分なキーを検証。
- `core.generated.cjs`: `index.html` と同じ数量計算・在庫警告・未入荷等の禁止規則。
- `worker.cjs`: ChatGPTプランproviderを差し替える境界と、本人が条件を確認した後の計算。
- `cloud-task.cjs`: Cloudタスクの合成試験。永続バックエンドではない。
- `local-probe-provider.cjs`: ローカルの合成試験だけに使う既存CLI認証。配布manifest対象外。遠隔ホストへ移さない。

## 画面への接続

画面側は `aiOrderConfigureProvider({ usage: 'chatgpt-plan', interpret(request, {signal}) })` で接続を受け取る。
本番providerは未設定なので、AI相談ボタンは表示しない。これは適格性判定そのものではない。
接続先選択・認証・Pro利用条件・iPhone往復が成立した後、正式providerを実装する。
AIは発注先、系列、サイズ、ケース数、入数、追加日数を解釈し、未確定なら質問する。
AIが条件を返しても計算・カート反映・発注保存は自動実行しない。

数量計算には分析ID・時刻・商品別集計・商品マスター・通常提案・分析後未入荷の商品コードが必要。
条件解釈のAIへは依頼文と仕入先候補のみ送る。将来の在庫や価格の説明用データ提供は別途設計する。
snapshotはリアルタイム在庫ではない。カート反映時は既存画面で分析世代と未入荷等を再照合する。

## 再生成と検証

```
node tools/build_ai_order_bundle.js
node tools/build_ai_order_bundle.js --check
node tests/test_ai_order_pipeline.js
node tests/test_ai_order_assist.js
```

Cloud環境はまだ未作成。GitHubへのコミット/公開も未実施。
先にプラン内呼出しの正規経路を確定する。API従量課金への自動フォールバックはない。
