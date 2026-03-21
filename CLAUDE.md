# 開発・自動化フォルダの設定

## このフォルダで扱う内容
- 業務効率化・自動化ツールの開発
- 社内業務フローの改善・システム化
- Googleカレンダー・LINE WORKS等との連携ツール
- データ集計・レポート自動化

## Claudeへの指示
- 使用言語は原則Pythonとすること
- コードには日本語コメントを入れること
- 必ず動作確認手順を一緒に示すこと
- 非エンジニアが使うことを前提に、操作をシンプルに保つこと

## バージョン管理ルール
- アプリケーションを作る際は、画面上の分かりやすい場所（タイトル横・フッター等）に必ずバージョン番号を表示すること
- バージョン形式は `v1.0.0`（メジャー.マイナー.パッチ）を基本とする
- コードの修正・エラー修正のたびにバージョンを更新し、最新バージョンを操作しているか即座に判断できるようにすること
  - バグ修正・軽微な修正 → パッチ番号を上げる（例: v1.0.0 → v1.0.1）
  - 機能追加・改善 → マイナー番号を上げる（例: v1.0.0 → v1.1.0）
  - 大幅な変更・再設計 → メジャー番号を上げる（例: v1.0.0 → v2.0.0）

## コード更新ルール
- コードを更新する際は**差分更新は行わず、必ず全文差し替えで提供**すること
- 部分的なコード提示は誤適用・混乱の原因になるため厳禁
- 更新後のコードは常に完全な状態で、そのまま動作できる形で提示すること

---

## GAS（Google Apps Script）WebApp 開発ルール

### ファイル構成
- `Code.gs`：バックエンド（データ読み書き・LINE WORKS通知）
- `index.html`：フロントエンド（HTML/CSS/JS全部入り）

### デプロイワークフロー
1. ClaudeがDropbox内のローカルファイルを更新
2. たかしさんがGASエディタに**全文コピー貼り付け**（部分差し替えNG）
3. 「デプロイ」→「デプロイを管理」→新バージョンで再デプロイ
4. URLは変わらない（同じデプロイIDを上書き更新）

### バージョン管理
- index.htmlのフッターに `HTML v2.x` を記載（画面右下または薄い透過文字で）
- Code.gsの `VERSION` 定数でGAS側バージョンを管理
- 管理パネルに両バージョンを表示して確認できるようにする

### 注意：複数Googleアカウント問題
- 同一ブラウザに複数のGoogleアカウントでログインしているとアクセスエラーが起きる
- 回避策：シークレットモードを使用するか、1アカウントのみのブラウザで開く

---

## GAS WebApp スマホ最適化（重要）

GASのWebAppはviewportを**強制的に980pxに固定**するため、
iPhoneでは全体が約0.4倍（393/980）に縮小表示される。

### 解決策（bodyタグ直後に記述）
```javascript
(function fixScale(){
  var iw = window.innerWidth; // GAS fake: 980px
  var sw = screen.width;      // 実際のデバイス幅（iPhone: 393px等）
  if (sw > 0 && sw < iw) {
    var zoom = iw / sw;                    // 980/393 ≈ 2.49
    document.body.style.zoom     = zoom;
    document.body.style.width    = sw + 'px';
    document.body.style.maxWidth = sw + 'px';
    document.body.style.overflowX = 'hidden';
  }
})();
```

### レイアウト動的設定（メディアクエリの代替）
```javascript
function applyLayout() {
  var mobile = window.screen.width < 768; // screen.widthで判定（innerWidthはNG）
  var grid = document.getElementById('myGrid');
  if (grid) grid.style.gridTemplateColumns = mobile ? 'repeat(2,1fr)' : 'repeat(4,1fr)';
}
// グリッドをJSで動的生成した後は必ず applyLayout() を再呼び出し
```

### ポイント
- `window.innerWidth` → 常に980px（GASの偽値）→ **使用禁止**
- `window.screen.width` → 実際のデバイス幅 → **こちらを使う**
- CSSメディアクエリは機能しない → JSで動的設定する
- ボタンテキストが長い場合は `white-space: nowrap` で折り返しを防止

---

## LINE WORKS Incoming Webhook 設定

### 設定手順
1. LINE WORKS管理画面 → 「Bot」 → 対象Botを選択
2. 「Incoming Webhook」タブ → URLを発行・コピー
3. Webhook URLを `Code.gs` の定数に設定

### GASからの送信コード（Code.gs）
```javascript
var LINEWORKS_WEBHOOK = 'https://talk.worksmobile.com/bot/hooks/xxxxxxxx';

function sendLineWorks(message) {
  try {
    UrlFetchApp.fetch(LINEWORKS_WEBHOOK, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ content: message }),
      muteHttpExceptions: true
    });
  } catch(e) {
    // 通知失敗してもメイン処理は継続
    Logger.log('LINE WORKS通知エラー: ' + e);
  }
}
```

### ポイント
- Webhook URLは外部に漏らさない（Code.gsの定数として管理）
- 通知失敗がメイン処理に影響しないよう `try-catch` で囲む
- `muteHttpExceptions: true` でHTTPエラーもキャッチする
- メッセージの改行は `\n` で表現（LINE WORKSで反映される）
