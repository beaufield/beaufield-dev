# 実装プラン：LINE WORKS 個人DM通知（毎日9時・期限超過アラート）

作成日: 2026-05-20  
ステータス: 未実装（プランのみ）

---

## 概要

貸出期限またはメーカー返却期限を超過した商品がある場合、その担当者のLINE WORKSに毎朝9時に個人DM通知を送る。  
通知先はグループではなく担当者個人のみ。既存のグループ通知（週次レポート等）はそのまま継続。

---

## アーキテクチャ方針

### LINE WORKS IDの一元管理

LINE WORKS ユーザーIDは **beaufield-auth の `users` シート**で一元管理する。  
→ 将来的に他のアプリでも個人通知が必要になったとき、beaufield-authを参照するだけでよい。

### 突合キー

担当者の紐づけは **user_id（beaufield-auth の user_id）** で行う（名前の表記揺れを避けるため）。

```
DeviceMaster.salesRep（名前）
    ↓ SalesRep シートで検索
SalesRep.auth_user_id（beaufield-auth の user_id）
    ↓ beaufield-auth.users で検索
users.lineworks_user_id（LINE WORKS ユーザーID）
    ↓
LINE WORKS Bot API でDM送信
```

---

## 通知対象の条件

| 種別 | 条件 |
|------|------|
| 貸出期限超過 | `status === '貸出中'` かつ `returnDueDate < 今日` |
| メーカー返却期限超過 | `makerReturnDueDate < 今日` かつ `status ≠ 'メーカー返却済'` かつ `status ≠ '廃棄'` |

- 超過案件が0件の担当者にはDMを送らない
- 各担当者には自分の担当分のみが届く

---

## 通知メッセージのイメージ

```
【期限超過アラート】2026-05-20

以下の商品の期限が超過しています。ご確認ください。

▼ 貸出期限超過（2件）
・BF-00001 フェイシャルスチーマー
  貸出先: 〇〇サロン / 期限: 2026-05-10

▼ メーカー返却期限超過（1件）
・BF-00003 カラーリングアイロン
  メーカー返却期限: 2026-04-30

早めのご対応をお願いします。
```

---

## 必要な変更一覧

### 1. beaufield-auth スプレッドシート（たかしさん作業）

`users` シートに **F列 `lineworks_user_id`** を追加する。

| 列 | 現在 | 変更後 |
|----|------|--------|
| A | user_id | user_id |
| B | name | name |
| C | pin | pin |
| D | active | active |
| E | created_at | created_at |
| **F** | —（空） | **lineworks_user_id（追加）** |

- 各ユーザーのLINE WORKSユーザーIDを入力する（例: `u1234567890123456789`）
- LINE WORKSユーザーIDはLINE WORKS Developer ConsoleまたはAdmin Consoleで確認可能

### 2. kiki-kanri SalesRep シート（たかしさん作業）

`SalesRep` シートに **`auth_user_id`** 列を追加する。  
各担当者の beaufield-auth `user_id`（例: `U001`）を入力する。

⚠️ **実装前に確認が必要**：DeviceMaster の `salesRep` フィールドが「名前」「ID」どちらを格納しているか確認すること。  
→ 名前の場合：SalesRep シートで名前 → auth_user_id を引く（現在の想定）  
→ IDの場合：直接 auth_user_id として扱えるので SalesRep シートの変更が不要になる可能性あり

### 3. kiki-kanri Script Properties（たかしさん作業）

以下を kiki-kanri の GAS スクリプトプロパティに追加する。  
値は expense-approval の GAS と同じものを使用（同一Botを流用）。

| キー名 | 内容 |
|--------|------|
| `LW_CLIENT_ID` | LINE WORKS Developer Console の Client ID |
| `LW_CLIENT_SECRET` | Client Secret |
| `LW_SA_ID` | Service Account ID |
| `LW_PRIVATE_KEY` | RSA秘密鍵（PEM形式・改行を `\n` に変換して1行で入力） |
| `LW_BOT_ID` | Bot ID（数字）|

※ `AUTH_SHEET_ID`（beaufield-auth の Spreadsheet ID）はすでに設定済みのはず。未設定の場合は `1cCQn16ubEN_Af7XWw8KerBscZtFomBnXHjIIiZUr6V8` を追加する。

### 4. kiki-kanri Code.gs（Claude が実装）

以下の関数を追加する：

```
getLwAccessToken()         // LINE WORKS OAuth2.0トークン取得（JWT Bearer）
sendLineWorksDm(userId, text)  // 個人DM送信
getLineWorksUserIdBySalesRep(salesRepName)  // 担当者名 → LINE WORKS ID の解決
sendDailyOverdueAlerts()   // 毎日9時に実行するメイン関数
```

### 5. GAS トリガー設定（たかしさん作業）

GAS の「トリガー」メニューから以下を設定する：
- 関数: `sendDailyOverdueAlerts`
- イベントのソース: 時間主導型
- 時間ベースのトリガーのタイプ: 日付ベースのタイマー
- 時刻: 午前9時〜10時

---

## LINE WORKS Bot について

- **既存Botを流用**（expense-approval で使用中のBot）
- Bot名を **「ViewField業務連絡Bot」** に変更する（LINE WORKS Developer Console で設定）
- expense-approval の通知には影響なし（送信先が異なるため）

---

## 実装時の注意点

### LINE WORKS API 認証（JWT Bearer）
GAS での実装は expense-approval の `poc_lw_auth.gs` / `Code.gs` を参考にそのままコピー流用できる。

### LINE WORKS ユーザーID の形式
LINE WORKS ユーザーIDは数字の文字列（例: `u1234567890123456789`）または管理コンソール上のアカウントID。  
expense-approval の `承認者マスタ` に登録済みの形式を確認・統一すること。

### beaufield-auth 参照パターン（既存と同じ）
```javascript
const authSheet = SpreadsheetApp.openById(AUTH_SHEET_ID);
const users = authSheet.getSheetByName('users').getDataRange().getValues();
const headers = users[0];
const lwIdCol = headers.indexOf('lineworks_user_id'); // 列名で動的取得
const userIdCol = headers.indexOf('user_id');
// user_id で検索して lineworks_user_id を返す
```

### メーカー返却期限フィールド名
DeviceMaster シートの「メーカー返却期限」列の内部名（`makerReturnDueDate` か別の名前か）を  
実装前に Code.gs で確認すること。

---

## 実装の順序

1. **たかしさん作業**
   - [ ] LINE WORKS Developer Console で Bot名を「ViewField業務連絡Bot」に変更
   - [ ] beaufield-auth `users` シートに F列 `lineworks_user_id` を追加・各ユーザーのIDを入力
   - [ ] kiki-kanri `SalesRep` シートに `auth_user_id` 列を追加・各担当者の user_id を入力
   - [ ] kiki-kanri GAS に Script Properties（5項目）を追加

2. **Claude が実装**
   - [ ] Code.gs に LINE WORKS DM通知関連の関数を追加（全文差し替えで提供）

3. **たかしさん作業**
   - [ ] Code.gs を GAS エディタに貼り付け・再デプロイ
   - [ ] GAS トリガーを設定（毎日9時）
   - [ ] テスト送信で動作確認

---

## 将来の拡張（今回は対象外）

- expense-approval の承認者マスタ内 LINE WORKS ID → beaufield-auth に一本化（移行）
- 新規アプリは最初から beaufield-auth 参照パターンで実装
