# SPEC.md — Beaufield ルート訪問チェッカー

## プロジェクト概要

**アプリ名**: Beaufield ルート訪問チェッカー  
**目的**: 美容ディーラー営業6名のルート訪問実績を日次で記録し、管理者・統括・部長が訪問履歴をマトリクスで確認できるWebアプリ  
**技術スタック**: HTML/CSS/JS（単一ファイル） + GitHub Pages + Google Apps Script（GAS） + Google Sheets  
**参考実装**: 既存の貸出管理アプリ（beaufield.github.io/kiki-kanri/）と同構成

---

## システム構成

```
[GitHub Pages] index.html（単一ファイル）
        ↕ fetch（CORS対応）
[Google Apps Script] doGet / doPost エンドポイント
        ↕
[Google Sheets]
  ├── users（ユーザーマスタ）
  ├── salons（サロンマスタ）
  └── visit_logs（訪問ログ）
```

---

## Google Sheets シート設計

### 1. `users` シート

| 列 | フィールド名 | 型 | 説明 |
|---|---|---|---|
| A | user_id | string | 一意ID（例: U001） |
| B | name | string | 氏名 |
| C | pin | string | 4桁PIN（平文） |
| D | role | string | `admin` / `director` / `manager` / `sales` |
| E | team | string | `A` / `B` / `all`（統括・管理者はall） |
| F | active | boolean | TRUE/FALSE |

**初期データ例**

| user_id | name | pin | role | team | active |
|---|---|---|---|---|---|
| U001 | Takashi（管理者） | 0000 | admin | all | TRUE |
| U002 | 統括者 | 0000 | director | all | TRUE |
| U003 | A部長 | 0000 | manager | A | TRUE |
| U004 | B部長 | 0000 | manager | B | TRUE |
| U005 | 営業A1 | 0000 | sales | A | TRUE |
| U006 | 営業A2 | 0000 | sales | A | TRUE |
| U007 | 営業B1 | 0000 | sales | B | TRUE |
| U008 | 営業B2 | 0000 | sales | B | TRUE |

---

### 2. `salons` シート

| 列 | フィールド名 | 型 | 説明 |
|---|---|---|---|
| A | salon_id | string | 一意ID（例: S0001） |
| B | salon_name | string | サロン名 |
| C | owner_user_id | string | 担当ユーザーID（usersのuser_id） |
| D | visit_day | string | 訪問曜日（月/火/水/木/金/土） |
| E | sort_order | number | 表示順（訪問順）。担当者ごとに1から採番 |
| F | active | boolean | TRUE/FALSE |

---

### 3. `visit_logs` シート

| 列 | フィールド名 | 型 | 説明 |
|---|---|---|---|
| A | log_id | string | 一意ID（タイムスタンプベース） |
| B | visited_at | datetime | チェック日時（ISO8601） |
| C | visit_date | date | 訪問日（YYYY-MM-DD） |
| D | user_id | string | チェックした営業のID |
| E | salon_id | string | サロンID |

---

## ロールと権限マトリクス

| 機能 | admin | director | manager | sales |
|---|---|---|---|---|
| 自分の訪問チェック | — | — | ✅ | ✅ |
| 全員の履歴閲覧 | ✅ | ✅ | — | — |
| 自部門の履歴閲覧 | ✅ | ✅ | ✅ | — |
| 自分の履歴閲覧 | ✅ | ✅ | ✅ | ✅ |
| 自担当サロン追加 | — | — | ✅ | ✅ |
| 自担当サロン名変更 | — | — | ✅ | ✅ |
| 自担当サロン曜日変更 | — | — | ✅ | ✅ |
| 自担当サロン表示順変更 | — | — | ✅ | ✅ |
| 全サロン編集 | ✅ | — | — | — |
| ユーザー管理 | ✅ | — | — | — |
| 自分のPIN変更 | ✅ | ✅ | ✅ | ✅ |
| 他者のPINリセット | ✅ | — | — | — |

---

## 画面一覧

### 画面A：ログイン画面

- ユーザー一覧からタップで氏名を選択
- 4桁PIN入力（テンキーUI）
- 認証成功 → ロールに応じた画面へ遷移
- 認証失敗 → エラーメッセージ表示

---

### 画面B：訪問チェック画面（sales / manager）

**表示条件**
- ログイン直後に表示
- 今日の曜日でフィルタされた担当サロンを `sort_order` 順に表示

**リスト表示**
- サロン名 ＋ チェックボタン
- 当日チェック済み → チェックマーク＋グレーアウト（当日は重複記録なし）
- 未チェック → タップで即時保存（GASへPOST）

**曜日外サロン追加チェック（イレギュラー対応）**
- 「＋他のサロンを追加」ボタン
- 自担当サロン全件をサロン名検索でフィルタ → タップでチェック

**担当外サロンへの訪問（新規開拓等）**
- 「＋新規サロンに訪問」ボタン
- サロン名を入力して新規サロン登録と同時にチェック記録

---

### 画面C：マイサロン管理画面（sales / manager）

**機能**
- 自担当サロン一覧表示（全曜日）
- 新規サロン追加（サロン名・訪問曜日を入力）
- サロン名編集
- 訪問曜日変更
- 表示順変更（上下ボタン or ドラッグ＆ドロップ）
- サロンの無効化（削除ではなくactive=FALSE）

---

### 画面D：訪問履歴ビュー（全ロール・閲覧範囲はロール依存）

**レイアウト**
- 縦軸：サロン名
- 横軸：週（例「3/第1水」「3/第2水」…）
- セル：✅（訪問済） / 空白（未訪問）
- **4週連続空白 → 該当セルを赤背景で警告表示**

**フィルタ**
- 担当者フィルタ（プルダウン）
  - sales: 自分のみ固定
  - manager: 自部門メンバーから選択
  - director: 全員から選択
  - admin: 全員から選択
- 曜日フィルタ（月〜土・全て）
- 表示月フィルタ（当月デフォルト）

---

### 画面E：管理者設定画面（admin のみ）

**ユーザー管理**
- ユーザー一覧表示
- 新規ユーザー追加（氏名・PIN・ロール・チーム）
- ユーザー編集（氏名・ロール・チーム・active）
- PINリセット（任意のユーザーの PIN を 0000 に戻す）
- ユーザー無効化（active=FALSE）

**全サロンマスタ閲覧**
- 全サロンの一覧表示
- サロン名・担当者・訪問曜日の編集
- ※担当変更・訪問曜日の一括変更はスプレッドシート直接編集でOK（アプリ内一括変更機能は不要）

---

## GAS エンドポイント設計

### 共通仕様
- `doGet(e)` / `doPost(e)` で全リクエストを受け付け
- レスポンスは JSON
- 認証はアプリ側でPIN照合済みの前提でGAS側は簡易チェックのみ（action＋user_idをパラメータで渡す）

### アクション一覧

| action | メソッド | 説明 |
|---|---|---|
| `login` | GET | user_id＋PINで認証。ユーザー情報を返す |
| `getMyRoute` | GET | 今日の曜日×担当者でサロン一覧取得 |
| `getMySalons` | GET | 担当者の全サロン取得（マイサロン管理用） |
| `getTodayLogs` | GET | 担当者の当日チェック済みサロンID一覧取得 |
| `checkVisit` | POST | 訪問チェックを visit_logs に追記 |
| `addSalon` | POST | 新規サロン登録 |
| `updateSalon` | POST | サロン情報更新（名称・曜日・表示順） |
| `deactivateSalon` | POST | サロン無効化 |
| `updateSortOrder` | POST | 表示順一括更新 |
| `getVisitHistory` | GET | 履歴マトリクス用データ取得（期間・担当者指定） |
| `getUsers` | GET | ユーザー一覧取得（admin用） |
| `addUser` | POST | ユーザー追加（admin用） |
| `updateUser` | POST | ユーザー更新（admin用） |
| `resetPin` | POST | PIN リセット（admin用） |
| `changePin` | POST | 本人によるPIN変更 |
| `getAllSalons` | GET | 全サロン取得（admin用） |
| `updateSalonAdmin` | POST | サロン更新（admin用・全件対象） |

---

## ビジネスロジック

### 重複チェック防止
- 同一ユーザー・同一サロン・同一日付の visit_logs が存在する場合は追記しない（GAS側でチェック）

### 4週連続未訪問の計算
- 履歴ビュー取得時に GAS 側で計算して返す
- 「当該サロンの訪問曜日」において直近4週分のレコードを確認
- 4週全て空の場合は `alert: true` フラグを返す

### 表示順（sort_order）の管理
- 担当者ごとに独立した連番
- 並び替え時は関係するサロンの sort_order を一括更新

---

## UI・UX 要件

- スマートフォン（iOS/Android Safari/Chrome）での操作を主想定
- タップターゲットは最小44px
- 訪問チェック画面はログイン後に最初に表示（ステップ数を最小化）
- オフライン時はエラーメッセージを表示（オフライン対応は不要）
- 日本語UI

---

## ファイル構成（GitHub リポジトリ）

```
route-checker/
├── index.html        ← アプリ本体（HTML/CSS/JS 単一ファイル）
├── gas/
│   └── Code.gs       ← Google Apps Script 本体
├── SPEC.md           ← 本仕様書
└── README.md         ← セットアップ手順
```

---

## 実装優先順位

| フェーズ | 内容 |
|---|---|
| Phase 1 | GASエンドポイント全実装 ＋ Google Sheetsセットアップ |
| Phase 2 | ログイン画面 ＋ 訪問チェック画面 |
| Phase 3 | 訪問履歴ビュー（マトリクス表示・アラート） |
| Phase 4 | マイサロン管理画面（追加・編集・並び替え） |
| Phase 5 | 管理者設定画面（ユーザー管理・全サロン管理） |
| Phase 6 | PIN変更・PINリセット機能 |

---

## セットアップ手順（README用メモ）

1. Google Sheetsを新規作成し、`users` / `salons` / `visit_logs` シートを作成
2. `users` シートに初期ユーザーデータを入力
3. GASスクリプトをスプレッドシートのApp Scriptに貼り付け
4. GASをWebアプリとしてデプロイ（アクセス：全員、実行：自分）
5. デプロイURLを `index.html` 内の `GAS_URL` 定数に設定
6. `index.html` を GitHubリポジトリ `beaufield/route-checker` にプッシュ
7. GitHub Pages を有効化（branch: main / root）
8. 公開URLをLINE WORKSで営業に共有・ブックマーク登録依頼

---

*作成日: 2026-03-21*  
*バージョン: 1.0.0*
