# 引継情報 — Beaufield ルート訪問チェッカー
最終更新: 2026-03-21

---

## プロジェクト概要
- **アプリ名**: Beaufield ルート訪問チェッカー
- **技術構成**: HTML単一ファイル + GitHub Pages + Google Apps Script + Google Sheets
- **仕様書**: SPEC.md（このディレクトリ内）

---

## GitHubリポジトリ
- リポジトリ: beaufield/beaufield-dev
- 作業ディレクトリ: 開発・自動化/route-checker/
- ブランチ: main

---

## Google Sheets 情報
- タイトル: BFルート訪問チェッカー
- URL: https://docs.google.com/spreadsheets/d/1yVd3yI9v8acjyKaM-fCs_VoBnOx44mnRDOBVGzBB288/edit
- スプレッドシートID: `1yVd3yI9v8acjyKaM-fCs_VoBnOx44mnRDOBVGzBB288`
- シート構成: `users` / `salons` / `visit_logs`（setupSheets で作成済み）

---

## GAS WebApp
- デプロイ済みURL: `https://script.google.com/macros/s/AKfycbxaJg7FmAxJBc-miR-Xc7Aa-PFO8gcjZqcCrFxAKLZ93E4syqyYHGxF2JcUHmRH8tr_/exec`
- 現在のバージョン: **Code.gs v1.1.2**（GASにデプロイ済み）
- 実行ユーザー: 自分 / アクセス: 全員

---

## バージョン状況
| ファイル | バージョン | GAS反映 |
|----------|-----------|---------|
| gas/Code.gs | v1.1.2 | ✅ デプロイ済み |
| index.html | v1.2.2 | ⚠️ ローカルのみ・GitHub Pages未反映 |

---

## フェーズ完了状況
| Phase | 内容 | 状態 |
|-------|------|------|
| Phase 1 | GAS + Sheetsセットアップ | ✅ 完了 |
| Phase 2 | ログイン・訪問チェック画面 | ✅ 完了 |
| Phase 3 | 訪問履歴マトリクス（フィルタ・4週アラート） | ✅ 完了 |
| Phase 4 | マイサロン管理（追加・編集・並び替え・無効化） | ✅ 完了 |
| Phase 5 | 管理者設定（ユーザー管理・全サロン管理） | ⬜ 未着手 |
| Phase 6 | PIN変更・PINリセット機能 | ⬜ 未着手 |

---

## 初期ユーザー（users シート）
| user_id | name | role | team | PIN |
|---------|------|------|------|-----|
| U001 | Takashi（管理者） | admin | all | 0000 |
| U002 | 統括者 | director | all | 0000 |
| U003 | A部長 | manager | A | 0000 |
| U004 | B部長 | manager | B | 0000 |
| U005 | 営業A1 | sales | A | 0000 |
| U006 | 営業A2 | sales | A | 0000 |
| U007 | 営業B1 | sales | B | 0000 |
| U008 | 営業B2 | sales | B | 0000 |

⚠️ usersシートのC列（pin）は「書式なしテキスト」に設定すること（数値に変換されると認証失敗）

---

## 既知の問題・解決済みトラブル
1. **PINが一致しない**: SheetsがPIN `0000` を数値 `0` に変換する問題
   - 対処: C列を「書式なしテキスト」に設定 → `0000` を再入力
   - Code.gsでも `padStart(4,'0')` で正規化済み（v1.1.1〜）

2. **並び替えが効かない**: sort_order 重複時に値の入れ替えで見た目が変わらない
   - 対処: 「位置入れ替え＋連番再採番」方式に変更（v1.2.1〜）

3. **チェック画面の順番が並び替えに追従しない**: `todayRoute` と `allMySalons` が別オブジェクト
   - 対処: 並び替え後に `todayRoute` も同期してソート（v1.2.2〜）

---

## 次セッションでやること（Phase 5 から）

### Phase 5：管理者設定画面（admin のみ）
- ユーザー一覧表示・新規追加・編集（氏名・ロール・チーム）・無効化
- 他者のPINリセット（0000 に戻す）
- 全サロンマスタ閲覧・サロン名/担当者/曜日の編集

### Phase 6：PIN変更機能（全ロール）
- 自分のPINを現在のPINで認証してから新PINに変更
- 管理者は他者のPINを0000にリセット可能

### その後
- GitHub Pages へ index.html をコミット・push
- 本番データ投入（実際の営業名・サロン名）
- LINE WORKS で営業メンバーにURL展開

---

## セッション開始時の指示
次のセッションでは以下のように指示してください：

```
Beaufield ルート訪問チェッカーの続きをお願いします。
HANDOVER.md を参照してください。
Phase 5（管理者設定画面）から進めてください。
作業ディレクトリ: D:\Dropbox\ClaudeWork\開発・自動化\route-checker
```
