# 引継情報 — Beaufield ルート訪問チェッカー
最終更新: 2026-04-03（v1.9.2 / GAS v1.5.3）

---

## プロジェクト概要
- **アプリ名**: Beaufield ルート訪問チェッカー
- **技術構成**: HTML単一ファイル + GitHub Pages + Google Apps Script + Google Sheets

---

## GitHubリポジトリ
- リポジトリ: beaufield/beaufield-dev
- 作業ディレクトリ: 開発・自動化/route-checker/
- ブランチ: main
- GitHub Pages URL: `https://beaufield.github.io/beaufield-dev/route-checker/`

---

## Google Sheets 情報
- スプレッドシートID: `1yVd3yI9v8acjyKaM-fCs_VoBnOx44mnRDOBVGzBB288`
- シート構成:
  - `users`: user_id / name / pin / role / team / active / display_order
  - `salons`: salon_id / salon_name / owner_user_id / visit_day / sort_order / active / code
  - `visit_logs`: log_id / visited_at / visit_date / user_id / salon_id

---

## GAS WebApp
- デプロイ済みURL: `https://script.google.com/macros/s/AKfycbxaJg7FmAxJBc-miR-Xc7Aa-PFO8gcjZqcCrFxAKLZ93E4syqyYHGxF2JcUHmRH8tr_/exec`
- 実行ユーザー: 自分 / アクセス: 全員

---

## バージョン状況（2026-04-03時点）
| ファイル | バージョン | 状態 |
|----------|-----------|------|
| gas/Code.gs | v1.5.3 | ✅ GASデプロイ済み |
| index.html | v1.9.2 | ✅ GitHub Pages 反映済み |

---

## フェーズ完了状況
| Phase | 内容 | 状態 |
|-------|------|------|
| Phase 1 | GAS + Sheetsセットアップ | ✅ 完了 |
| Phase 2 | ログイン・訪問チェック画面 | ✅ 完了 |
| Phase 3 | 訪問履歴マトリクス（フィルタ・4週アラート） | ✅ 完了 |
| Phase 4 | マイサロン管理（追加・編集・並び替え・無効化） | ✅ 完了 |
| Phase 5 | 管理者設定（ユーザー管理・全サロン管理） | ✅ 完了 |
| Phase 6 | PIN変更・PINリセット機能 | ✅ 完了 |

---

## 実装済み機能一覧

### 認証・セッション
- ユーザー選択 → 4桁PIN入力でログイン
- **localStorage に30日間セッション保持**（再ログイン不要）
- ログアウトでセッション削除

### チェック画面（sales/manager）
- **日付選択機能**: 今日〜5日前のクイックボタン + 日付ピッカー
- 選択日の曜日に対応するルートを表示
- 過去日付へのチェック記録対応
- 曜日外サロン追加・新規サロン訪問
- **高速化（v1.8.0〜）**: チェック操作は即座にUI更新（バックグラウンド送信）、日付変更はローカル計算、起動時はサロンキャッシュ（localStorage・1時間）活用
- **チェック取り消し（v1.9.0〜）**: チェック済みサロンを再タップで取り消し（今日〜5日前分）
- **チェック日時表示（v1.9.0〜）**: チェック画面はボタン左に「M/D HH:MM」、履歴画面は✅の下に表示

### 履歴マトリクス（全ロール）
- 月・曜日・ユーザーフィルタ
- 4週以上未訪問アラート
- サロン列幅固定（130px）・折り返し表示

### マイサロン管理（sales/manager）
- 曜日フィルタ（火〜土・随時）
- 並び替え・編集・無効化
- コード表示

### 管理者設定（admin）
- ユーザー管理（追加・編集・無効化・PINリセット）
- 表示順並び替え（ログイン画面の順番に反映）
- 全サロン管理（独立画面・サロン名/コード検索）

---

## 曜日仕様
- **月曜日は定休日のため曜日選択肢に含まれない**（火〜土・随時のみ）
- 履歴フィルタも同様（月なし）

---

## 既知の問題・解決済みトラブル
1. **PINが一致しない**: SheetsがPIN `0000` を数値 `0` に変換
   - 対処: C列「書式なしテキスト」設定 + Code.gsで `padStart(4,'0')` 正規化
2. **並び替えが効かない**: sort_order重複時に値交換で見た目が変わらない
   - 対処: 位置入れ替え＋連番再採番方式（v1.2.1〜）
3. **チェック画面の順番が追従しない**: todayRoute と allMySalons が別オブジェクト
   - 対処: 並び替え後に todayRoute も同期（v1.2.2〜）
4. **無効化が失敗する**: closeSalonEditModal()でeditingSalonIdがnullになる前に実行
   - 対処: targetId変数に退避してからclose（v1.3.6〜）
5. **git commit できない**: Dropbox が .git/objects をロック
   - 対処: GitHub API（PowerShell Invoke-RestMethod）で直接プッシュ
6. **チェック取り消しが反映されない**: uncheckVisit で visit_date をgetValues()直読みするとDate型になり文字列比較が失敗
   - 対処: `instanceof Date` チェックして formatDate で正規化（v1.5.3〜）
7. **時刻が再読み込みで消える**: _readSheet の Date 変換が 'yyyy-MM-dd' のみで時刻を切り捨てていた
   - 対処: 時刻情報がある場合は "yyyy-MM-dd'T'HH:mm:ss" 形式で保持（v1.5.1〜）

---

## git が使えない場合のプッシュ方法

```powershell
# 現在のファイルSHAを取得
curl -s -H "Authorization: token TOKEN" \
  "https://api.github.com/repos/beaufield/beaufield-dev/contents/route-checker/index.html" \
  | grep '"sha"'

# PowerShellでプッシュ
powershell.exe -Command "& {
  $token = 'TOKEN'
  $path  = 'D:\Dropbox\ClaudeWork\開発・自動化\route-checker\index.html'
  $bytes = [System.IO.File]::ReadAllBytes($path)
  $b64   = [Convert]::ToBase64String($bytes)
  $body  = @{ message='コミットメッセージ'; content=$b64; sha='FILE_SHA' } | ConvertTo-Json -Depth 3
  $headers = @{ Authorization='token '+$token; 'Content-Type'='application/json'; 'User-Agent'='ps' }
  $res = Invoke-RestMethod -Uri 'https://api.github.com/repos/beaufield/beaufield-dev/contents/route-checker/index.html' -Method Put -Headers $headers -Body $body
  Write-Output ('OK: ' + $res.commit.sha.Substring(0,12))
}"
```

---

## セッション開始時の指示

```
Beaufield ルート訪問チェッカーの続きをお願いします。
HANDOVER.md を参照してください。
作業ディレクトリ: D:\Dropbox\ClaudeWork\開発・自動化\route-checker
```
