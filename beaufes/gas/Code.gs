// ============================================================
// ビューフェス申込アプリ - Google Apps Script
// ============================================================
// [重要] コードにIDを直書きしない。以下の手順でスクリプトプロパティに設定すること。
//
// GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」→「プロパティを追加」
//   SPREADSHEET_ID    : ビューフェス申込データのスプレッドシートID（新規に「beaufes2026」という名前で作成する）
//   LINEWORKS_WEBHOOK : （任意・v0.14.0）新規申込のLINE WORKS通知先Incoming Webhook URL。
//                        未設定でも動作する（通知が飛ばないだけ）。発行手順は_handoff.md参照。
//
// 初回セットアップ手順:
//   1. 新規Googleスプレッドシートを作成（名前: beaufes2026）
//   2. 拡張機能 → Apps Script でこのファイルの内容を貼り付け
//   3. プロジェクトの設定でスクリプトプロパティ SPREADSHEET_ID を設定
//   4. GASエディタの関数選択で setupSheets を選び、一度だけ実行（シート・見出し・configの初期値を作成）
//   5. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
//        実行ユーザー: スプレッドシート作成に使った管理用Googleアカウント
//        アクセスできるユーザー: 全員（お客様の申込フォームが認証なしで動く必要があるため）
//   6. 発行されたウェブアプリURLを index.html / pass.html の GAS_URL に設定
//
// メール送信元（beaufes@gmail.com）を実際に使うには、事前に
// GAS実行アカウントのGmail設定で「他のメールアドレスを追加」から
// beaufes@gmail.com を送信元アドレス（send-as）として登録しておくこと。
// 未登録の場合は自動的に MailApp（送信元は実行アカウント・差出人名とReply-Toで代替）にフォールバックする。
//
// 🔴 送信量の注意（2026-08-04 実測確認済み）: GAS実行アカウントはGoogle Workspace Business
// Standardを契約しているにもかかわらず、Apps Scriptの送信枠は個人アカウント扱いで1日100通だった
//（MailApp.getRemainingDailyQuota()で実測確認済み。契約と実測の食い違いの原因は未特定・追跡はせず
// 実測値を設計上の正として採用。checkMailQuota()関数でいつでも再確認できる）。
// 申込受付中の分散送信は問題ないが、前日リマインド等で200名に一括送信する機能を作る際は、
// 50〜80件ずつ複数回に分けて送るなどの対策が必須。
// 詳細: LINEHarness/ビューフェス申込_設計.md §7-0-1〜7-0-2
//
// 🆕 v0.5.0（2026-08-05・設計書v3対応）: applications シートに business_type 列（U列=21列目）を追加。
// 既存の本番シート（beaufes2026）には自動で列が増えないため、デプロイ後に一度だけ
// migrateAddBusinessType() をGASエディタから手動実行してヘッダーを追加すること。
// 新規にsetupSheetsでシートを作る場合はヘッダーに最初から含まれるため不要。
//
// 🆕 v0.6.0（2026-08-06）: エリア項目をフォームから廃止（用途は地域分布の集計のみで、
// お客様に選ばせる必要がないと判断。Takashiさん指示）。J列(area)は列位置維持のため
// 残すが常に空文字を書き込む。詳細: LINEHarness/ビューフェス申込_設計.md §4-1
// 🆕 v0.6.1: 別端末で先行push済みのbusiness_type機能(v0.5.0)とのマージ調整のみ。機能変更なし
//
// 🆕 v0.7.0（2026-08-06）: 重複申込対策（設計書§4-3の未実装分を実装）
//   1. 氏名の照合を正規化（`_normalizeName`）。「山田太郎」「山田 太郎」「山田　太郎」を同一人物として扱う
//   2. 同じメールアドレスの申込が既にある場合、無言で更新せず確認画面用の応答を返す
//      （`duplicate_found:true`）。クライアントが mode='update'/'new' を付けて再送信して確定する。
//   3. `resendPass` アクションを新設。パスURLを紛失した人がメールアドレスだけで再発行できる
//      （フォーム再送信で重複行が生まれる経路を塞ぐため）
//   → v0.8.0でこの確認画面方式は**廃止**（下記参照）。
//
// 🆕 v0.8.0（2026-08-06・L0-X）: 重複確認画面を廃止し「トークンURLを知っている人だけ編集できる」方式へ移行。
//   問題: v0.7.0の`applyApplication`は、メールアドレスが一致すると`duplicate_found`とともに
//   **既存申込のsalon_name/staff_name/created_atをそのまま返していた**。他人のメールアドレスを
//   入力すれば、その人の参加事実・サロン名・氏名が分かってしまう情報漏洩経路だった。
//   （`resendPass`は列挙対策済みだったが`apply`側が抜けていた）
//
//   対策: 確認画面そのものを廃止。`applyApplication`から`mode`/`target_app_id`/`duplicate_found`を削除し、
//   以下の3分岐のみにする。
//     - メールに一致なし               → そのまま新規登録
//     - メール一致 かつ 正規化氏名も一致 → 既存行は一切変更せず・情報も一切返さず、
//                                        パスURLをメール再送。応答は{existing_notified:true}のみ
//     - メール一致 だが氏名が違う       → 別人として新規登録（同一サロン複数名の代表メール申込に対応）
//   漏れる情報は「そのメール＋その氏名の組み合わせが登録済みか」という真偽値のみになる。
//
//   新設 `updateApplication`: ticket_tokenで行を特定して更新する唯一の経路（capability URL方式）。
//   ticket_token・app_id・statusは書き換えない。重複判定は行わない（本人の編集のため）。
//   `getPass`もpass.htmlの編集UI向けにphone/email/business_type/has_transaction/address/referrer/noteを返すよう拡張。
//
//   受容したリスク: QR提示時に第三者が撮影するとticket_tokenを入手でき、内容を閲覧・編集できてしまう。
//   イベント規模・性質から実害は小さいと判断し受容（Takashiさん了解済み・token は16桁の乱数で総当たり不可能）。
//   詳細: 開発・自動化/beaufes/LINE連携_実装プラン.md §2-2・§4 L0-X
//
// 🆕 v0.8.1（2026-08-06）: 電話番号の先頭「0」が消える不具合を修正（pass.htmlの編集フォームで発覚）。
//   原因: 全て数字の文字列を書き込むと、列の書式がAutomaticのままだとGoogle Sheetsが
//   数値型に変換し先頭の0が失われる。applyApplication・updateApplicationとも、電話番号セルを
//   書き込み直前にPlain Text指定するよう修正。既存データの復旧は migrateFixPhoneColumn() で行う
//  （本番シートに対して一度だけ手動実行すること）。
//
// 🆕 v0.9.0（2026-08-06・L1-a）: LINE Harness連携の土台（UIなし・doGet/doPostからはまだ呼ばれない）。
//   新設: `_lhReq` / `_lhFetchAllFriends` / `syncLineFriends` / `_findFriendByLineUserId` / `probeLineHarness`。
//   友だちは `line_user_id` での直接検索手段がAPIに無いため、全件を `line_friends_cache` シートに
//   キャッシュし、時間トリガー（6時間ごと・GASエディタで手動設定）で同期する方式にした（§2-1）。
//   Script Propertiesに `LINE_HARNESS_API_URL` / `LINE_HARNESS_API_KEY` / `BEAUFES_TAG_ID` が必要
//  （2026-08-06 Takashiさんが登録済み。デプロイ・実機確認済み: probeLineHarnessでtotal:103・
//   syncLineFriendsで103件同期・6時間ごとの時間トリガーも設定済み）。
//   詳細: 開発・自動化/beaufes/LINE連携_実装プラン.md §1・§2-1・§4 L1-a
//
// 🆕 v0.9.1（2026-08-06・L1-b）: LINEのIDトークンをサーバー側で検証する `_verifyLineIdToken` を新設。
//   `liff.getProfile()` の結果をそのまま信用せず、`POST https://api.line.me/oauth2/v2.1/verify` の
//   応答（aud/exp/iss）を自分で検証してから `sub`（=lineUserId）を使う（§2-3）。まだ呼び出し元は無い。
//   診断用に `testVerifyBadIdToken()` を追加（不正なトークンで確実に例外になることをGASエディタから確認できる）。
//   実機確認済み: 不正なトークンはLINE側からHTTP_400（JWS format error）で拒否され、正しく例外化される。
//
// 🆕 v0.10.0（2026-08-06・L1-c）: LIFF申込アクション `liffPrefill` / `applyLiff` を新設。
//   applyApplication・updateApplicationのバリデーション・書き込みロジックを共通関数
//  （`_validateApplicationFields` / `_appendApplicationRow` / `_updateApplicationRow`）に切り出し、
//   Web版・LIFF版で二重管理しないようにリファクタリング（挙動は変えていない）。
//   liffPrefill: IDトークン検証→line_friends_cacheと既存申込(line_user_id)からプリフィル値を返す。
//   6択に無い業態（旧表記「エステ」等）は省略する（§1-4 地雷5）。
//   applyLiff: line_user_idで既存申込が引ければ更新（メール送信・確認画面なし。§2-10）、
//   引けなければsource='liff'で新規登録。まだ呼び出し元（liff.html）は無い（L1-dで新規作成）。
//
// 🆕 v0.11.0（2026-08-06・L2）: LINE側の仕上げ（`_syncApplicationToLine`）を新設し、applyLiffに接続。
//   ① 新規登録時のみ入場パスURL付きでLINEプッシュ（更新時は送らない。メールと同じ考え方）
//   ② 「ビューフェス2026申込」タグ付与（新規・更新とも試す）
//   ⑤ friendのmetadataの**空欄だけ**書き戻す（§2-4。既に値がある項目は絶対に送らない）。
//      判定はキャッシュではなく`_lhGetFriend`で取得した最新値を使う（stale判定での誤上書き事故を防ぐ）
//   🔴 いずれも失敗しても申込自体は失敗させない。各ステップを個別にtry/catchし、
//   失敗は新設の`line_sync_log`シートに記録して後から手動で復旧できるようにした。
//
// 🆕 v0.12.0（2026-08-06・診断）: 通信失敗の実測用に `ping` / `pingHeavy` アクションを追加（読み取りのみ）。
//   背景: 実行ログの実測で「全実行が0.3〜3.1秒で完了・エラーゼロなのに、その結果がブラウザに
//   一度も届かず、クライアントが15秒でタイムアウト→3回リトライして諦める」パターンを確認した
//   （2026-08-06 17:20台の3連続実行と、直後の空欄プリフィル画面が一致）。
//   → スクリプトの遅さではなく**結果の配送層**が疑わしいため、`diag.html` から連続実行して
//   実際の失敗率を数える。pingLightはシート無し・pingHeavyはliffPrefill同等の全行読み込み。
//
// 🆕 v0.13.0（2026-08-06・配送障害対策P1/P2）: `GAS配送障害_対策計画.md` に基づく対症療法。
// 根本原因（Google側の結果配送層）は直せないため、「リトライしても壊れない」構造にする。
//   P1: `apply`（Web申込）に冪等キー(`request_id`)を導入。
//     クライアントはページ読み込み時に1つUUIDを生成し、全リトライで同じ値を送る。
//     applications シートにV列(22列目)=`request_id`を追加。
//     同じrequest_idの行が既に存在すれば「配送失敗による再試行」と判断し、
//     新規登録・重複通知(existing_notified)のどちらの経路にも入れず、
//     その行のapp_id/pass_urlをそのまま返す（メールは送らない・行も書き換えない）。
//     判定はemail+氏名判定より必ず先に行う。request_id が空（旧HTMLキャッシュ）なら
//     従来どおりの動作にフォールバックする。既存シートには migrateAddRequestIdColumn() が必要。
//   P2: `resendPass` の重複メール抑止。配送失敗が続くとクライアントが最大6回再試行し、
//     同一内容のパス再送メールが何通も届いていた。mail_logの直近10分以内に同じ宛先への
//     resend_pass送信成功記録があれば送信をスキップする（CacheServiceは使わず
//     シート参照のみ。project_gas_cache_lessonの教訓）。
//
// 🆕 v0.14.0（2026-08-07）: 新規申込のLINE WORKS通知を追加。
//   Incoming Webhook方式（認証不要・`{body:{text}}`形式。ref_lineworks_two_notify_paths）。
//   スクリプトプロパティ `LINEWORKS_WEBHOOK`（任意）にビューフェス専用ルーム宛のURLを
//   設定すると、新規申込（apply/applyLiffの新規登録時のみ・更新時は送らない）のたびに
//   サロン名・お名前・業態・電話を通知する。未設定なら何もしない（既存の運用に影響なし）。
//   失敗しても申込自体は成立済みのため、呼び出し側でtry/catchして握りつぶす。
//
// 🆕 v0.15.0（2026-08-15・優先項目④ 公開フォーム対策）: 認証なしの公開フォーム(`case 'apply'`)に
//   多層の歯止めを入れる。設計の正は `公開フォーム対策_実装方針.md`。
//   🔴 変更は applyApplication の中だけに閉じている。`doPost`のエントリ・
//   `_validateApplicationFields`（LIFFと共有）・`applyLiff`・`liff.html` には一切触れていない
//   （`LINE連携境界_調査レポート.md` §2「触ってはいけない3箇所」）。
//
//   ① ロック前ゲート `_publicFormGate()`: スプレッドシートに触る前・LockServiceを取る前に
//      I/Oゼロで弾く。スパムの実害はメール枠だけでなく、**スクリプトロックの占有による
//      受付停止**（正規の申込者がwaitLockのタイムアウトで弾かれ、さらに6回リトライして
//      負荷を増やす）でもあるため、ここで捨てられるかどうかが受付能力を左右する。
//      - 確定的な拒否: 入力長の上限・業態が6択にない（＝細工されたPOST）
//      - ボット兆候はスコア方式（ハニーポット+1／滞在3秒未満+1／備考にURL+1）で
//        **2点以上のときだけ**拒否する。単独の兆候で人間を弾かないための設計
//        （自動入力がハニーポットを埋める事故・備考にサロンのURLを書く人が実在するため）
//   ② メール枠の安全弁: `MailApp.getRemainingDailyQuota()` を見て、残50通で警告通知・
//      残20通で確認メールを停止する。**申込行は必ず書き、pass_urlも返す**ので、
//      申込者は画面で入場パスを受け取れる（メールは控え）。
//      🔴 行数ではなく実残枠を見るのは、確認メールの消費元が `applyApplication` だけでなく
//      `applyLiff`(L647)・`resendPass` にもあり、apply側で行数を数えても全体が見えないため。
//   ③ 同一宛先への確認メールは1日5通まで（本日分の行数で判定・追加のシート読み込みなし）
//   ④ `existing_notified` 経路の `_sendPassResendMail` に10分抑止を適用。
//      v0.13.0のP2で `resendPass` だけ塞いだ穴が、apply経由では開いたままだった
//      （他人のメール＋既知の氏名で無制限に再送させられる）
//   ⑤ 新規申込のLINE WORKS通知は本日30件を超えたら停止（1回だけ要約を通知）。
//      通知の洪水を防ぐと同時に、1件ごとのUrlFetchApp往復が応答時間を伸ばすのも止める
//
//   スクリプトプロパティ（すべて任意・未設定なら既定値）:
//     PUBLIC_FORM_GUARD   'off'にすると①②③⑤を全て無効化（再デプロイ不要のキルスイッチ）
//     MAIL_QUOTA_WARN     既定50 / MAIL_QUOTA_STOP 既定20
//     PER_EMAIL_MAIL_MAX  既定5  / DAILY_NOTIFY_LIMIT 既定30
//   🔴 新しいシートも新しいOAuthスコープも追加していない（MailAppは既に使用中）。
//   したがってデプロイ後の手動マイグレーションは不要。
//
// 🆕 v0.15.1（2026-08-15・名札印刷badges.html準備 S0）: 診断専用の `testAuthSheetAccess()` を追加。
//   `badges.html`（名札印刷）はbeaufield-authの共通セッションに相乗りする設計（名札印刷_badges設計.md）。
//   認証基盤(validateSession等)を実装する前に、beaufesのGAS実行アカウントが
//   beaufield-auth スプレッドシートを開けることをこの関数で確認する。既存の申込経路への影響なし。
//
// 🆕 v0.16.0（2026-08-15・名札印刷badges.html本体 S1〜S3）: `badges.html` から呼ばれる
//   読み取り専用アクション `listBadges`/`listSpareBadges` を追加。認証は
//   beaufield-authの共通セッション（`validateSession`をorder-appから移植・§5-2）。
//   `spare_badges`シート・configの`badge_*`キーは`setupSpareBadges()`/`seedBadgeConfig()`で
//   手動投入（新シート・新OAuthスコープはこれ以外に無し）。
//   🔴 変更は新しい関数と`doPost`の新しい`case`2行のみ。既存の申込経路
//  （apply/applyLiff/updateApplication/_validateApplicationFields/liff.html）は一切変更していない。
//   詳細・実装計画は `名札印刷_badges設計.md`（総チェック3周・25件の落とし穴を反映済み）。
// ============================================================

const VERSION  = '0.33.0';
const APP_NAME = 'beaufes';

// スクリプトプロパティから機密値を取得（コードへの直書き禁止）
const _PROPS         = PropertiesService.getScriptProperties();
const SPREADSHEET_ID = _PROPS.getProperty('SPREADSHEET_ID');

// 🆕 名札印刷badges.html用（名札印刷_badges設計.md §5-1）。
// beaufield-authの共通セッションに相乗りするためのスプレッドシートID。
const AUTH_SHEET_ID = _PROPS.getProperty('AUTH_SHEET_ID');
const CACHE_TTL_SESSION = 60; // 権限変更を最大1分で反映（order-appと同じ）

// シート名定数
const SHEET_APPLICATIONS       = 'applications';
const SHEET_SESSIONS           = 'sessions';
const SHEET_RESERVATIONS       = 'reservations';
const SHEET_CHECKINS           = 'checkins';
const SHEET_MAIL_LOG           = 'mail_log';
const SHEET_CONFIG             = 'config';
const SHEET_LINE_FRIENDS_CACHE = 'line_friends_cache'; // 🆕 L1-a（LINE Harness友だちのキャッシュ）
const SHEET_LINE_SYNC_LOG      = 'line_sync_log';       // 🆕 L2（プッシュ・タグ付与・metadata書き戻しの失敗ログ）
const SHEET_SPARE_BADGES       = 'spare_badges';        // 🆕 名札印刷（予備名札プール・§5-5）

// メール送信元（機密ではないため直書きでよい。§7-0-2で確定）
const MAIL_FROM_ADDR = 'beaufes@gmail.com';
const MAIL_FROM_NAME = '株式会社ビューフィールド ビューフェス事務局';

// GitHub Pagesの公開URL（QRコード・メール内のパスリンク生成に使用）
const SITE_BASE_URL = 'https://beaufield.github.io/beaufield-dev/beaufes/';

// 🆕 L1-a: LINE Harness連携用（機密値はScript Propertiesから取得。コードへの直書き禁止）
const LINE_HARNESS_API_URL = _PROPS.getProperty('LINE_HARNESS_API_URL');
const LINE_HARNESS_API_KEY = _PROPS.getProperty('LINE_HARNESS_API_KEY');
const BEAUFES_TAG_ID       = _PROPS.getProperty('BEAUFES_TAG_ID'); // タグ「ビューフェス2026申込」のtagId（秘密情報ではない）

// 🆕 L1-b: LIFF/LINEログインチャネルID（秘密情報ではないため直書きでよい。§0で発行済み）
const LIFF_CHANNEL_ID = '2010404613';

// 🆕 L1-c: 業態6択。index.html/pass.html/liff.htmlと1文字も違えないこと（LINEフォーム実物準拠）
const BUSINESS_TYPE_OPTIONS = ['美容室', '理容室', 'エステサロン', 'ネイルサロン', 'アイサロン', 'その他'];

// 🆕 v0.14.0: 新規申込のLINE WORKS通知（任意・未設定なら何もしない）。
// Incoming Webhook方式（[[ref_lineworks_two_notify_paths]]・認証不要・{body:{text}}形式）。
// ビューフェス専用ルーム宛に新規発行したURLをここに設定する。
const LINEWORKS_WEBHOOK = _PROPS.getProperty('LINEWORKS_WEBHOOK');

// ============================================================
// 起動時チェック（プロパティ未設定を早期検知）
// ============================================================
function _checkProps() {
  if (!SPREADSHEET_ID) throw new Error('スクリプトプロパティ SPREADSHEET_ID が未設定です');
}

// 🆕 L1-a: LINE Harness関連の関数だけが呼ぶ（apply/getPass等の基本機能はLINE連携が
// 未設定でも動き続けてほしいため、_checkPropsとは分離している）
function _checkLineHarnessProps() {
  if (!LINE_HARNESS_API_URL) throw new Error('スクリプトプロパティ LINE_HARNESS_API_URL が未設定です');
  if (!LINE_HARNESS_API_KEY) throw new Error('スクリプトプロパティ LINE_HARNESS_API_KEY が未設定です');
}

// ============================================================
// 🆕 名札印刷badges.html用の認証（S1・名札印刷_badges設計.md §5-2）
//
// beaufield-authの共通セッション（sessions→users→user_app_rolesの3シート）に相乗りする。
// order-app/gas/Code.gs:88-164 の validateSession をそのまま移植したもの
// （APP_NAMEとキャッシュキーだけ変更）。自分で考えた別方式にしないこと（§2-4）。
//
// 🔴 この節の関数は listBadges/listSpareBadges（＝badges.html関連アクション）からしか
//    呼ばない。既存の申込経路（apply/applyLiff/updateApplication等）には一切関与しない。
// ============================================================

// セッション検証。beaufield-authの sessions シートでトークンを照合する。
// CacheServiceで60秒キャッシュ（権限変更・ログアウトを最大1分で反映）。
// 戻り値: { valid:true, user_id, name, is_admin, role } または
//        { valid:false } または { valid:false, transient:true }（一時障害・負キャッシュしない）
function validateSession(token) {
  if (!token) return { valid: false, reason: 'no_token' };
  if (!AUTH_SHEET_ID) return { valid: false, transient: true }; // 未設定はプロパティ不備＝一時的な設定不足として扱う

  const cache    = CacheService.getScriptCache();
  const cacheKey = 'sess_beaufes_v1_' + token.slice(-32); // 🔴 他アプリと衝突させないプレフィックス
  const cached   = cache.get(cacheKey);
  if (cached !== null) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  try {
    const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);
    const sh = ss.getSheetByName('sessions');
    // シート取得失敗は「セッション無効」ではなく一時障害。負キャッシュしない
    // （Google側の一時的な応答不良でも起こりうるため。負キャッシュすると
    //  有効なトークンがTTLの間ブロックされ続けてしまう）
    if (!sh) return { valid: false, transient: true };

    const data = sh.getDataRange().getValues();
    const now  = Date.now();

    for (let i = 1; i < data.length; i++) {
      const rowToken   = String(data[i][0]);
      const rowUserId  = String(data[i][1]);
      const rowExpires = Number(data[i][2]);

      if (rowToken === token) {
        if (rowExpires < now) {
          // 期限切れ → 行を削除してから拒否
          sh.deleteRow(i + 1);
          const r = { valid: false, reason: 'expired' };
          cache.put(cacheKey, JSON.stringify(r), 60);
          return r;
        }
        const usersSh = ss.getSheetByName('users');
        const rolesSh = ss.getSheetByName('user_app_roles');
        if (!usersSh || !rolesSh) return { valid: false, transient: true };

        const users = usersSh.getDataRange().getValues();
        let userRow = null;
        for (let j = 1; j < users.length; j++) {
          if (String(users[j][0]) === rowUserId) { userRow = users[j]; break; }
        }
        if (!userRow || !(userRow[3] === true || userRow[3] === 'TRUE')) {
          return { valid: false, reason: 'inactive' };
        }

        const roles = rolesSh.getDataRange().getValues();
        let role = '';
        for (let j = 1; j < roles.length; j++) {
          if (String(roles[j][0]) === rowUserId && String(roles[j][1]) === APP_NAME) {
            role = String(roles[j][2] || '').trim().toLowerCase();
            break;
          }
        }
        if (!role || role === 'none') return { valid: false, reason: 'no_role' };

        const r = {
          valid: true,
          user_id: rowUserId,
          name: String(userRow[1] || rowUserId),
          is_admin: userRow[5] === true || userRow[5] === 'TRUE',
          role: role
        };
        cache.put(cacheKey, JSON.stringify(r), CACHE_TTL_SESSION);
        return r;
      }
    }
  } catch (e) {
    // 認証シートを読めなかった＝一時障害。ここも負キャッシュしない（上記と同じ理由）
    Logger.log('beaufes validateSession エラー: ' + e);
    return { valid: false, transient: true };
  }
  // ここに到達＝シートは読めたがトークンが見つからなかった＝本物の無効
  const r = { valid: false, reason: 'not_found' };
  cache.put(cacheKey, JSON.stringify(r), 60);
  return r;
}

// 認証必須アクションの入口で必ず呼ぶ。
// 🔴 例外を投げないこと。doPost の外側 catch(err) が全部 INTERNAL_ERROR に潰してしまい、
//    画面側が「未ログイン」と「サーバー障害」を区別できなくなるため（§5-2）。
// 戻り値: { ok:true, session } または { ok:false, error:'SESSION_INVALID'|'AUTH_TRANSIENT' }
function _requireSession(data) {
  const v = validateSession(String((data && data.session_token) || ''));
  if (v.valid) return { ok: true, session: v };

  // 認証シートを読めなかっただけの場合は「ログインし直せ」ではなく「一時障害」として返す。
  if (v.transient) return { ok: false, error: 'AUTH_TRANSIENT' };

  // 🆕 v0.18.0（2026-08-26）: 「未ログイン」と「権限なし」を区別して返す。
  // 従来は両方 SESSION_INVALID で、画面側で切り分けられず原因特定に時間がかかった
  // （名札印刷設計 §5-2 の「区別しない」方針をこの版で変更した）。
  // 使うのは社員専用画面のアクションだけなので、この粒度なら返してよいと判断している。
  //   NO_TOKEN       : リクエストに session_token が入っていない（未ログイン、または送り方のバグ）
  //   SESSION_INVALID: トークンが見つからない・期限切れ（ログインし直せば直る）
  //   USER_INACTIVE  : users シートで利用停止になっている
  //   NO_ROLE        : ログインは有効だが user_app_roles に beaufes の行がない（または none）
  switch (v.reason) {
    case 'no_token': return { ok: false, error: 'NO_TOKEN' };
    case 'inactive': return { ok: false, error: 'USER_INACTIVE' };
    case 'no_role':  return { ok: false, error: 'NO_ROLE' };
    default:         return { ok: false, error: 'SESSION_INVALID' }; // not_found / expired
  }
}

// ============================================================
// エントリーポイント（GET） ―― 全て認証なし・公開アクション
// ============================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  let data = {};
  try {
    if (e && e.parameter && e.parameter.data) data = JSON.parse(e.parameter.data);
  } catch (jsonErr) {
    return _jsonResponse(_err('INVALID_REQUEST'));
  }

  try {
    switch (action) {
      case 'getPass':   return _jsonResponse(getPass(data));
      // 🆕 v0.19.0 booth.html用。認証なし・booth_tokenで認可（booth実装設計_確定版.md §4-1）
      case 'boothInit': return _jsonResponse(boothInit(data));
      case 'listSessions': return _jsonResponse(listSessions(data));  // 🆕 v0.20.0 セミナー枠一覧（公開）
      // 🆕 診断用（diag.html）。読み取りのみ・データを一切変更しない
      case 'ping':      return _jsonResponse(pingLight(data));
      case 'pingHeavy': return _jsonResponse(pingHeavy(data));
      default:          return _jsonResponse(_err('不明なアクション: ' + action));
    }
  } catch (err) {
    Logger.log('doGet error: ' + err);
    return _jsonResponse(_err('INTERNAL_ERROR'));
  }
}

// ============================================================
// エントリーポイント（POST） ―― 全て認証なし・公開アクション
// ============================================================
function doPost(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action || '';
  let data = {};
  try {
    if (params.data) data = JSON.parse(params.data);
  } catch (jsonErr) {
    return _jsonResponse(_err('INVALID_REQUEST'));
  }

  try {
    switch (action) {
      case 'apply':             return _jsonResponse(applyApplication(data, params.client_attempt));
      case 'resendPass':        return _jsonResponse(resendPass(data));
      case 'updateApplication': return _jsonResponse(updateApplication(data));
      case 'liffPrefill':       return _jsonResponse(liffPrefill(data));
      case 'applyLiff':         return _jsonResponse(applyLiff(data));
      // 🆕 v0.20.0 セミナー予約（認証なし・ticket_tokenが本人確認を兼ねる）
      case 'listSessions':      return _jsonResponse(listSessions(data));
      case 'reserveSessions':   return _jsonResponse(reserveSessions(data));
      case 'listReservations':  return _jsonResponse(listReservations(data));   // 🆕 v0.31.0 社員用 🔒
      // 🆕 診断用（diag.html）。読み取りのみ・データを一切変更しない
      case 'ping':              return _jsonResponse(pingLight(data));
      case 'pingHeavy':         return _jsonResponse(pingHeavy(data));
      // 🆕 名札印刷badges.html用（S3）。🔒 認証必須（_requireSession・上記以外は認証なしのまま）
      case 'listBadges':        return _jsonResponse(listBadges(data));
      case 'listSpareBadges':   return _jsonResponse(listSpareBadges(data));
      // 🆕 v0.17.0 申込者一覧 admin.html 用。🔒 認証必須（_requireSession）
      case 'listApplications':  return _jsonResponse(listApplications(data));
      case 'setTantou':         return _jsonResponse(setTantou(data));
      // 🆕 v0.19.0 booth.html用（booth実装設計_確定版.md §4）。
      // 上2つは認証なし・booth_tokenで認可（社外端末から叩くため）。下4つは🔒認証必須（_requireSession）
      case 'boothSubmit':            return _jsonResponse(boothSubmit(data));
      case 'boothVoid':              return _jsonResponse(boothVoid(data));
      case 'boothRecent':            return _jsonResponse(boothRecent(data));
      case 'boothSummary':           return _jsonResponse(boothSummary(data));
      case 'boothResolveUnresolved': return _jsonResponse(boothResolveUnresolved(data));
      case 'boothExportCsv':         return _jsonResponse(boothExportCsv(data));
      case 'boothImportQueue':       return _jsonResponse(boothImportQueue(data));
      // 🆕 v0.26.0 受付 scan.html 用（当日運用_堅牢化設計.md §3）。🔒 3本とも認証必須（_requireSession）
      case 'scanRoster':             return _jsonResponse(scanRoster(data));
      case 'scanCheckin':            return _jsonResponse(scanCheckin(data));
      case 'scanCheckedIn':          return _jsonResponse(scanCheckedIn(data));
      // 🆕 v0.33.0 admin.html の当日機能（代理登録・予備名札の割当・パス再送）。🔒 認証必須
      case 'adminCreateApplication': return _jsonResponse(adminCreateApplication(data));
      case 'assignSpare':            return _jsonResponse(assignSpare(data));
      case 'adminResendPass':        return _jsonResponse(adminResendPass(data));
      default:                  return _jsonResponse(_err('不明なアクション: ' + action));
    }
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return _jsonResponse(_err('INTERNAL_ERROR'));
  }
}

// ============================================================
// ユーティリティ
// ============================================================
function _jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function _ok(data)  { return { success: true,  data: data }; }
function _err(msg)  { return { success: false, error: msg }; }
function _now() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}
function _getSheet(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('シートが見つかりません: ' + name + '（setupSheetsを実行してください）');
  return sh;
}
function _genToken() {
  // QR用ランダム16文字（推測不能・§3設計どおり）
  return Utilities.getUuid().replace(/-/g, '').substring(0, 16);
}
function _escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// 🆕 氏名の「照合キー」を作る（v0.7.0・§4-3）
// シートに保存する表示用の氏名は入力どおりのまま。照合のときだけこの関数を通す。
// 「山田太郎」「山田 太郎」「山田　太郎（全角空白）」を同一人物として扱うため。
// 本人は同じ表記のつもりでも空白の有無で別行ができてしまう事故を防ぐ。
function _normalizeName(s) {
  return String(s || '')
    .normalize('NFKC')     // 全角英数・全角空白・半角カナなどを正規形に統一
    .replace(/\s+/g, '')   // 空白を全て除去（半角・全角とも）
    .toLowerCase();        // ローマ字表記の大文字小文字の揺れを吸収
}

// 既存app_idの最大連番+1を発番（例: F2026-0007 → F2026-0008）
function _nextAppId(rows) {
  const year = new Date().getFullYear();
  let maxSeq = 0;
  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][0] || '');
    const m  = id.match(/-(\d+)$/);
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > maxSeq) maxSeq = n;
    }
  }
  const seq = maxSeq + 1;
  return 'F' + year + '-' + ('0000' + seq).slice(-4);
}

// config シートを key/value オブジェクトとして取得
function _getConfig() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh  = ss.getSheetByName(SHEET_CONFIG);
  const cfg = {};
  if (!sh) return cfg;
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) cfg[String(rows[i][0])] = rows[i][1];
  }
  return cfg;
}

// ============================================================
// 申込フォームの共通処理（applyApplication / updateApplication / applyLiff で共用）
// 🆕 L1-c: Web版とLIFF版でバリデーション・書き込みロジックを二重管理しないための切り出し
// ============================================================

// 入力値のバリデーション。opts.requireAgree=false のときはagree_capability必須チェックを省く
// （updateApplicationは編集画面に同意チェックボックスが無いため）。
// 戻り値: 成功時 { fields }／失敗時 { error }
function _validateApplicationFields(data, opts) {
  opts = opts || {};
  const f = {
    salonName:      String(data.salon_name || '').trim(),
    staffName:      String(data.staff_name || '').trim(),
    email:          String(data.email || '').trim(),
    phone:          String(data.phone || '').trim(),
    businessType:   String(data.business_type || '').trim(), // U列。名札の色分けに使用（§4-1-2）
    hasTransaction: String(data.has_transaction || '').trim(), // 'yes' / 'no'
    address:        String(data.address || '').trim(),        // 新規客のみ
    referrer:       String(data.referrer || '').trim(),       // 新規客のみ
    note:           String(data.note || '').trim(),
    agree:          !!data.agree_capability
  };

  if (!f.salonName) return { error: 'サロン名を入力してください' };
  if (!f.staffName) return { error: 'お名前を入力してください' };
  if (!f.email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(f.email)) return { error: 'メールアドレスの形式が正しくありません' };
  if (!f.phone) return { error: 'お電話番号を入力してください' };
  if (!f.businessType) return { error: '業態を選択してください' };
  if (opts.requireAgree && !f.agree) return { error: '美容従事者であることの確認にチェックしてください' };

  f.emailNorm = f.email.toLowerCase();
  return { fields: f };
}

// 新規行を追加する。source: 'web' | 'liff'。lineFriendId/lineUserId はWeb版では常に''。
// requestId は冪等キー（v0.13.0・§2-2）。applyLiffからは渡されない（LIFFはline_user_idで
// 既に冪等なため不要）→ その場合は省略され''になる。
// 戻り値: { appId, ticketToken }
function _appendApplicationRow(sh, rows, f, source, lineFriendId, lineUserId, requestId) {
  const appId       = _nextAppId(rows);
  const ticketToken = _genToken();
  const now         = _now();
  const newRow      = rows.length + 1;

  const values = [
    appId, now, now, source,
    f.salonName, f.staffName, f.email, f.emailNorm, f.phone, '', // J列(area)はフォームから削除済み（2026-08-06）
    f.hasTransaction, f.address, f.referrer, f.agree,
    lineFriendId || '', lineUserId || '',
    ticketToken, 'confirmed', '', f.note,
    f.businessType, // U列（§4-1-2）
    requestId || '' // 🆕 V列（v0.13.0・§2-2 冪等キー）
  ];
  // 🔴 電話番号列(I列=9列目)は書き込み前に必ずPlain Text指定する（migrateFixPhoneColumnと同じ理由）。
  // appendRowだと書式を挟めないため、行番号を自分で計算してsetValuesに置き換えている。
  sh.getRange(newRow, 9, 1, 1).setNumberFormat('@');
  sh.getRange(newRow, 1, 1, values.length).setValues([values]);

  return { appId: appId, ticketToken: ticketToken };
}

// 既存行を更新する。app_id・ticket_token・statusは書き換えない。
function _updateApplicationRow(sh, row, f) {
  const now = _now();
  // 🔴 電話番号列(I列=9列目)は書き込み前に必ずPlain Text指定する（migrateFixPhoneColumnと同じ理由）
  sh.getRange(row, 9, 1, 1).setNumberFormat('@');
  sh.getRange(row, 3, 1, 1).setValue(now); // updated_at
  sh.getRange(row, 5, 1, 9).setValues([[
    // J列(area)はフォームから削除済み（2026-08-06）。列位置を保つため空文字を書き続ける
    f.salonName, f.staffName, f.email, f.emailNorm, f.phone, '', f.hasTransaction, f.address, f.referrer
  ]]);
  sh.getRange(row, 20, 1, 1).setValue(f.note);
  sh.getRange(row, 21, 1, 1).setValue(f.businessType); // U列（§4-1-2）
}

// ============================================================
// 🆕 v0.15.0: 公開フォーム専用のガード（優先項目④）
//
// 🔴 この節の関数は applyApplication（＝`case 'apply'`）からしか呼ばない。
//    applyLiff・updateApplication・liff.html からは絶対に呼ばないこと。
//    LINE経由の申込はIDトークンで本人確認済みであり、ここで想定している
//    「認証なしの公開フォームへの機械的な投稿」とは前提がまったく違うため。
// ============================================================

// 入力長の上限。用途は「50KBの文字列を投げ込まれてシートを膨らませる」ことの防止であって
// 入力内容の整形ではない。したがって**実用上ぶつからない寛容な値**にしてある
// （厳しくすると、少し長めに書いた正規の申込者を弾く side effect のほうが確実に大きい）。
const PUBLIC_FIELD_MAX = {
  salonName: 100, staffName: 60, email: 254, phone: 40,
  address: 200, referrer: 100, note: 2000
};

// ハニーポットの入力欄名。index.html 側と一致させること。
// 🔴 `website` `company` `fax` 等の「意味のある名前」は使わない。
// ブラウザやパスワードマネージャーの自動入力が拾って、人間の申込を誤爆させる実例があるため。
const PUBLIC_HONEYPOT_FIELD = 'bf_note2';

// ガード設定。getProperties()の1回呼び出しで済ませる（getPropertyを個別に叩くより安い）。
function _guardConfig() {
  const p = _PROPS.getProperties();
  function num(key, def) {
    const v = Number(p[key]);
    return (isNaN(v) || v < 0) ? def : v;
  }
  return {
    enabled:         String(p['PUBLIC_FORM_GUARD'] || 'on').toLowerCase() !== 'off',
    mailQuotaWarn:   num('MAIL_QUOTA_WARN',    50),
    mailQuotaStop:   num('MAIL_QUOTA_STOP',    20),
    perEmailMailMax: num('PER_EMAIL_MAIL_MAX',  5),
    dailyNotifyMax:  num('DAILY_NOTIFY_LIMIT', 30)
  };
}

// ロック前ゲート。**スプレッドシートに一切触らない**（＝ロックも実行時間も消費させない）。
// 戻り値: null＝通過 ／ 文字列＝拒否理由（呼び出し側が _err で返す）
//
// 🔴 拒否メッセージは index.html の TRANSPORT_ERROR_PATTERN（`不明なアクション|INVALID_REQUEST`）に
// 絶対にマッチさせないこと。マッチすると輸送路の破損とみなされ、1回の拒否が6回のリトライに増幅する。
function _publicFormGate(data, f) {
  // --- (1) 確定的な拒否 ---------------------------------------------------
  // 画面のプルダウンでは選べない値＝手で組み立てたPOST。ここは兆候ではなく確定とみなす。
  if (BUSINESS_TYPE_OPTIONS.indexOf(f.businessType) < 0) {
    return '業態を選択してください';
  }
  const tooLong =
    (f.salonName.length > PUBLIC_FIELD_MAX.salonName) ||
    (f.staffName.length > PUBLIC_FIELD_MAX.staffName) ||
    (f.email.length     > PUBLIC_FIELD_MAX.email)     ||
    (f.phone.length     > PUBLIC_FIELD_MAX.phone)     ||
    (f.address.length   > PUBLIC_FIELD_MAX.address)   ||
    (f.referrer.length  > PUBLIC_FIELD_MAX.referrer)  ||
    (f.note.length      > PUBLIC_FIELD_MAX.note);
  if (tooLong) {
    return '入力が長すぎる項目があります。お手数ですが短くしてご入力ください';
  }

  // --- (2) ボット兆候のスコア ----------------------------------------------
  // 🔴 1つでも当たったら拒否、にはしない。人間が誤って1つ踏むことは現実に起きるが
  // （自動入力・備考にサロンのURLを書く等）、2つ同時に踏むことは実質起きないため。
  // 🔴 いずれの兆候も「キー自体が無ければ加点しない」。旧HTMLキャッシュや
  // liff.htmlからの避難経路（フィールドを持たない）を素通しさせるための後方互換
  //（`request_id` が空でも動くのと同じ考え方）。
  let score = 0;
  const hits = [];

  if (String(data[PUBLIC_HONEYPOT_FIELD] || '').trim() !== '') {
    score++; hits.push('honeypot');
  }
  const elapsedMs = Number(data.elapsed_ms);
  if (!isNaN(elapsedMs) && elapsedMs >= 0 && elapsedMs < 3000) {
    score++; hits.push('too_fast:' + elapsedMs + 'ms');
  }
  if (/https?:\/\/|www\./i.test(f.note)) {
    score++; hits.push('url_in_note');
  }

  if (score >= 2) {
    Logger.log('publicFormGate: ボット判定で拒否 score=' + score + ' hits=' + hits.join(',') +
               ' email=' + f.email + ' name=' + f.staffName);
    // 理由は明かさない（ボットに学習させないため）。
    // ただし人間が誤爆した場合の逃げ道として、LINEからの申込を案内する。
    return '送信できませんでした。お手数ですが時間をおいて再度お試しいただくか、' +
           'LINE公式アカウントからお申し込みください';
  }
  if (score > 0) {
    // 1点は通す。実際に何が引っかかっているかを後から見られるようにログだけ残す。
    Logger.log('publicFormGate: 兆候1点（通過） hits=' + hits.join(',') + ' email=' + f.email);
  }
  return null;
}

// created_at セルが「今日」かどうか。
// 🔴 setValues で 'yyyy-MM-dd HH:mm:ss' 文字列を書いても、Sheetsが日時値として
// 解釈して Date オブジェクトで返してくる場合がある。両方の表現を受けられるようにする。
function _isTodayCell(v, todayStr) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd') === todayStr;
  return String(v).slice(0, 10) === todayStr;
}

// 残りのメール送信枠。取得に失敗したら null を返す（＝呼び出し側はフェイルオープン）。
// 🔴 フェイルオープンにするのは、一時的な取得失敗で確認メールを全部止めるほうが
// 実害が大きいため。止めるのは「残枠が確かに少ないと分かったとき」だけにする。
function _remainingMailQuota() {
  try {
    return MailApp.getRemainingDailyQuota();
  } catch (e) {
    Logger.log('getRemainingDailyQuota失敗（フェイルオープン）: ' + e);
    return null;
  }
}

// LINE WORKSへ任意の文字列を通知する。
// 🔴 既存の `_notifyNewApplicationLineWorks` をリファクタして共用しないのは、
// あの関数が applyLiff からも呼ばれており、LINE経路に触れないという今回の大前提を
// 守るため。5行の重複は、LINE申込を巻き込むリスクより安い。
function _notifyLineWorksText(text) {
  if (!LINEWORKS_WEBHOOK) return;
  UrlFetchApp.fetch(LINEWORKS_WEBHOOK, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ body: { text: text } }),
    muteHttpExceptions: true
  });
}

// 同じ種類の警告は1日1回しか送らない。日付はスクリプトプロパティに置く
// （シートI/Oを増やさないため。CacheServiceは使わない＝project_gas_cache_lessonの教訓）。
// 🔴 同時実行で2通出ることはありうるが、警告が重複するだけなので許容する。
function _alertOncePerDay(propKey, text) {
  try {
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    if (_PROPS.getProperty(propKey) === today) return;
    _PROPS.setProperty(propKey, today);
    _notifyLineWorksText(text);
  } catch (e) {
    Logger.log('_alertOncePerDay失敗（無視して継続）: ' + e);
  }
}

// ============================================================
// ① 申込受付（doPost: action=apply）
//    キーは email_norm + 正規化した staff_name（§4-3）
//
//    🆕 v0.8.0（L0-X）: 確認画面は廃止。以下の3分岐のみ。
//      - メールに一致なし               → そのまま新規登録
//      - メール一致 かつ 正規化氏名も一致 → 既存行は一切変更せず・情報も一切返さず、
//                                        パスURLをメール再送するだけ（{existing_notified:true}）
//      - メール一致 だが氏名が違う       → 別人として新規登録（代表メールでの複数名申込に対応）
//    既存申込の内容を変更する経路は updateApplication（ticket_token方式）のみに一本化。
//
//    🆕 v0.13.0（配送障害対策P1・§2-2）: 冪等キー(`request_id`)による再試行の吸収。
//    GAS結果配送層が失敗すると、クライアントは「サーバーに届いたかどうか分からない」まま
//    自動リトライする。従来はこの再試行が上記の「メール一致・氏名一致」経路に落ちて
//    existing_notified を返していた＝初めて申込んだお客様が「すでにお申し込みをお預かりして
//    います」と言われる不具合があった（GAS配送障害_調査レポート.md §7-1）。
//    request_id はクライアントがページ読み込み時に1つ生成し、全リトライで同じ値を送る。
//    同じrequest_idの行が既にあれば「配送失敗による再試行」と確定できるので、
//    email+氏名判定より必ず先に見て、新規登録もexisting_notifiedもどちらにも入らせず
//    その場でapp_id/pass_urlを返す（メール再送はしない・行も書き換えない）。
// ============================================================
function applyApplication(data, clientAttempt) {
  _checkProps();

  const v = _validateApplicationFields(data, { requireAgree: true });
  if (v.error) return _err(v.error);
  const f = v.fields;
  const nameKey    = _normalizeName(f.staffName);
  const requestId  = String(data.request_id || '').trim().slice(0, 64);

  // 🆕 v0.15.0: ロック前ゲート。ここで弾けたリクエストは
  // スプレッドシートにもLockServiceにも到達しない（＝受付能力を消費しない）。
  const cfg = _guardConfig();
  if (cfg.enabled) {
    const gateErr = _publicFormGate(data, f);
    if (gateErr) return _err(gateErr);
  }

  if (clientAttempt && Number(clientAttempt) >= 2) {
    Logger.log('apply: リトライ経由の到達 client_attempt=' + clientAttempt + ' request_id=' + requestId);
  }

  let appId, ticketToken;
  // 🆕 v0.20.0: セミナー予約。ここで例外を出さない・申込を止めないこと（設計書§4-2の絶対条件）
  const wantSessions = _parseSessionIds(data.sessions);
  let sessionsResult = null;
  let replayed          = false; // 🆕 request_id一致＝配送失敗による再試行と確定した場合
  let existingNotified  = false;
  let resendList        = null; // メール一致・氏名一致が見つかった場合の再送先（ロック解放後に送信）

  // 🆕 v0.15.0: 本日分の集計（既存の行スキャンの中で数えるので追加のシート読み込みはゼロ）
  const todayStr       = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  let todayNewCount    = 0; // 本日作成された申込行の数
  let todayMailToEmail = 0; // 本日この宛先で作成された行の数（＝送った確認メール数の近似）

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh  = _getSheet(ss, SHEET_APPLICATIONS);
    const rows = sh.getDataRange().getValues();

    // 🆕 request_id が一致する既存行を最優先で探す（空文字同士は誤マッチしないよう対象外）。
    if (requestId) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][21]) !== requestId) continue;
        replayed    = true;
        appId       = rows[i][0];
        ticketToken = rows[i][16];
        break;
      }
    }

    if (!replayed) {
      // 同じメール＋同じ氏名（正規化）の既存申込を探す。見つかっても内容は一切書き換えない・
      // 一切返さない（他人のメールで氏名・サロン名が引ける経路を作らないため・§2-2）。
      // 🆕 v0.15.0: 同じ1回のスキャンで本日分の件数も数えるため、一致を見つけても break せず
      // 最後まで回す（rowsは既にメモリ上にあるので、コストはCPU上の走査だけ）。
      for (let i = 1; i < rows.length; i++) {
        if (_isTodayCell(rows[i][1], todayStr)) {
          todayNewCount++;
          if (String(rows[i][7]).toLowerCase() === f.emailNorm) todayMailToEmail++;
        }

        if (existingNotified) continue;                    // 既に見つけている（先頭一致を採用）
        if (String(rows[i][7]).toLowerCase() !== f.emailNorm) continue;
        if (String(rows[i][17]) === 'cancelled') continue; // 取消済みは既存として扱わない
        if (_normalizeName(rows[i][5]) !== nameKey) continue;

        existingNotified = true;
        resendList = [{
          salonName: String(rows[i][4]),
          staffName: String(rows[i][5]),
          passUrl:   SITE_BASE_URL + 'pass.html?t=' + String(rows[i][16])
        }];
      }

      if (!existingNotified) {
        // 新規登録（メールが一致しても氏名が違えば別人として登録＝同一サロン複数名の代表メール申込に対応）
        const created = _appendApplicationRow(sh, rows, f, 'web', '', '', requestId);
        appId       = created.appId;
        ticketToken = created.ticketToken;

        // 🆕 v0.20.0: セミナー予約を同じロックの中で書く。
        // 🔴 満席・枠切れでも申込は成立させる（_writeReservationsは例外を投げない）。
        // 🔴 予約の書き込みが落ちても申込行は残す。ここでthrowすると「申込できない」に化ける。
        if (wantSessions.length) {
          try {
            sessionsResult = _writeReservations(ss, appId, wantSessions,
                               { replace: false, cancelledAppIds: _cancelledAppIdSet(rows) });
          } catch (e) {
            Logger.log('セミナー予約の書き込みに失敗（申込自体は成立済み）: ' + e);
            sessionsResult = { reserved: [], full: [], invalid: wantSessions.slice(), cancelled: [], error: String(e) };
          }
        }
      }
    }
  } finally {
    lock.releaseLock();
  }

  if (replayed) {
    // 配送失敗による再試行と確定済み。1回目の実行で確認メールは送信済みのため再送しない。
    return _ok({ app_id: appId, pass_url: SITE_BASE_URL + 'pass.html?t=' + ticketToken, is_update: false, replayed: true });
  }

  if (existingNotified) {
    // 🆕 v0.15.0: ここに10分抑止が無かったため、他人のメールアドレスと既知の氏名さえ分かれば
    // 無制限にパス再送メールを送りつけられた（v0.13.0のP2は`resendPass`側だけを塞いでいた）。
    // 応答は送信の有無にかかわらず同じにする（送ったかどうかを外から観測させない）。
    if (!cfg.enabled || !_recentResendMailSent(f.email)) {
      _sendPassResendMail(f.email, resendList);
    } else {
      Logger.log('apply: existing_notified のメールを10分抑止でスキップ email=' + f.email);
    }
    return _ok({ existing_notified: true });
  }

  const passUrl = SITE_BASE_URL + 'pass.html?t=' + ticketToken;

  // 🆕 v0.15.0: 確認メールの安全弁。
  // 🔴 どの分岐に落ちても「申込行は書けている・pass_url は返す」ことは変えない。
  // メールが送れないことより、申込そのものが通らないことのほうが害が大きいため。
  let mailSkipped = '';
  if (cfg.enabled) {
    const quota = _remainingMailQuota(); // 取得失敗時は null＝フェイルオープン
    if (todayMailToEmail >= cfg.perEmailMailMax) {
      mailSkipped = 'per_email_cap';
    } else if (quota !== null && quota <= cfg.mailQuotaStop) {
      mailSkipped = 'quota_stop';
      _alertOncePerDay('ALERT_MAIL_QUOTA_STOP_DATE',
        '🛑 ビューフェス申込: 残メール枠が' + quota + '通になったため、確認メールの自動送信を停止しました。\n' +
        '申込の受付自体は継続しており、申込者には画面で入場パスが表示されています。\n' +
        '（枠は翌日リセット。急ぎの場合は applications シートの pass_url を手動で送付してください）');
    } else if (quota !== null && quota <= cfg.mailQuotaWarn) {
      _alertOncePerDay('ALERT_MAIL_QUOTA_WARN_DATE',
        '⚠️ ビューフェス申込: 残メール枠が' + quota + '通です（1日の上限は100通）。\n' +
        '残' + cfg.mailQuotaStop + '通で確認メールの自動送信を停止します。');
    }
  }

  if (mailSkipped) {
    Logger.log('apply: 確認メールをスキップ reason=' + mailSkipped + ' email=' + f.email + ' app_id=' + appId);
  } else {
    _sendConfirmationMail(f.email, f.salonName, f.staffName, passUrl, false,
                          _reservedSessionLabels(appId),    // 🆕 v0.20.0 予約した枠を本文に載せる
                          !!(sessionsResult && sessionsResult.full && sessionsResult.full.length));
  }

  // 🆕 v0.15.0: 1件ごとの通知は本日30件までにする。
  // 通知の洪水で本物を見落とすのを防ぐと同時に、1件ごとのUrlFetchApp往復が
  // 応答時間を伸ばして配送404を増やすのも止める（_health の知見）。
  // 🔴 applyLiff 側の通知には手を入れない（LINE申込は本人確認済みで洪水にならないため）。
  if (!cfg.enabled || todayNewCount < cfg.dailyNotifyMax) {
    try {
      _notifyNewApplicationLineWorks(appId, f, 'web');
    } catch (e) {
      Logger.log('LINE WORKS通知に失敗（申込自体は成立済み）: ' + e);
    }
  } else {
    _alertOncePerDay('ALERT_NOTIFY_MUTED_DATE',
      '🔕 ビューフェス申込: 本日のWeb申込が' + cfg.dailyNotifyMax + '件を超えたため、' +
      '1件ごとの新規申込通知を本日ぶんは停止しました。\n' +
      '申込の受付は継続しています。件数が想定外に多い場合は applications シートをご確認ください。');
  }

  const res = { app_id: appId, pass_url: passUrl, is_update: false };
  if (mailSkipped) res.mail_skipped = mailSkipped; // 診断用。クライアントは参照していない
  // 🆕 v0.20.0: 満席等で予約できなかった枠を画面に伝える（申込自体は成立している）
  if (sessionsResult) res.sessions_result = sessionsResult;
  return _ok(res);
}

// ============================================================
// 🆕 ④ 申込内容の変更（doPost: action=updateApplication）―― v0.8.0（L0-X・§2-2）
//    ticket_token を知っている本人だけが自分の申込を編集できる（capability URL方式）。
//    pass.html の編集UI（L1-e）から呼ばれる。Web版の変更経路はこれのみ。
//
//    🔴 ticket_token（列17）・app_id（列1）・status（列18）は絶対に書き換えない。
//    🔴 重複判定は行わない（token を持っている時点で本人の編集と確定しているため）。
// ============================================================
function updateApplication(data) {
  _checkProps();

  const token = String(data.ticket_token || '').trim();
  if (!token) return _err('INVALID_TOKEN');

  const v = _validateApplicationFields(data, { requireAgree: false });
  if (v.error) return _err(v.error);
  const f = v.fields;

  let appId;
  // 🆕 v0.20.0: data.sessions が来ていればセミナー予約も差し替える。
  // 🔴 キーが無い場合（旧HTMLキャッシュからの送信）は予約に一切触らない。
  // undefined を「全部取消」と解釈すると、古い画面から編集しただけで予約が黙って消える。
  const touchSessions = (data.sessions !== undefined && data.sessions !== null && data.sessions !== '');
  const wantSessions  = touchSessions ? _parseSessionIds(data.sessions) : [];
  let sessionsResult  = null;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh  = _getSheet(ss, SHEET_APPLICATIONS);
    const rows = sh.getDataRange().getValues();

    let foundRow = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][16]) === token) { foundRow = i + 1; break; }
    }
    if (foundRow < 0) return _err('NOT_FOUND');

    appId = rows[foundRow - 1][0];
    _updateApplicationRow(sh, foundRow, f);

    if (touchSessions) {
      try {
        sessionsResult = _writeReservations(ss, appId, wantSessions,
                           { replace: true, cancelledAppIds: _cancelledAppIdSet(rows) });
      } catch (e) {
        Logger.log('セミナー予約の書き込みに失敗（申込内容の更新は成立済み）: ' + e);
        sessionsResult = { reserved: [], full: [], invalid: wantSessions.slice(), cancelled: [], error: String(e) };
      }
    }
  } finally {
    lock.releaseLock();
  }

  const upRes = { app_id: appId, pass_url: SITE_BASE_URL + 'pass.html?t=' + token };
  if (sessionsResult) upRes.sessions_result = sessionsResult;
  return _ok(upRes);
}

// ============================================================
// 🆕 L1-c: LIFF申込のプリフィル取得（doPost: action=liffPrefill）
//    IDトークンを検証し、line_friends_cacheと既存申込（line_user_id）から埋められる項目を返す。
//    友だちが見つからなくても success:true（空のフォームを出す。エラーにしない・§2-9）。
// ============================================================
function liffPrefill(data) {
  _checkProps();

  const idToken = String(data.id_token || '').trim();
  let lineUserId;
  try {
    lineUserId = _verifyLineIdToken(idToken);
  } catch (e) {
    return _err('INVALID_ID_TOKEN');
  }

  const friend = _findFriendByLineUserId(lineUserId);
  const result = { found: !!friend };

  if (friend) {
    result.friend_id = friend.friendId;

    const prefill = {};
    if (friend.salonName) prefill.salon_name = friend.salonName;
    if (friend.staffName) prefill.staff_name = friend.staffName;
    if (friend.phone)     prefill.phone = friend.phone;
    if (friend.email)     prefill.email = friend.email;
    if (friend.businessType && BUSINESS_TYPE_OPTIONS.indexOf(friend.businessType) >= 0) {
      prefill.business_type = friend.businessType; // 🔴 6択に無い値(旧表記「エステ」等)は省略（§1-4 地雷5）
    }
    result.prefill = prefill;
    result.has_transaction = friend.extId ? 'yes' : 'no';
  }

  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh  = _getSheet(ss, SHEET_APPLICATIONS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][15]) === lineUserId && String(rows[i][17]) !== 'cancelled') {
      result.existing_application = { app_id: String(rows[i][0]), ticket_token: String(rows[i][16]) };
      break;
    }
  }

  return _ok(result);
}

// ============================================================
// 🆕 L1-c: LIFF申込の確定（doPost: action=applyLiff）―― §2-10
//    line_user_id で既存申込が引ければそのまま更新（本人確定なのでメール送信・確認画面なし）。
//    引けなければ新規登録（source='liff'・line_user_id/line_friend_idを保存）。
//    duplicate_found / mode / target_app_id は存在しない（L0-Xで廃止済み。Web版と同じ）。
//    🔴 §2-2のメール＋氏名一致ロジックはLIFFでは使わない（line_user_idが本人確認そのものなので不要）。
// ============================================================
function applyLiff(data) {
  _checkProps();

  const idToken = String(data.id_token || '').trim();
  let lineUserId;
  try {
    lineUserId = _verifyLineIdToken(idToken);
  } catch (e) {
    return _err('INVALID_ID_TOKEN');
  }

  const v = _validateApplicationFields(data, { requireAgree: true });
  if (v.error) return _err(v.error);
  const f = v.fields;

  const friend       = _findFriendByLineUserId(lineUserId);
  const lineFriendId = friend ? friend.friendId : '';

  let appId, ticketToken, isUpdate;
  // 🆕 v0.20.0: セミナー予約。LIFFは本人確定なので、再申込＝内容変更として予約も差し替える
  const wantSessions = _parseSessionIds(data.sessions);
  let sessionsResult = null;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh  = _getSheet(ss, SHEET_APPLICATIONS);
    const rows = sh.getDataRange().getValues();

    let foundRow = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][15]) === lineUserId && String(rows[i][17]) !== 'cancelled') { foundRow = i + 1; break; }
    }

    if (foundRow > 0) {
      isUpdate    = true;
      appId       = rows[foundRow - 1][0];
      ticketToken = rows[foundRow - 1][16];
      _updateApplicationRow(sh, foundRow, f);
    } else {
      isUpdate = false;
      const created = _appendApplicationRow(sh, rows, f, 'liff', lineFriendId, lineUserId);
      appId       = created.appId;
      ticketToken = created.ticketToken;
    }

    // 🔴 申込を止めない（設計書§4-2）。satisfiedでなくても申込行は既に書けている。
    try {
      sessionsResult = _writeReservations(ss, appId, wantSessions,
                         { replace: true, cancelledAppIds: _cancelledAppIdSet(rows) });
    } catch (e) {
      Logger.log('セミナー予約の書き込みに失敗（申込自体は成立済み）: ' + e);
      sessionsResult = { reserved: [], full: [], invalid: wantSessions.slice(), cancelled: [], error: String(e) };
    }
  } finally {
    lock.releaseLock();
  }

  const passUrl = SITE_BASE_URL + 'pass.html?t=' + ticketToken;
  // 更新時はメールを送らない（本人が画面上でパスURLをそのまま受け取れるため。updateApplicationと同じ考え方）
  if (!isUpdate) {
    _sendConfirmationMail(f.email, f.salonName, f.staffName, passUrl, false,
                          _reservedSessionLabels(appId),    // 🆕 v0.20.0
                          !!(sessionsResult && sessionsResult.full && sessionsResult.full.length));
    try {
      _notifyNewApplicationLineWorks(appId, f, 'liff');
    } catch (e) {
      Logger.log('LINE WORKS通知に失敗（申込自体は成立済み）: ' + e);
    }
  }

  // 🆕 v0.24.0: 予約が変わったときの控えメール（更新の場合も送る）。
  // 🔴 新規申込のときは申込完了メールに予約が載っているので、二重に送らない。
  if (isUpdate && sessionsResult &&
      ((sessionsResult.added && sessionsResult.added.length) ||
       (sessionsResult.cancelled && sessionsResult.cancelled.length))) {
    try {
      _sendReservationMail(f.email, f.salonName, f.staffName, passUrl,
                           _reservedSessionLabels(appId),
                           !!(sessionsResult.full && sessionsResult.full.length));
    } catch (e) {
      Logger.log('予約の控えメール送信に失敗（申込・予約自体は成立済み）: ' + e);
    }
  }

  // 🆕 L2: LINEプッシュ・タグ付与・metadata書き戻し。
  // 🔴 ここで何が起きても申込自体は既に成立済み（スプレッドシートには書けている）なので、
  // 予期せぬ例外もここで握りつぶす（各ステップ内部でも個別にtry/catchしている・§2 L2冒頭）。
  try {
    _syncApplicationToLine(appId, lineFriendId, isUpdate, f, passUrl);
  } catch (e) {
    Logger.log('_syncApplicationToLine予期せぬ失敗: ' + e);
  }

  const liffRes = { app_id: appId, pass_url: passUrl, is_update: isUpdate };
  if (sessionsResult) liffRes.sessions_result = sessionsResult;
  return _ok(liffRes);
}

// ============================================================
// 🆕 L2: LINE側の仕上げ（申込完了プッシュ・タグ付与・metadata書き戻し）
//
// 🔴 これらが失敗しても申込自体を失敗させない。スプレッドシートに書けた時点で申込は
// 成立している。各ステップはtry/catchで独立させ、失敗は line_sync_log に記録して
// 後から手動で復旧できるようにする。
// ============================================================

function _ensureLineSyncLogSheet(ss) {
  let sh = ss.getSheetByName(SHEET_LINE_SYNC_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LINE_SYNC_LOG);
    sh.getRange(1, 1, 1, 6).setValues([[
      'logged_at', 'app_id', 'friend_id', 'step', 'status', 'error'
    ]]);
    sh.setFrozenRows(1);
    Logger.log('line_sync_logシート作成完了');
  }
  return sh;
}

function _logLineSync(appId, friendId, step, status, error) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = _ensureLineSyncLogSheet(ss);
    sh.appendRow([_now(), appId, friendId, step, status, error || '']);
  } catch (e) {
    Logger.log('line_sync_log書き込み失敗: ' + e);
  }
}

// friendIdで1件取得。キャッシュではなく最新値（metadata書き戻しの空欄判定に使う・§2-4の
// 「既に値が入っているキーは絶対に送らない」を守るため、数時間古いキャッシュに頼らない）。
function _lhGetFriend(friendId) {
  const json = _lhReq('GET', '/api/friends/' + encodeURIComponent(friendId), null);
  if (!json || json.success !== true || !json.data) {
    throw new Error('friend取得に失敗しました: ' + JSON.stringify(json).substring(0, 300));
  }
  return json.data;
}

function _lhAddTag(friendId, tagId) {
  const json = _lhReq('POST', '/api/friends/' + encodeURIComponent(friendId) + '/tags', { tagId: tagId });
  if (!json || json.success !== true) {
    throw new Error('タグ付与に失敗しました: ' + JSON.stringify(json).substring(0, 300));
  }
}

function _lhPutMetadata(friendId, patch) {
  const json = _lhReq('PUT', '/api/friends/' + encodeURIComponent(friendId) + '/metadata', patch);
  if (!json || json.success !== true) {
    throw new Error('metadata更新に失敗しました: ' + JSON.stringify(json).substring(0, 300));
  }
}

function _lhSendMessage(friendId, content) {
  const json = _lhReq('POST', '/api/friends/' + encodeURIComponent(friendId) + '/messages',
    { messageType: 'text', content: content });
  if (!json || json.success !== true) {
    throw new Error('メッセージ送信に失敗しました: ' + JSON.stringify(json).substring(0, 300));
  }
}

// 申込確定後のLINE側の仕上げ本体。friendIdが無い（友だちでない・キャッシュに未登録）場合は
// 何もしない（§2-9。エラーにしない）。各ステップは独立して失敗を許容する。
function _syncApplicationToLine(appId, friendId, isUpdate, f, passUrl) {
  if (!friendId) return;

  // ① 申込完了をLINEプッシュ（新規登録時のみ。更新時はメールと同じく送らない）
  if (!isUpdate) {
    try {
      const cfg       = _getConfig();
      const eventDate = cfg.event_date || '2026年10月26日（月）';
      const eventTime = cfg.event_time || '10:00〜16:00';
      const venueName = cfg.venue_name || '青島屋（AOSHIMAYA）';
      const venueAddr = cfg.venue_addr || '宮崎市青島2丁目12-11';
      const content =
        'ビューフェス2026へのお申し込みを受付いたしました。\n\n' +
        '日時: ' + eventDate + ' ' + eventTime + '\n' +
        '会場: ' + venueName + ' ' + venueAddr + '\n\n' +
        '▼入場パス\n' + passUrl;
      _lhSendMessage(friendId, content);
      _logLineSync(appId, friendId, 'push', 'ok', '');
    } catch (e) {
      _logLineSync(appId, friendId, 'push', 'error', String(e));
    }
  }

  // ② 「ビューフェス2026申込」タグ付与（新規・更新とも試す。既に付与済みでも害はない前提）
  try {
    if (!BEAUFES_TAG_ID) throw new Error('スクリプトプロパティ BEAUFES_TAG_ID が未設定です');
    _lhAddTag(friendId, BEAUFES_TAG_ID);
    _logLineSync(appId, friendId, 'tag', 'ok', '');
  } catch (e) {
    _logLineSync(appId, friendId, 'tag', 'error', String(e));
  }

  // ⑤ metadataの空欄だけ書き戻す（§2-4）。🔴 既に値がある項目は絶対に送らない。
  try {
    const current = _lhGetFriend(friendId); // 最新値。キャッシュのstale判定に頼らない
    const meta    = current.metadata || {};
    const patch   = {};
    if (!meta.salon_name    && f.salonName)    patch.salon_name    = f.salonName;
    if (!meta.staff_name    && f.staffName)    patch.staff_name    = f.staffName;
    if (!meta.email         && f.email)        patch.email         = f.email;
    if (!meta.phone         && f.phone)        patch.phone         = f.phone;
    if (!meta.business_type && f.businessType) patch.business_type = f.businessType;

    if (Object.keys(patch).length > 0) {
      _lhPutMetadata(friendId, patch);
      _logLineSync(appId, friendId, 'metadata', 'ok', JSON.stringify(patch));
    } else {
      _logLineSync(appId, friendId, 'metadata', 'skipped', '埋める空欄なし');
    }
  } catch (e) {
    _logLineSync(appId, friendId, 'metadata', 'error', String(e));
  }
}

// ============================================================
// 🆕 名札印刷badges.html用の読み取りアクション（S3・名札印刷_badges設計.md §5-3・§5-4）
//    doPost: action=listBadges / listSpareBadges ―― 🔒 認証必須（_requireSession）
//
// 🔴 この節は読み取り専用。申込データを一切変更しない。既存の申込経路
//    （apply/applyLiff/updateApplication等）とは完全に独立している。
// ============================================================

// created_at/updated_at セルを 'yyyy-MM-dd' に正規化して文字列比較できるようにする。
// 🔴 通常は_now()が返す文字列だが、人がシートを編集するとDate型になりうるので両方吸収する
// （JSTで統一。_isTodayCellと同じ考え方）。
function _dateKey(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

// 業態文字列からA/B/C帯を決定する。6値のどれでもなければ（空欄・旧表記「エステ」等）
// 例外を書き足さずdefaultへ落とす（§5-3の設計方針と同じ）。
function _resolveBand(cfg, businessType) {
  const key = 'badge_color_' + String(businessType || '').trim();
  const v = cfg[key];
  return (v === 'A' || v === 'B' || v === 'C') ? v : String(cfg['badge_color_default'] || 'C');
}

// config の badge_band_*/badge_label_* から bands オブジェクトを組み立てる
function _buildBandsFromConfig(cfg) {
  const bands = {};
  ['A', 'B', 'C'].forEach(function (k) {
    bands[k] = {
      color: String(cfg['badge_band_' + k] || ''),
      label: String(cfg['badge_label_' + k] || '')
    };
  });
  return bands;
}

// 来場者名札の名簿を返す（doPost: action=listBadges）。
// data: { session_token, business_types?, date_field?, created_from?, created_to? }
function listBadges(data) {
  const auth = _requireSession(data);
  if (!auth.ok) return _err(auth.error);
  _checkProps();

  const businessTypes = Array.isArray(data.business_types) ? data.business_types : [];
  const dateField  = (data.date_field === 'updated_at') ? 'updated_at' : 'created_at';
  const fromKey    = String(data.created_from || '').trim();
  const toKey      = String(data.created_to || '').trim();
  const wantAll    = businessTypes.length === 0; // 省略 or 空配列 → 全業態
  const wantEmpty  = businessTypes.indexOf('__empty__') >= 0;
  const wantSet    = {};
  businessTypes.forEach(function (bt) { if (bt !== '__empty__') wantSet[bt] = true; });

  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh  = _getSheet(ss, SHEET_APPLICATIONS);
  const rows = sh.getDataRange().getValues();
  const cfg  = _getConfig(); // 🔴 _getConfig()は自分でopenByIdする実装のため呼び出しが2回になるが、
                              // 既存関数を書き換えない（§0-1の原則）。読み取り2回は誤差なので最適化しない

  let unknownCount   = 0;
  let skippedNoToken = 0;
  const badges = [];

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][17]) !== 'confirmed') continue; // status

    const token = String(rows[i][16] || '');
    if (!token) { skippedNoToken++; continue; } // ticket_tokenが空＝データ異常。名札は出せない

    const businessType = String(rows[i][20] || '');
    const known = BUSINESS_TYPE_OPTIONS.indexOf(businessType) >= 0;
    if (!known) unknownCount++;

    if (!wantAll) {
      const bizOk = known ? !!wantSet[businessType] : wantEmpty;
      if (!bizOk) continue;
    }

    if (fromKey || toKey) {
      const dk = _dateKey(rows[i][dateField === 'updated_at' ? 2 : 1]);
      if (fromKey && dk < fromKey) continue;
      if (toKey && dk > toKey) continue;
    }

    badges.push({
      app_id:        String(rows[i][0]),
      salon_name:    String(rows[i][4]),
      staff_name:    String(rows[i][5]),
      business_type: businessType,
      band:          _resolveBand(cfg, businessType),
      ticket_token:  token,
      created_at:    String(rows[i][1])
    });
  }

  badges.sort(function (a, b) { return a.app_id < b.app_id ? -1 : (a.app_id > b.app_id ? 1 : 0); });

  return _ok({
    pass_base: SITE_BASE_URL + 'pass.html',
    bands: _buildBandsFromConfig(cfg),
    unknown_business_type_count: unknownCount,
    skipped_no_token: skippedNoToken,
    total: badges.length,
    badges: badges
  });
}

// 予備名札の一覧を返す（doPost: action=listSpareBadges）。data: { session_token }
// 未割当のみ/すべての絞り込みは画面側で行う（§5-4。既定＝未割当のみ）。
function listSpareBadges(data) {
  const auth = _requireSession(data);
  if (!auth.ok) return _err(auth.error);
  _checkProps();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = _ensureSpareBadgesSheet(ss);
  const rows = sh.getDataRange().getValues();

  const spares = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    spares.push({
      spare_no:         String(rows[i][0]),
      ticket_token:      String(rows[i][1]),
      assigned_app_id:   String(rows[i][2] || ''),
      assigned_at:       String(rows[i][3] || '')
    });
  }

  return _ok({
    pass_base: SITE_BASE_URL + 'pass.html',
    spares: spares
  });
}

// ============================================================
// ② 入場パス取得（doGet: action=getPass）―― 完全公開・認証なし
//    token を知っている人だけが自分の情報を見られる
//
//    🆕 v0.8.0（L0-X）: pass.html の編集UI（L1-e）向けに編集対象項目を追加。
//    token を持つ本人しか開けないので返してよい（§8のcapability方針）。
// ============================================================
function getPass(data) {
  _checkProps();
  const token = String(data.ticket_token || '').trim();
  if (!token) return _err('INVALID_TOKEN');

  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh  = _getSheet(ss, SHEET_APPLICATIONS);
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][16]) === token) {
      const cfg = _getConfig();
      return _ok({
        app_id:         rows[i][0],
        salon_name:     rows[i][4],
        staff_name:     rows[i][5],
        status:         rows[i][17],
        checked_in_at:  rows[i][18] || null,
        // 🆕 編集UI用
        email:           rows[i][6],
        phone:           rows[i][8],
        business_type:   rows[i][20],
        has_transaction: rows[i][10],
        address:         rows[i][11],
        referrer:        rows[i][12],
        note:            rows[i][19],
        event: {
          date:       cfg.event_date       || '',
          time:       cfg.event_time       || '',
          venue_name: cfg.venue_name       || '',
          venue_addr: cfg.venue_addr       || ''
        },
        pass_url: SITE_BASE_URL + 'pass.html?t=' + token
      });
    }
  }
  return _err('NOT_FOUND');
}

// ============================================================
// 🆕 ③ 入場パスの再発行（doPost: action=resendPass）―― v0.7.0
//    パスURLを紛失した方が、メールアドレスだけで自分のパスを取り戻せるようにする。
//    これが無いと「URLを失った人がフォームを再送信 → 氏名の表記揺れで既存申込に
//    辿り着けず重複行ができる」という経路が残ってしまう（§4-3）。
//
//    🔴 メールアドレスの存在有無を応答で区別しない（該当が無くても success を返す）。
//    そうしないと「このメールアドレスは申込済みか」を外部から総当たりで調べられてしまう。
// ============================================================
function resendPass(data) {
  _checkProps();
  const email = String(data.email || '').trim();
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return _err('メールアドレスの形式が正しくありません');
  }
  const emailNorm = email.toLowerCase();

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh   = _getSheet(ss, SHEET_APPLICATIONS);
  const rows = sh.getDataRange().getValues();

  const list = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][7]).toLowerCase() === emailNorm && String(rows[i][17]) !== 'cancelled') {
      list.push({
        salonName: String(rows[i][4]),
        staffName: String(rows[i][5]),
        passUrl:   SITE_BASE_URL + 'pass.html?t=' + String(rows[i][16])
      });
    }
  }

  // 該当があるときだけ送る。応答は該当の有無にかかわらず同じ（列挙攻撃の防止）
  // 🆕 v0.13.0（配送障害対策P2・§7-2）: 配送失敗が続くとクライアントが最大6回再試行し、
  // 同じ宛先に同一内容のメールが何通も届いていた。直近10分以内に送信成功済みなら送らない。
  if (list.length > 0 && !_recentResendMailSent(email)) {
    _sendPassResendMail(email, list);
  }

  return _ok({ requested: true });
}

// mail_logの末尾を見て、直近10分以内に同じ宛先へtype=resend_pass の送信成功
// （status='ok'または'fallback'。'error'は送れていないので除外）が記録済みならtrueを返す。
// 🔴 CacheServiceは使わない（project_gas_cache_lessonの教訓）。末尾50行を見るだけで十分軽い。
function _recentResendMailSent(email) {
  const emailNorm = email.toLowerCase();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_MAIL_LOG);
  if (!sh) return false;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const scanRows  = Math.min(50, lastRow - 1);
  const startRow  = lastRow - scanRows + 1;
  const values    = sh.getRange(startRow, 1, scanRows, 4).getValues(); // sent_at, to, type, status
  const cutoffMs  = Date.now() - 10 * 60 * 1000;

  for (let i = 0; i < values.length; i++) {
    const sentAt = values[i][0];
    const to     = values[i][1];
    const type   = values[i][2];
    const status = values[i][3];
    if (type !== 'resend_pass') continue;
    if (String(to).toLowerCase() !== emailNorm) continue;
    if (status !== 'ok' && status !== 'fallback') continue;
    // sent_atは_now()の'yyyy-MM-dd HH:mm:ss'（Asia/Tokyo）文字列
    const sentDate = new Date(String(sentAt).replace(' ', 'T') + '+09:00');
    if (!isNaN(sentDate.getTime()) && sentDate.getTime() >= cutoffMs) return true;
  }
  return false;
}

// ============================================================
// 🆕 v0.14.0: 新規申込のLINE WORKS通知
//    LINEWORKS_WEBHOOK未設定なら何もしない（任意機能）。失敗しても申込自体は
//    既に成立済みのため、呼び出し側でtry/catchして握りつぶす（他の通知系と同じ方針）。
// ============================================================
function _notifyNewApplicationLineWorks(appId, f, source) {
  if (!LINEWORKS_WEBHOOK) return;
  const sourceLabel = source === 'liff' ? 'LINE' : 'Web';
  const text =
    '🎪 ビューフェス2026 新規申込\n' +
    'app_id: ' + appId + '（' + sourceLabel + '経由）\n' +
    'サロン名: ' + f.salonName + '\n' +
    'お名前: ' + f.staffName + '\n' +
    '業態: ' + f.businessType + '\n' +
    '電話: ' + f.phone;
  UrlFetchApp.fetch(LINEWORKS_WEBHOOK, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ body: { text: text } }),
    muteHttpExceptions: true
  });
}

// ============================================================
// メール送信
//    送信元は beaufes@gmail.com（send-as登録済み前提）。
//    GmailAppでの送信に失敗した場合（send-as未登録等）は
//    MailApp（送信元は実行アカウントのまま・差出人名とReply-Toで代替）にフォールバックする。
// ============================================================
function _sendConfirmationMail(email, salonName, staffName, passUrl, isUpdate, seminarLines, fullNotice) {
  const cfg       = _getConfig();
  const eventDate = cfg.event_date || '2026-10-26';
  const eventTime = cfg.event_time || '10:00〜16:00';
  const venueName = cfg.venue_name || '青島屋（AOSHIMAYA）';
  const venueAddr = cfg.venue_addr || '宮崎市青島2丁目12-11';

  const subject = isUpdate
    ? '【ビューフェス2026】お申し込み内容を更新しました'
    : '【ビューフェス2026】お申し込みありがとうございます（入場パス）';

  // 🆕 v0.20.0: 予約した枠（無ければ空文字＝これまでと同じ本文になる）
  const semList = (seminarLines && seminarLines.length) ? seminarLines : [];
  const semText = semList.length
    ? '▼ ご予約いただいたセミナー・体験会\n' + semList.map(function (s) { return '  ・' + s; }).join('\n') + '\n\n'
    : '';
  const semHtml = semList.length
    ? '<p>▼ ご予約いただいたセミナー・体験会<br>' +
      semList.map(function (s) { return '・' + _escapeHtml(s); }).join('<br>') + '</p>'
    : '';

  // 🆕 v0.23.0: 満席で予約が取れなかった場合。
  // 🔴 「申し込めていない」と誤解されないよう、入場申込は完了していることを必ず添える。
  const fullText = fullNotice
    ? '▼ ご希望のセミナー・体験会は満席のため、ご予約をお取りできませんでした\n' +
      '  ご入場のお申し込みは完了しております。\n' +
      '  空いている枠は、下記の入場パスの画面からご予約いただけます。\n\n'
    : '';
  const fullHtml = fullNotice
    ? '<p style="background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:10px 12px;">' +
      '▼ ご希望のセミナー・体験会は満席のため、ご予約をお取りできませんでした<br>' +
      '<span style="color:#92400e;">ご入場のお申し込みは完了しております。' +
      '空いている枠は、下記の入場パスの画面からご予約いただけます。</span></p>'
    : '';

  const textBody =
    salonName + '\n' +
    staffName + ' 様\n\n' +
    'ビューフェス2026へのお申し込みを' + (isUpdate ? '更新' : '受付') + 'いたしました。\n\n' +
    '  日時 : ' + eventDate + ' ' + eventTime + '\n' +
    '  会場 : ' + venueName + ' ' + venueAddr + '\n\n' +
    semText + fullText +
    '▼ 当日は入場口でこちらの入場パスをご提示ください\n' +
    '  ' + passUrl + '\n\n' +
    '※このメールを保存いただくか、上のリンクをスマホのホーム画面に追加しておくと当日スムーズです\n' +
    '\n' +
    '--\n' +
    'ビューフェス事務局（' + MAIL_FROM_ADDR + '）\n';

  const htmlBody =
    '<p>' + _escapeHtml(salonName) + '<br>' + _escapeHtml(staffName) + ' 様</p>' +
    '<p>ビューフェス2026へのお申し込みを' + (isUpdate ? '更新' : '受付') + 'いたしました。</p>' +
    '<p>日時: ' + _escapeHtml(eventDate) + ' ' + _escapeHtml(eventTime) + '<br>' +
    '会場: ' + _escapeHtml(venueName) + ' ' + _escapeHtml(venueAddr) + '</p>' +
    semHtml + fullHtml +
    '<p><a href="' + passUrl + '">▼ 入場パスを開く</a></p>' +
    '<p style="color:#666;font-size:13px;">' +
    '※このメールを保存いただくか、上のリンクをスマホのホーム画面に追加しておくと当日スムーズです</p>' +
    '<p style="color:#999;font-size:12px;">ビューフェス事務局（' + MAIL_FROM_ADDR + '）</p>';

  _sendMail(email, subject, textBody, htmlBody, isUpdate ? 'resend' : 'apply');
}

// 🆕 v0.24.0: セミナー・体験会の予約を変更/取消したときの控えメール。
// 🔴 「取り消しました」だけを読んだお客様が、来場申込ごと消えたと誤解しないよう、
//    ご入場のお申し込みは有効であることを必ず本文に入れる。
function _sendReservationMail(email, salonName, staffName, passUrl, labels, fullNotice) {
  const list    = labels || [];
  const subject = list.length
    ? '【ビューフェス2026】セミナー・体験会のご予約内容'
    : '【ビューフェス2026】セミナー・体験会のご予約を取り消しました';

  // 🔴 件名と書き出しを揃える。全部取り消したのに「承りました」だと読み手が混乱する。
  const lead = list.length
    ? 'セミナー・体験会のご予約内容を承りました。'
    : 'セミナー・体験会のご予約を取り消しました。';

  const bodyLines = list.length
    ? '▼ 現在のご予約\n' + list.map(function (x) { return '  ・' + x; }).join('\n') + '\n\n'
    : '▼ 現在のご予約\n  ご予約はありません（すべて取り消しました）\n\n';

  const fullText = fullNotice
    ? '▼ ご希望のセミナー・体験会は満席のため、ご予約をお取りできませんでした\n' +
      '  空いている枠は、下記の画面からご予約いただけます。\n\n'
    : '';

  const textBody =
    salonName + '\n' +
    staffName + ' 様\n\n' +
    lead + '\n\n' +
    bodyLines + fullText +
    '▼ ご予約の変更・取り消し、入場パスの表示はこちら\n' +
    '  ' + passUrl + '\n\n' +
    '※ビューフェス2026へのご入場のお申し込みは有効です。\n' +
    '　当日は上のリンクの入場パスをご提示ください。\n' +
    '\n--\n' +
    'ビューフェス事務局（' + MAIL_FROM_ADDR + '）\n';

  const htmlList = list.length
    ? '<p>▼ 現在のご予約<br>' + list.map(function (x) { return '・' + _escapeHtml(x); }).join('<br>') + '</p>'
    : '<p>▼ 現在のご予約<br>ご予約はありません（すべて取り消しました）</p>';

  const htmlFull = fullNotice
    ? '<p style="background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:10px 12px;">' +
      '▼ ご希望のセミナー・体験会は満席のため、ご予約をお取りできませんでした<br>' +
      '<span style="color:#92400e;">空いている枠は、下記の画面からご予約いただけます。</span></p>'
    : '';

  const htmlBody =
    '<p>' + _escapeHtml(salonName) + '<br>' + _escapeHtml(staffName) + ' 様</p>' +
    '<p>' + _escapeHtml(lead) + '</p>' +
    htmlList + htmlFull +
    '<p><a href="' + passUrl + '">▼ ご予約の変更・取り消し、入場パスの表示</a></p>' +
    '<p style="color:#666;font-size:13px;">※ビューフェス2026へのご入場のお申し込みは有効です。<br>' +
    '当日は上のリンクの入場パスをご提示ください。</p>' +
    '<p style="color:#999;font-size:12px;">ビューフェス事務局（' + MAIL_FROM_ADDR + '）</p>';

  _sendMail(email, subject, textBody, htmlBody, 'reservation');
}

// 🆕 入場パスの再発行メール（v0.7.0）
// 同じメールアドレスで複数名の申込があることは普通に起きるため、
// 1通にまとめて「◯◯様の入場パス」と担当者名を明記して並べる（§4-3）。
function _sendPassResendMail(email, list) {
  const subject = '【ビューフェス2026】入場パスのリンクをお送りします';

  let textBody = 'ビューフェス2026の入場パスのリンクをお送りします。\n\n';
  let htmlBody = '<p>ビューフェス2026の入場パスのリンクをお送りします。</p>';

  for (let i = 0; i < list.length; i++) {
    const it = list[i];
    textBody += '■ ' + it.staffName + ' 様（' + it.salonName + '）\n' +
                '  ' + it.passUrl + '\n\n';
    htmlBody += '<p style="margin:14px 0;">■ ' + _escapeHtml(it.staffName) + ' 様' +
                '（' + _escapeHtml(it.salonName) + '）<br>' +
                '<a href="' + it.passUrl + '">入場パスを開く</a></p>';
  }

  textBody += '当日は入場口でこちらの画面をご提示ください。\n' +
              '※このメールを保存いただくか、リンクをスマホのホーム画面に追加しておくと当日スムーズです\n\n' +
              '--\nビューフェス事務局（' + MAIL_FROM_ADDR + '）\n';
  htmlBody += '<p style="color:#666;font-size:13px;">当日は入場口でこちらの画面をご提示ください。<br>' +
              '※このメールを保存いただくか、リンクをスマホのホーム画面に追加しておくと当日スムーズです</p>' +
              '<p style="color:#999;font-size:12px;">ビューフェス事務局（' + MAIL_FROM_ADDR + '）</p>';

  _sendMail(email, subject, textBody, htmlBody, 'resend_pass');
}

function _sendMail(to, subject, textBody, htmlBody, type) {
  let status = 'ok', error = '';
  try {
    GmailApp.sendEmail(to, subject, textBody, {
      name:     MAIL_FROM_NAME,
      from:     MAIL_FROM_ADDR,
      replyTo:  MAIL_FROM_ADDR,
      htmlBody: htmlBody
    });
  } catch (e1) {
    // send-as未登録・Gmailサービス未使用等でGmailAppが使えない場合のフォールバック
    try {
      MailApp.sendEmail({
        to:       to,
        subject:  subject,
        name:     MAIL_FROM_NAME,
        replyTo:  MAIL_FROM_ADDR,
        body:     textBody,
        htmlBody: htmlBody
      });
      status = 'fallback';
      error  = 'GmailApp失敗のためMailAppにフォールバック: ' + e1;
    } catch (e2) {
      status = 'error';
      error  = String(e2);
    }
  }
  _logMail(to, type, status, error);
}

function _logMail(to, type, status, error) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = _getSheet(ss, SHEET_MAIL_LOG);
    sh.appendRow([_now(), to, type, status, error]);
  } catch (e) {
    Logger.log('mail_log書き込み失敗: ' + e);
  }
}

// ============================================================
// 🆕 LINE Harness連携（L1-a・UIなし。doGet/doPostからはまだ呼ばれない）
//
// 実装前提（LINE連携_実装プラン.md §1・§2で実測済み。設計書やbcart-approvalの
// TypeScript型と食い違う場合はそちらを疑う）:
//   - friendオブジェクトのキーは lineUserId（キャメルケース）。line_user_idは誤り（§1-2）
//   - ?line_user_id= と ?page= は完全に無視される。ページングは ?limit=&offset= のみ有効（§1-3）
//   - デフォルトlimitは50・友だちは102名いる → 素直に呼ぶと半分しか返らない（§1-3 地雷2）
//   - lineUserIdで直接検索する手段がAPIに無い → 全件キャッシュ方式にする（§2-1）
//   - CacheServiceは使わない（project_gas_cache_lessonの教訓）。シートに持つ
// ============================================================

// Bearer認証つきUrlFetchApp。非2xx・非JSONはどちらも例外にする（無言で握りつぶさない）
function _lhReq(method, path, payload) {
  _checkLineHarnessProps();

  const options = {
    method: method,
    headers: { Authorization: 'Bearer ' + LINE_HARNESS_API_KEY },
    muteHttpExceptions: true
  };
  if (payload) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }

  const res  = UrlFetchApp.fetch(LINE_HARNESS_API_URL + path, options);
  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('LINE Harness API エラー: HTTP_' + code + ' ' + method + ' ' + path + ' / ' + text.substring(0, 300));
  }
  try {
    return JSON.parse(text);
  } catch (e) {
    throw new Error('LINE Harness API 非JSON応答: ' + method + ' ' + path + ' / ' + text.substring(0, 300));
  }
}

// 全友だちを limit=100 + offset ループで取得する。
// 🔴 page は完全に無視されるため使わない（§1-3）。offset>=total または空配列で終了。
// 安全弁として最大50ループ（1ページ100件なので5000件まで対応・102件なら1〜2ループで終わる）
function _lhFetchAllFriends() {
  const LIMIT     = 100;
  const MAX_LOOPS = 50;
  let offset = 0;
  let total  = null;
  const items = [];

  for (let loop = 0; loop < MAX_LOOPS; loop++) {
    const json = _lhReq('GET', '/api/friends?limit=' + LIMIT + '&offset=' + offset, null);
    if (!json || json.success !== true || !json.data) {
      throw new Error('LINE Harness友だち一覧の取得に失敗しました: ' + JSON.stringify(json).substring(0, 300));
    }
    const pageItems = json.data.items || [];
    if (total === null) total = json.data.total;

    Array.prototype.push.apply(items, pageItems);

    if (pageItems.length === 0) break;
    offset += pageItems.length;
    if (typeof total === 'number' && offset >= total) break;
  }

  return items;
}

// line_friends_cache シートが無ければヘッダー付きで作成する
function _ensureLineFriendsCacheSheet(ss) {
  let sh = ss.getSheetByName(SHEET_LINE_FRIENDS_CACHE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LINE_FRIENDS_CACHE);
    sh.getRange(1, 1, 1, 9).setValues([[
      'line_user_id', 'friend_id', 'salon_name', 'staff_name', 'phone',
      'business_type', 'email', 'ext_id', 'synced_at'
    ]]);
    sh.setFrozenRows(1);
    Logger.log('line_friends_cacheシート作成完了');
  }
  return sh;
}

// LINE Harnessから全友だちを取得し、line_friends_cacheシートを全書き換えする。
// 時間トリガーで6時間ごとに実行する想定（GASエディタのトリガー画面から手動設定。§4 L1-a）。
function syncLineFriends() {
  _checkProps();
  const friends = _lhFetchAllFriends();

  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh  = _ensureLineFriendsCacheSheet(ss);
  const now = _now();

  // 既存データを消してから書き直す（ヘッダーは残す）
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, 9).clearContent();
  }

  if (friends.length > 0) {
    const rows = friends.map(function (f) {
      const meta = f.metadata || {};
      return [
        f.lineUserId || f.line_user_id || '', // 🔴 両対応（§1-2の地雷。実際はlineUserIdのみ存在）
        f.id || '',
        meta.salon_name     || '',
        meta.staff_name     || '',
        meta.phone          || '',
        meta.business_type  || '',
        meta.email          || '',
        meta.ext_id         || '',
        now
      ];
    });
    // 🔴 電話番号(E列)・ext_id(H列)はPlain Text指定してから書き込む。
    // 全て数字の文字列を書き込むと列の書式がAutomaticのままだと数値型に変換され、
    // 先頭の0が消える（applications.phoneで実際に発生した不具合と同じ原因。§migrateFixPhoneColumn参照）
    sh.getRange(2, 5, rows.length, 1).setNumberFormat('@');
    sh.getRange(2, 8, rows.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, 9).setValues(rows);
  }

  Logger.log('syncLineFriends完了: ' + friends.length + '件を同期しました（friends.total想定: 102件前後）。');
  return friends.length;
}

// キャッシュシートから line_user_id で1件検索する（内部ヘルパー）
function _searchLineFriendsCache(sh, lineUserId) {
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === lineUserId) {
      return {
        lineUserId:   String(rows[i][0]),
        friendId:     String(rows[i][1]),
        salonName:    String(rows[i][2]),
        staffName:    String(rows[i][3]),
        phone:        String(rows[i][4]),
        businessType: String(rows[i][5]),
        email:        String(rows[i][6]),
        extId:        String(rows[i][7])
      };
    }
  }
  return null;
}

// line_user_id（IDトークンのsub）から友だちを引く。キャッシュミス時は1回だけ
// syncLineFriends()を実行して再検索する（新規友だち対応・§2-1）。
// それでも見つからなければnullを返す（呼び出し側は空のフォームを出す。エラーにしない）
function _findFriendByLineUserId(lineUserId) {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = _ensureLineFriendsCacheSheet(ss);

  const found = _searchLineFriendsCache(sh, lineUserId);
  if (found) return found;

  syncLineFriends();
  return _searchLineFriendsCache(_ensureLineFriendsCacheSheet(ss), lineUserId);
}

// 診断用（GASエディタから手動実行）。1ページ目だけ取得し、件数・キー名・lineUserIdの
// 有無をログ出力する。L0-1（GAS→LINE Harness疎通）の完了条件確認に使う。
function probeLineHarness() {
  const json  = _lhReq('GET', '/api/friends?limit=1&offset=0', null);
  const data  = (json && json.data) ? json.data : {};
  const items = data.items || [];
  const first = items[0] || {};

  Logger.log('total: ' + data.total);
  Logger.log('1件目のキー: ' + Object.keys(first).join(', '));
  Logger.log('lineUserIdあり: ' + (first.lineUserId !== undefined));
  Logger.log('line_user_idあり（誤表記チェック・falseが正常）: ' + (first.line_user_id !== undefined));
}

// ============================================================
// 🆕 L1-b: LINEのIDトークン検証
//
// 🔴 liff.getProfile()の結果をブラウザから送られてきたまま信用してはいけない
// （なりすまし申込が可能になる・§2-3）。IDトークンを必ずサーバー側で検証し、
// 応答（aud/exp/iss）を鵜呑みにせず自分で判定する。
// ============================================================

// LINEのIDトークンを検証し、本人を一意に表すsub（= lineUserId）を返す。
// 不正・改竄・期限切れ・他チャネル宛のトークンは必ず例外にする（無言でnullを返さない）。
function _verifyLineIdToken(idToken) {
  if (!idToken) throw new Error('IDトークンがありません');

  const options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      id_token:  idToken,
      client_id: LIFF_CHANNEL_ID
    },
    muteHttpExceptions: true // 🔴 自動では例外にならないため、必ず自分でステータス判定する（§2-3）
  };

  const res  = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', options);
  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code !== 200) {
    throw new Error('IDトークン検証に失敗しました: HTTP_' + code + ' / ' + text.substring(0, 300));
  }

  let json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new Error('IDトークン検証: 非JSON応答 / ' + text.substring(0, 300));
  }

  // 🔴 LINE側が200を返しても、内容は自分で検証する（§2-3の3項目）
  if (json.iss !== 'https://access.line.me') {
    throw new Error('IDトークンのissが不正です: ' + json.iss);
  }
  if (json.aud !== LIFF_CHANNEL_ID) {
    throw new Error('IDトークンのaudが自チャネル宛ではありません: ' + json.aud);
  }
  const nowSec = Math.floor(Date.now() / 1000);
  if (!json.exp || json.exp <= nowSec) {
    throw new Error('IDトークンの有効期限が切れています');
  }
  if (!json.sub) {
    throw new Error('IDトークンにsub（ユーザー識別子）がありません');
  }

  return json.sub;
}

// 診断用（GASエディタから手動実行）。実在しない/改竄されたIDトークンを渡し、
// 確実に例外になることを確認する。正しいIDトークンでの確認はliff.html実装後（L1-d）に実機で行う。
function testVerifyBadIdToken() {
  try {
    _verifyLineIdToken('this-is-not-a-real-id-token');
    Logger.log('🔴 異常: 不正なトークンなのに例外になりませんでした（要調査）');
  } catch (e) {
    Logger.log('✅ 正常: 不正なトークンで例外になりました → ' + e.message);
  }
}

// ============================================================
// 🆕 セミナー予約（sessions / reservations）—— v0.20.0（2026-09-04）
//    設計: `LINEHarness/ビューフェス申込_設計.md` §4-2
//
//    【この設計の要点】
//    - 枠は `sessions` シートが正。**枠を増やすのはシートに1行足すだけ**で、
//      コード修正もGAS再デプロイもいらない（画面はこのAPIの返り値だけを見て描く）。
//    - `slot`（時間帯グループ）が同じ枠は**排他**。画面ではslotごとのラジオになり、
//      サーバー側でも同一slotの二重予約を拒否する。同じ時間に2枠取る事故を構造的に潰す。
//    - `capacity` 空欄 = 定員なし。設定されていれば満席で自動的に予約不可。
//    - 🔴🔴 **セミナーが満席でも来場申込は絶対に止めない**（設計書§4-2の絶対条件）。
//      申込処理の中で満席に当たった場合も申込行は書き、`sessions_result.full` で
//      「予約だけ取れなかった」ことを返す。ここを例外にしたりエラー返しにしないこと。
//    - キャンセル待ちは作らない（要件）。取消は pass.html から（席がその場で戻る）。
// ============================================================

const RES_STATUS_RESERVED  = 'reserved';
const RES_STATUS_CANCELLED = 'cancelled';

// sessions シートの列（0始まり）。J・K は v0.20.0 で追加した詳細表示用。
const SES_COL = {
  id: 0, slot: 1, title: 2, speaker: 3, room: 4,
  starts: 5, ends: 6, capacity: 7, active: 8,
  bullets: 9, overview: 10
};
// reservations シートの列（0始まり）
const RES_COL = { id: 0, appId: 1, sessionId: 2, createdAt: 3, status: 4, attendedAt: 5 };

// 🔴 時刻セルは Date になっている場合がある（"10:30" と打つとスプレッドシートが
// 勝手に時刻値へ変換するため）。文字列のまま String() すると
// "Sat Dec 30 1899 10:30:00 GMT+0919" のような値が画面に出る。必ずこの関数を通す。
function _fmtTimeCell(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'HH:mm');
  }
  return String(v).trim();
}

// "HH:mm" を0時からの分に直す。読めなければ null。
function _hhmmToMinutes(v) {
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(v == null ? '' : v).trim());
  if (!m) return null;
  return Number(m[1]) * 60 + Number(m[2]);
}

// 2つの枠の時間が重なるか。
// 🔴 隣り合う枠（10:45終わり ／ 10:45始まり）は**重ならない**扱いにする（`<` で比較する理由）。
//    バランス革命の連続枠がすべて競合してしまうため。
// 🔵 開始時刻が無い枠は判定できないので「重ならない」とする（判定できないことを理由に
//    予約を止めるより、取れてしまうほうが害が小さい。当日の受付で調整できる）。
// 🔵 終了時刻が無い枠は、開始が同じときだけ重複とみなす。
function _sessionsOverlap(a, b) {
  const as = _hhmmToMinutes(a.starts_at), bs = _hhmmToMinutes(b.starts_at);
  if (as === null || bs === null) return false;
  const ae = _hhmmToMinutes(a.ends_at),  be = _hhmmToMinutes(b.ends_at);
  if (ae === null || be === null) return as === bs;
  return as < be && bs < ae;
}

// 改行区切りのセルを配列にする（セミナー内容の箇条書き用）
function _splitLines(v) {
  return String(v == null ? '' : v)
    .split(/\r?\n/)
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s !== ''; });
}

// sessions シートを読む（無ければ空配列。シート未作成でも画面を壊さない）
function _readSessions(ss) {
  const sh = ss.getSheetByName(SHEET_SESSIONS);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][SES_COL.id] == null ? '' : rows[i][SES_COL.id]).trim();
    if (!id) continue;
    const capRaw = rows[i][SES_COL.capacity];
    const capNum = (capRaw === '' || capRaw === null || capRaw === undefined) ? null : Number(capRaw);
    out.push({
      session_id: id,
      slot:       String(rows[i][SES_COL.slot]  == null ? '' : rows[i][SES_COL.slot]).trim() || id,
      title:      String(rows[i][SES_COL.title] == null ? '' : rows[i][SES_COL.title]).trim(),
      speaker:    String(rows[i][SES_COL.speaker] == null ? '' : rows[i][SES_COL.speaker]).trim(),
      room:       String(rows[i][SES_COL.room]  == null ? '' : rows[i][SES_COL.room]).trim(),
      starts_at:  _fmtTimeCell(rows[i][SES_COL.starts]),
      ends_at:    _fmtTimeCell(rows[i][SES_COL.ends]),
      capacity:   (capNum === null || isNaN(capNum)) ? null : capNum,
      // 🔴 空欄は「有効」。書き忘れで枠が黙って消えるほうが害が大きい（booth と同じ判断）
      is_active:  _boothIsActive(rows[i][SES_COL.active]),
      bullets:    _splitLines(rows[i][SES_COL.bullets]),
      overview:   String(rows[i][SES_COL.overview] == null ? '' : rows[i][SES_COL.overview]).trim()
    });
  }
  return out;
}

// reservations シートの生データ（ヘッダー込み）。書き込み側は行番号が要るのでそのまま返す。
function _readReservationRows(ss) {
  const sh = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!sh) return [];
  return sh.getDataRange().getValues();
}

// 🆕 v0.31.0: キャンセル済み申込の app_id 集合。
// 🔴 申込のキャンセルはスプレッドシートの status 列を手で書き換えて行われる（専用の画面が無い）。
//    つまり「キャンセルされた瞬間」をコードで捕まえることはできない。
//    そこで**数えるときに毎回除外する**ことで、手作業のキャンセルでも席が確実に戻るようにしている。
//    予約行そのものを消しに行く方式にしないこと（手編集を取りこぼす）。
function _cancelledAppIdSet(appRows) {
  const set = {};
  for (let i = 1; i < appRows.length; i++) {
    if (String(appRows[i][17]) === 'cancelled') set[String(appRows[i][0])] = true;
  }
  return set;
}

// applications シートを読んでキャンセル済み集合を作る
function _readCancelledAppIds(ss) {
  const sh = ss.getSheetByName(SHEET_APPLICATIONS);
  if (!sh) return {};
  return _cancelledAppIdSet(sh.getDataRange().getValues());
}

// 予約済み（reserved）の件数を session_id ごとに数える。
// 🔴 cancelled（キャンセル済み申込の集合）を渡すと、その人の予約は数えない＝席が戻る。
function _countReserved(resRows, cancelled) {
  const counts = {};
  for (let i = 1; i < resRows.length; i++) {
    if (String(resRows[i][RES_COL.status]) !== RES_STATUS_RESERVED) continue;
    if (cancelled && cancelled[String(resRows[i][RES_COL.appId])]) continue;
    const sid = String(resRows[i][RES_COL.sessionId] == null ? '' : resRows[i][RES_COL.sessionId]).trim();
    if (!sid) continue;
    counts[sid] = (counts[sid] || 0) + 1;
  }
  return counts;
}

// その人（app_id）が現在予約している session_id の配列
function _reservedSessionIdsOf(resRows, appId) {
  const out = [];
  if (!appId) return out;
  for (let i = 1; i < resRows.length; i++) {
    if (String(resRows[i][RES_COL.status]) !== RES_STATUS_RESERVED) continue;
    if (String(resRows[i][RES_COL.appId]) !== String(appId)) continue;
    out.push(String(resRows[i][RES_COL.sessionId]).trim());
  }
  return out;
}

// res_id の採番。既存の R0001 形式の最大値の次から発番するクロージャを返す。
function _resIdIssuer(resRows) {
  let max = 0;
  for (let i = 1; i < resRows.length; i++) {
    const m = /^R(\d+)$/.exec(String(resRows[i][RES_COL.id] == null ? '' : resRows[i][RES_COL.id]).trim());
    if (m) max = Math.max(max, Number(m[1]));
  }
  return function () { max++; return 'R' + _boothPad(max, 4); };
}

// 公開されている枠（is_active）だけを、残席つきで返す共通部品
function _sessionCatalog(ss, resRows, cancelled) {
  const counts = _countReserved(resRows, cancelled);
  return _readSessions(ss)
    .filter(function (s) { return s.is_active; })
    .map(function (s) {
      const used = counts[s.session_id] || 0;
      const remaining = (s.capacity === null) ? null : Math.max(0, s.capacity - used);
      return {
        session_id: s.session_id,
        slot:       s.slot,
        title:      s.title,
        speaker:    s.speaker,
        room:       s.room,
        starts_at:  s.starts_at,
        ends_at:    s.ends_at,
        capacity:   s.capacity,
        remaining:  remaining,
        is_full:    (remaining !== null && remaining <= 0),
        bullets:    s.bullets,
        overview:   s.overview
      };
    });
}

// 読み込み済みの applications 行から ticket_token で申込者を引く
function _findApplicantInRows(rows, token) {
  if (!token) return null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][16]) !== token) continue;
    return {
      appId:     String(rows[i][0]),
      salonName: String(rows[i][4]),
      staffName: String(rows[i][5]),
      email:     String(rows[i][6]),
      status:    String(rows[i][17])
    };
  }
  return null;
}

// ticket_token から申込者を引く（見つからなければ null）。
// 予約の控えメールに宛先と氏名が要るので、app_id だけでなく行の内容も返す。
function _applicantByTicketToken(ss, token) {
  if (!token) return null;
  const sh = _getSheet(ss, SHEET_APPLICATIONS);
  return _findApplicantInRows(sh.getDataRange().getValues(), token);
}

// ticket_token から app_id を引く（見つからなければ null）
function _appIdByTicketToken(ss, token) {
  const a = _applicantByTicketToken(ss, token);
  return a ? a.appId : null;
}

// 予約1件の表示名。"11:00〜12:00  フェムケアセミナー" の形。
// 🔴 バランス革命の計測会のように**タイトルが時刻そのもの**（"10:00〜10:45の回"）の枠では、
//    時刻＋タイトルにすると「10:00〜10:45  10:00〜10:45の回」と二重になる。
//    その場合はタイトルではなく slot 名（"バランス革命 無料計測会"）を使う。
function _sessionDisplayName(s) {
  const time = (s.starts_at && s.ends_at) ? (s.starts_at + '〜' + s.ends_at)
             : (s.starts_at ? (s.starts_at + '〜') : '');
  const titleIsTime = !!(s.starts_at && String(s.title).indexOf(s.starts_at) >= 0);
  const name = (titleIsTime && s.slot) ? s.slot : s.title;
  return time ? (time + '  ' + name) : name;
}

// その人が予約している枠の表示名の配列（確認メール用）。
// 🔴 ロックの外から呼ぶこと（シートを読むだけ）。失敗しても空配列を返し、メール送信を止めない。
function _reservedSessionLabels(appId) {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const resRows = _readReservationRows(ss);
    const ids     = _reservedSessionIdsOf(resRows, appId);
    if (!ids.length) return [];
    const byId = {};
    _readSessions(ss).forEach(function (s) { byId[s.session_id] = s; });
    return ids.map(function (id) {
      const s = byId[id];
      return s ? _sessionDisplayName(s) : id;
    });
  } catch (e) {
    Logger.log('_reservedSessionLabels失敗（メールは予約欄なしで送る）: ' + e);
    return [];
  }
}

// ============================================================
// 🆕 セミナー枠の一覧（doGet/doPost: action=listSessions）— 認証なし・公開
//    ticket_token を添えると「その人の現在の予約」も一緒に返す（pass.html用）。
//    🔴 氏名・メール等の個人情報は返さない。返すのは枠の情報と自分の予約だけ。
// ============================================================
function listSessions(data) {
  _checkProps();
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const resRows = _readReservationRows(ss);
  // 🔴 キャンセル済み申込の予約は席を占有しない（v0.31.0）。ここを外すと
  //    キャンセルした人の分だけ定員が減ったままになる。
  const appSh    = _getSheet(ss, SHEET_APPLICATIONS);
  const appRows  = appSh.getDataRange().getValues();
  const cancelled = _cancelledAppIdSet(appRows);
  const catalog  = _sessionCatalog(ss, resRows, cancelled);

  const token = String((data && data.ticket_token) || '').trim();
  let reserved = [];
  if (token) {
    let appId = null;
    for (let i = 1; i < appRows.length; i++) {
      if (String(appRows[i][16]) === token) { appId = String(appRows[i][0]); break; }
    }
    if (appId) reserved = _reservedSessionIdsOf(resRows, appId);
  }
  return _ok({ sessions: catalog, reserved: reserved });
}

// ============================================================
// 🆕 v0.31.0 社員用: 予約状況の一覧（doPost: action=listReservations）🔒
//    admin.html から呼ぶ。枠ごとの予約数と、誰が予約しているかを返す。
//    🔴 キャンセル済み申込の予約は `app_status='cancelled'` を付けて返し、
//       件数（reserved_count）には数えない。画面側で「席は戻っている」と分かるようにするため。
// ============================================================
function listReservations(data) {
  const auth = _requireSession(data);
  if (!auth.ok) return _err(auth.error);

  _checkProps();
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const appSh   = _getSheet(ss, SHEET_APPLICATIONS);
  const appRows = appSh.getDataRange().getValues();
  const resRows = _readReservationRows(ss);
  const cancelled = _cancelledAppIdSet(appRows);

  // app_id -> 申込の表示用情報
  const byApp = {};
  for (let i = 1; i < appRows.length; i++) {
    const id = String(appRows[i][0]);
    if (!id) continue;
    byApp[id] = {
      salon_name: String(appRows[i][4]),
      staff_name: String(appRows[i][5]),
      phone:      String(appRows[i][8]),
      tantou:     String(appRows[i][22] == null ? '' : appRows[i][22]),
      app_status: String(appRows[i][17])
    };
  }

  const rows = [];
  for (let i = 1; i < resRows.length; i++) {
    if (String(resRows[i][RES_COL.status]) !== RES_STATUS_RESERVED) continue;
    const appId = String(resRows[i][RES_COL.appId]);
    const info  = byApp[appId] || { salon_name: '(申込が見つかりません)', staff_name: '', phone: '', tantou: '', app_status: '' };
    rows.push({
      res_id:     String(resRows[i][RES_COL.id]),
      app_id:     appId,
      session_id: String(resRows[i][RES_COL.sessionId]).trim(),
      created_at: String(resRows[i][RES_COL.createdAt]),
      salon_name: info.salon_name,
      staff_name: info.staff_name,
      phone:      info.phone,
      tantou:     info.tantou,
      app_status: info.app_status
    });
  }

  // 枠の一覧（非公開の枠も社員には見せる。当日の名簿づくりで必要になるため）
  const counts = _countReserved(resRows, cancelled);
  const sessions = _readSessions(ss).map(function (s) {
    const used = counts[s.session_id] || 0;
    return {
      session_id: s.session_id,
      slot:       s.slot,
      title:      s.title,
      starts_at:  s.starts_at,
      ends_at:    s.ends_at,
      capacity:   s.capacity,
      is_active:  s.is_active,
      reserved_count: used,
      remaining:  (s.capacity === null) ? null : Math.max(0, s.capacity - used)
    };
  });

  return _ok({ sessions: sessions, reservations: rows, viewer: auth.session.name || '' });
}

// ============================================================
// 🆕 v0.31.0 手動実行: キャンセル済み申込の予約行を cancelled にする（後片付け）
//    席の計算は _countReserved が毎回除外するので**実行しなくても定員は正しい**。
//    reservations シートを人が見たときに紛らわしくないよう整えるためだけのもの。
// ============================================================
function releaseCancelledReservations() {
  _checkProps();
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh  = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!sh) { Logger.log('reservationsシートがありません。'); return; }

  const cancelled = _readCancelledAppIds(ss);
  const rows = sh.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][RES_COL.status]) !== RES_STATUS_RESERVED) continue;
    if (!cancelled[String(rows[i][RES_COL.appId])]) continue;
    sh.getRange(i + 1, RES_COL.status + 1).setValue(RES_STATUS_CANCELLED);
    n++;
  }
  Logger.log('キャンセル済み申込の予約を' + n + '件 cancelled にしました。');
}

// ============================================================
// 予約の書き込み（内部関数）
// 🔴🔴 必ず LockService のロックの中から呼ぶこと。残席の判定と書き込みの間に
//      他のリクエストが割り込むと定員を超える。
//
//    wantIds : 予約したい session_id の配列（[] なら「全部取消」の意味になる・replace時）
//    opts.replace : true なら wantIds に無い自分の既存予約を cancelled にする
//                   （pass.html からの変更・取消）。false なら追加のみ（申込時）。
//    opts.cancelledAppIds : キャンセル済み申込の集合（席の計算から除外する）。
//                   🔵 呼び出し元が applications を既に読んでいるなら渡すこと。
//                   渡さないとここで読み直す＝ロックの中でシート読み込みが1回増える。
//
//    返り値 { reserved: [], added: [], full: [], invalid: [], conflict: [], cancelled: [] }
//    conflict = 既に予約している枠と**時間が重なる**ため取れなかったもの（slotをまたいで判定）
//    🔴 reserved は「保存後に予約されている枠」、added は「今回あらたに入った枠」。
//       控えメールを送るかどうかは added / cancelled で判断する（同じ内容の再保存では送らない）。
//    🔴 例外は投げない。呼び出し元（申込処理）を巻き込んで申込そのものを失敗させないため。
// ============================================================
function _writeReservations(ss, appId, wantIds, opts) {
  const replace = !!(opts && opts.replace);
  const result  = { reserved: [], added: [], full: [], invalid: [], conflict: [], cancelled: [] };
  if (!appId) return result;

  const sh = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!sh) {                     // セミナーを使わない運用（シート未作成）なら何もしない
    result.invalid = (wantIds || []).slice();
    return result;
  }

  const all   = _readSessions(ss);
  const byId  = {};
  all.forEach(function (s) { byId[s.session_id] = s; });

  const cancelled = (opts && opts.cancelledAppIds) || _readCancelledAppIds(ss);
  const rows     = sh.getDataRange().getValues();
  const counts   = _countReserved(rows, cancelled);
  const nextId   = _resIdIssuer(rows);
  const mineRow  = {};           // session_id -> 自分の予約行（1始まり）
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][RES_COL.status]) !== RES_STATUS_RESERVED) continue;
    if (String(rows[i][RES_COL.appId]) !== String(appId)) continue;
    mineRow[String(rows[i][RES_COL.sessionId]).trim()] = i + 1;
  }

  const want = [];
  (wantIds || []).forEach(function (raw) {
    const sid = String(raw == null ? '' : raw).trim();
    if (sid && want.indexOf(sid) < 0) want.push(sid);
  });

  // --- 1) 取消（replace時のみ）。先に取り消してから追加する。
  //     同じslot内で「Aをやめて Bにする」場合に、先に席を戻さないと自分の予約が
  //     自分の席を塞いで満席扱いになるため、この順序は変えないこと。
  if (replace) {
    Object.keys(mineRow).forEach(function (sid) {
      if (want.indexOf(sid) >= 0) return;
      sh.getRange(mineRow[sid], RES_COL.status + 1).setValue(RES_STATUS_CANCELLED);
      counts[sid] = Math.max(0, (counts[sid] || 1) - 1);
      delete mineRow[sid];
      result.cancelled.push(sid);
    });
  }

  // --- 2) 追加
  const slotTaken = {};          // slot -> session_id（同一slotの二重予約を防ぐ）
  const kept      = [];          // いま保持している枠（時間の重複判定に使う）
  Object.keys(mineRow).forEach(function (sid) {
    const s = byId[sid];
    if (!s) return;
    slotTaken[s.slot] = sid;
    kept.push(s);
  });

  const appends = [];
  want.forEach(function (sid) {
    const s = byId[sid];
    if (!s || !s.is_active) { result.invalid.push(sid); return; }       // 存在しない/終了した枠
    if (mineRow[sid]) { result.reserved.push(sid); return; }            // すでに予約済み＝冪等（再送で増えない）
    if (slotTaken[s.slot]) { result.invalid.push(sid); return; }        // 同じ時間帯を二重に取ろうとした
    // 🔴 slot をまたいで時間が重なる枠は取れない（2026-09-05・Takashiさん指定）。
    // 満席の判定より前に見る（「重なっている」ほうが利用者に伝えるべき理由として的確なため）。
    // 🔴 既に予約済みの枠（mineRow にあるもの）はこのチェックにかけない。
    //    過去に入った重なりを、無関係な保存のたびに黙って消してしまうのを防ぐ。
    if (kept.some(function (k) { return _sessionsOverlap(k, s); })) {
      result.conflict.push(sid);
      return;
    }

    const used = counts[sid] || 0;
    if (s.capacity !== null && used >= s.capacity) { result.full.push(sid); return; }  // 満席

    appends.push([nextId(), String(appId), sid, _now(), RES_STATUS_RESERVED, '']);
    counts[sid]      = used + 1;
    slotTaken[s.slot] = sid;
    kept.push(s);
    result.reserved.push(sid);
    result.added.push(sid);
  });

  if (appends.length) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, 6).setValues(appends);
  }
  return result;
}

// data.sessions（JSON配列 or カンマ区切り）を配列にする。壊れていても例外にしない。
function _parseSessionIds(v) {
  if (v === null || v === undefined || v === '') return [];
  if (Object.prototype.toString.call(v) === '[object Array]') {
    return v.map(function (x) { return String(x).trim(); }).filter(function (x) { return x; });
  }
  const s = String(v).trim();
  if (!s) return [];
  if (s.charAt(0) === '[') {
    try {
      const arr = JSON.parse(s);
      if (Object.prototype.toString.call(arr) === '[object Array]') {
        return arr.map(function (x) { return String(x).trim(); }).filter(function (x) { return x; });
      }
    } catch (e) { /* 壊れたJSONは下のカンマ区切り解釈に落とす */ }
  }
  return s.split(',').map(function (x) { return x.trim(); }).filter(function (x) { return x; });
}

// ============================================================
// 🆕 セミナー予約の変更・取消（doPost: action=reserveSessions）
//    ticket_token を知っている本人だけが自分の予約を差し替えられる（capability URL方式・
//    updateApplication と同じ考え方）。pass.html の予約UIから呼ばれる。
//    sessions: [] を送ると全て取消になる。
// ============================================================
function reserveSessions(data) {
  _checkProps();

  const token = String(data.ticket_token || '').trim();
  if (!token) return _err('INVALID_TOKEN');
  const want = _parseSessionIds(data.sessions);

  let result, appId, applicant;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
    const appSh  = _getSheet(ss, SHEET_APPLICATIONS);
    const appRows = appSh.getDataRange().getValues();
    applicant = _findApplicantInRows(appRows, token);
    if (!applicant) return _err('NOT_FOUND');
    // 🔴 キャンセル済みの申込から予約させない。放置すると「来場しない人が席を持つ」ことになる。
    if (applicant.status === 'cancelled') return _err('APPLICATION_CANCELLED');
    appId  = applicant.appId;
    result = _writeReservations(ss, appId, want,
               { replace: true, cancelledAppIds: _cancelledAppIdSet(appRows) });
  } finally {
    lock.releaseLock();
  }

  // 🆕 v0.24.0: 控えメール。
  // 🔴 ロックの外で送る（メール送信は数秒かかる。ロック内に入れると他の予約が待たされる）。
  // 🔴 実際に変わったときだけ送る。同じ内容を保存し直しただけで毎回届くと鬱陶しいため。
  // 🔴 メールが失敗しても予約の保存は成功として返す。メールは控えであって本体ではない。
  let mailSent = false;
  const changed = !!((result.added && result.added.length) || (result.cancelled && result.cancelled.length));
  if (changed && applicant.email) {
    try {
      const cfg   = _guardConfig();
      const quota = cfg.enabled ? _remainingMailQuota() : null;   // 取得失敗時はnull＝送る
      if (quota !== null && quota <= cfg.mailQuotaStop) {
        Logger.log('reserveSessions: 残メール枠が' + quota + '通のため控えメールを送りませんでした app_id=' + appId);
      } else {
        _sendReservationMail(
          applicant.email, applicant.salonName, applicant.staffName,
          SITE_BASE_URL + 'pass.html?t=' + token,
          _reservedSessionLabels(appId),
          !!(result.full && result.full.length));
        mailSent = true;
      }
    } catch (e) {
      Logger.log('予約の控えメール送信に失敗（予約自体は保存済み）: ' + e);
    }
  }

  // 変更後の最新の残席をそのまま返す（画面側で再取得させない＝往復を1回減らす。
  // GASの結果配送は約7%失敗するので、往復回数そのものを減らすことに意味がある）
  const ss2      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const resRows2 = _readReservationRows(ss2);
  // 🔴 v0.32.0: ここでも キャンセル済み申込を除外する。listSessions は除外しているのに
  //    ここだけ渡し忘れていたため、予約直後の画面だけ残席が少なく（＝満席に）見えていた。
  return _ok({
    app_id:    appId,
    result:    result,
    mail_sent: mailSent,   // 画面に「控えをメールでお送りしました」と出すため
    sessions:  _sessionCatalog(ss2, resRows2, _readCancelledAppIds(ss2)),
    reserved:  _reservedSessionIdsOf(resRows2, appId)
  });
}

// ============================================================
// 🆕 マイグレーション: 既存の sessions シートに bullets（J列）/ overview（K列）を追加する
//    【本番シートに対して一度だけ手動実行する】GASエディタで migrateAddSessionDetailColumns
//    を選び▷実行。2回実行しても冪等。
// ============================================================
function migrateAddSessionDetailColumns() {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_SESSIONS);
  if (!sh) { Logger.log('sessionsシートがありません。先に setupSheets() を実行してください。'); return; }

  const j = String(sh.getRange(1, 10).getValue() || '').trim();
  const k = String(sh.getRange(1, 11).getValue() || '').trim();
  if (j === 'bullets' && k === 'overview') { Logger.log('bullets/overview は既にあります。何もしませんでした。'); return; }
  if ((j && j !== 'bullets') || (k && k !== 'overview')) {
    throw new Error('sessionsのJ1/K1に想定外の値があります（"' + j + '" / "' + k + '"）。手動で確認してください。');
  }
  sh.getRange(1, 10, 1, 2).setValues([['bullets', 'overview']]);
  Logger.log('sessionsシートに bullets(J) / overview(K) を追加しました。');
}

// ============================================================
// 🆕 予約枠の投入（手動実行・冪等）
//    2026年のラインナップを sessions シートに入れる。既にある session_id は飛ばすので
//    何度実行してもよい（既存行の内容は上書きしない）。
//
//    🔴 枠を増やす・時間や定員を変えるときは、この関数を書き換えるのではなく
//       **スプレッドシートの sessions シートを直接編集する**こと。
//       この関数は初回投入を楽にするためだけのもの。
//
//    🔴 slot（B列）が同じ枠どうしは排他になる（お客様は1つしか選べない）。
//       バランス革命の計測会は7枠すべて同じ slot なので、1人1枠しか取れない。
// ============================================================
function seedSeminarSessions() {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_SESSIONS);
  if (!sh) { setupSheets(); sh = _getSheet(ss, SHEET_SESSIONS); }
  migrateAddSessionDetailColumns();

  const femcareBullets = [
    '更年期は"予防"できる！',
    '結局更年期の原因ってなにから来ているの？',
    '身体の変化に寄り添ったエッセンス新商品のご紹介',
    '今日からできるフェムケア'
  ].join('\n');

  const femcareOverview =
    '年齢を重ねるにつれて感じる、理由のはっきりしないゆらぎや、すっきりしない毎日。\n' +
    '頑張っているのに気分が晴れない、以前とは違う自分に戸惑う、そんなオトナ女子のために生まれたのが' +
    '「Meguru+（メグルプラス）」です。\n' +
    '「ゆらぐ私に穏やかな巡りを」をコンセプトに、これからの毎日を前向きに、自分らしく過ごすための' +
    '新しい習慣をご提案します。\n' +
    '本セミナーでは、更年期の仕組みとその解決するための秘密を詳しくご紹介。' +
    '変化の時期を、美しく心地よく迎えるためのヒントをお届けします。';

  // 🔵 炭酸ガスパック体験会の説明文は仮置き（2026-09-05時点で本文未定）。決まり次第シートを直接編集する。
  const co2Bullets = [
    '炭酸ガスパックの特徴をご紹介',
    'サロンでの取り入れ方',
    '実際にご体験いただけます'
  ].join('\n');
  const co2Overview = '※この説明文は仮のものです。内容が決まり次第、差し替えます。';

  const balanceBullets = [
    'コンピューター精密計測で、足裏の圧力バランスと骨盤のゆがみをデータで見える化',
    '微細振動波（圧電体）による体軸の体感テスト',
    'お一人ずつのオーダーメイド インソール型骨格調整具のご提案'
  ].join('\n');
  const balanceOverview =
    '長引く「ひざ痛・腰痛・猫背・肩こり・疲労感」の原因は、骨格のゆがみかもしれません。\n' +
    '足裏コンピューター計測で、骨盤と足裏のバランスを精密に測定します。\n' +
    'お一人ずつ丁寧に計測・ご説明を行うため完全予約制です（各回1名様）。';

  // 🔵 フェイスメーカーの説明文は `内容/体験会/フェイスメーカー案内文.docx` から起こした（GitHub管理外）
  const faceBullets = [
    '【顔を動かす】低周波の立体電場をベースに、表情筋などへアプローチ',
    '【整える】温熱・近赤外線などを組み合わせ、肌や筋肉が反応しやすい環境へ',
    '【仕上げる】専用の「ハニカムAgマスク」で顔全体を包み、美容液とのなじみをサポート'
  ].join('\n');
  const faceOverview =
    '＼ 顔は、つくれる。／\n' +
    '「肌」だけでなく、筋肉・フェイスライン・顔全体の印象までトータルにアプローチする' +
    'フェイシャル機器「Face Maker（フェイスメーカー）」をご体験いただけます。\n' +
    '目指すのは、スッキリしたフェイスライン／若々しく整った顔の印象／ハリ・うるおいのある肌印象。\n' +
    'サロンメニュー例は「EFFプローブ 30分」＋「ハニカムAgマスク 15分」の約45分。\n' +
    '「フェイシャルの新メニューを増やしたい」「他店と差別化したい」サロン様におすすめです。';

  // [session_id, slot, title, speaker, room, starts_at, ends_at, capacity, is_active, bullets, overview]
  // 🔴 starts_at / ends_at は文字列で入れる。ends_at が空欄でも画面では「10:00〜」と表示される。
  // 🔵 バランス革命は各回45分（2026-09-04 Takashiさん確認）。申込書には開始時刻しか無いので、
  //    45分を足した終了時刻をここで入れている。変更するときはシートのG列を直す。
  // 🔵 フェイスメーカーは各回50分・10:10スタートの連続7枠（2026-09-04 Takashiさん指定）。
  //    定員は指定が無かったため、1対1のフェイシャル体験であることから**各回1名**とした。
  //    違う場合はシートのH列を直すだけでよい（コード変更・再デプロイ不要）。
  // 🔴 バランス革命とフェイスメーカーは**別の slot** なので、時間が重なる枠を両方予約できてしまう
  //    （例: バランス革命 10:00〜10:45 と フェイスメーカー 10:10〜11:00）。
  //    slot をまたいだ時間の重複チェックは未実装。当日の運用で見るか、実装するかは要判断。
  const rows = [
    // 🔵 2026-09-05 変更: 11:00〜12:00 → 11:30〜12:30（Takashiさん指示）
    ['S1', 'セミナー 第1部', 'オトナ女子の"巡り"に寄り添うフェムケアセミナー', '', '',
     '11:30', '12:30', 20, true, femcareBullets, femcareOverview],   // 🔵 定員20名（2026-09-05）
    // 🔵 2026-09-05 改称: ジャンパーニュ体験セミナー → 整形級！炭酸ガスパック体験会。
    //    🔴 slot（見出し）は「セミナー 第2部」のまま（2026-09-05 Takashiさん指示で確定）。
    //    一度「体験会」に変えたが元に戻した。変えるならシートのB列だけ直せばよい。
    ['S2', 'セミナー 第2部', '整形級！炭酸ガスパック体験会', '', '',
     '14:00', '15:00', 20, true, co2Bullets, co2Overview],   // 🔵 定員20名（2026-09-05）
    ['B1', 'バランス革命 無料計測会', '10:00〜10:45の回', '', '', '10:00', '10:45', 1, true, balanceBullets, balanceOverview],
    ['B2', 'バランス革命 無料計測会', '10:45〜11:30の回', '', '', '10:45', '11:30', 1, true, balanceBullets, balanceOverview],
    ['B3', 'バランス革命 無料計測会', '11:30〜12:15の回', '', '', '11:30', '12:15', 1, true, balanceBullets, balanceOverview],
    ['B4', 'バランス革命 無料計測会', '12:45〜13:30の回', '', '', '12:45', '13:30', 1, true, balanceBullets, balanceOverview],
    ['B5', 'バランス革命 無料計測会', '13:30〜14:15の回', '', '', '13:30', '14:15', 1, true, balanceBullets, balanceOverview],
    ['B6', 'バランス革命 無料計測会', '14:15〜15:00の回', '', '', '14:15', '15:00', 1, true, balanceBullets, balanceOverview],
    ['B7', 'バランス革命 無料計測会', '15:00〜15:45の回', '', '', '15:00', '15:45', 1, true, balanceBullets, balanceOverview],
    ['F1', 'フェイスメーカー 体験会', '10:10〜11:00の回', '', '', '10:10', '11:00', 1, true, faceBullets, faceOverview],
    ['F2', 'フェイスメーカー 体験会', '11:00〜11:50の回', '', '', '11:00', '11:50', 1, true, faceBullets, faceOverview],
    ['F3', 'フェイスメーカー 体験会', '11:50〜12:40の回', '', '', '11:50', '12:40', 1, true, faceBullets, faceOverview],
    ['F4', 'フェイスメーカー 体験会', '12:40〜13:30の回', '', '', '12:40', '13:30', 1, true, faceBullets, faceOverview],
    ['F5', 'フェイスメーカー 体験会', '13:30〜14:20の回', '', '', '13:30', '14:20', 1, true, faceBullets, faceOverview],
    ['F6', 'フェイスメーカー 体験会', '14:20〜15:10の回', '', '', '14:20', '15:10', 1, true, faceBullets, faceOverview],
    ['F7', 'フェイスメーカー 体験会', '15:10〜16:00の回', '', '', '15:10', '16:00', 1, true, faceBullets, faceOverview]
  ];

  const existing = {};
  const cur = sh.getDataRange().getValues();
  for (let i = 1; i < cur.length; i++) {
    const id = String(cur[i][0] == null ? '' : cur[i][0]).trim();
    if (id) existing[id] = true;
  }

  const toAdd = rows.filter(function (r) { return !existing[r[0]]; });
  if (!toAdd.length) { Logger.log('追加する枠はありませんでした（すべて既にあります）。'); return; }

  sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, 11).setValues(toAdd);
  Logger.log('予約枠を' + toAdd.length + '件投入しました: ' +
             toAdd.map(function (r) { return r[0]; }).join(', ') +
             '\n定員・時間・説明文の変更は、以後スプレッドシートを直接編集してください。');
}


// ============================================================
// スプレッドシート初期セットアップ
// GASエディタから手動で一度だけ実行すること
// ============================================================
function setupSheets() {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // --- applications シート（1行 = 来場者1人）---
  let appSh = ss.getSheetByName(SHEET_APPLICATIONS);
  if (!appSh) {
    appSh = ss.insertSheet(SHEET_APPLICATIONS);
    appSh.getRange(1, 1, 1, 23).setValues([[
      'app_id', 'created_at', 'updated_at', 'source',
      'salon_name', 'staff_name', 'email', 'email_norm', 'phone', 'area', // areaは2026-08-06にフォームから削除・列は維持（空文字のみ）
      'has_transaction', 'address', 'referrer', 'agree_capability',
      'line_friend_id', 'line_user_id',
      'ticket_token', 'status', 'checked_in_at', 'note',
      'business_type',          // 🆕 U列（§4-1-2・v0.5.0で追加）
      'request_id',             // 🆕 V列（v0.13.0・§2-2 冪等キー）
      'tantou'                  // 🆕 W列（v0.17.0・営業担当・admin.htmlからのみ書き込む）
    ]]);
    appSh.setFrozenRows(1);
    appSh.setColumnWidth(1, 110);
    Logger.log('applicationsシート作成完了');
  } else {
    Logger.log('applicationsシートは既に存在します');
  }

  // --- sessions シート（セミナー枠マスタ。1行=1枠。枠を増やすのはここに行を足すだけ）---
  let sesSh = ss.getSheetByName(SHEET_SESSIONS);
  if (!sesSh) {
    sesSh = ss.insertSheet(SHEET_SESSIONS);
    sesSh.getRange(1, 1, 1, 11).setValues([[
      'session_id', 'slot', 'title', 'speaker', 'room',
      'starts_at', 'ends_at', 'capacity', 'is_active',
      'bullets', 'overview'   // 🆕 v0.20.0（画面でタップして開く詳細。bulletsは1行1項目）
    ]]);
    sesSh.setFrozenRows(1);
    Logger.log('sessionsシート作成完了');
  } else {
    Logger.log('sessionsシートは既に存在します');
  }

  // --- reservations シート（1行 = 1人 × 1セミナー）---
  let resSh = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!resSh) {
    resSh = ss.insertSheet(SHEET_RESERVATIONS);
    resSh.getRange(1, 1, 1, 6).setValues([[
      'res_id', 'app_id', 'session_id', 'created_at', 'status', 'attended_at'
    ]]);
    resSh.setFrozenRows(1);
    Logger.log('reservationsシート作成完了');
  } else {
    Logger.log('reservationsシートは既に存在します');
  }

  // --- checkins シート（監査用の生ログ）---
  let ckSh = ss.getSheetByName(SHEET_CHECKINS);
  if (!ckSh) {
    ckSh = ss.insertSheet(SHEET_CHECKINS);
    ckSh.getRange(1, 1, 1, 8).setValues([[
      'checkin_id', 'app_id', 'ticket_token', 'session_id',
      'checked_at', 'staff_user_id', 'device', 'note'
    ]]);
    ckSh.setFrozenRows(1);
    Logger.log('checkinsシート作成完了');
  } else {
    Logger.log('checkinsシートは既に存在します');
  }

  // --- mail_log シート ---
  let mlSh = ss.getSheetByName(SHEET_MAIL_LOG);
  if (!mlSh) {
    mlSh = ss.insertSheet(SHEET_MAIL_LOG);
    mlSh.getRange(1, 1, 1, 5).setValues([[
      'sent_at', 'to', 'type', 'status', 'error'
    ]]);
    mlSh.setFrozenRows(1);
    Logger.log('mail_logシート作成完了');
  } else {
    Logger.log('mail_logシートは既に存在します');
  }

  // --- config シート（イベント情報。コードを触らず運用で変更できる）---
  let cfgSh = ss.getSheetByName(SHEET_CONFIG);
  if (!cfgSh) {
    cfgSh = ss.insertSheet(SHEET_CONFIG);
    cfgSh.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    cfgSh.getRange(2, 1, 6, 2).setValues([
      ['event_date',     '2026年10月26日（月）'],
      ['event_time',     '10:00〜16:00'],
      ['venue_name',     '青島屋（AOSHIMAYA）'],
      ['venue_addr',     '宮崎市青島2丁目12-11'],
      ['apply_deadline', '2026年10月20日（火）'],
      ['mail_from',      MAIL_FROM_ADDR]
    ]);
    cfgSh.setFrozenRows(1);
    Logger.log('configシート作成完了（初期値を投入済み）');
  } else {
    Logger.log('configシートは既に存在します');
  }

  // --- line_friends_cache シート（🆕 L1-a。LINE Harness友だちのキャッシュ）---
  // 本番の既存シートには_ensureLineFriendsCacheSheet()が初回syncLineFriends()実行時に
  // 自動作成するため、setupSheetsでの作成は「新規シートを最初から作る場合」の網羅目的
  if (!ss.getSheetByName(SHEET_LINE_FRIENDS_CACHE)) {
    _ensureLineFriendsCacheSheet(ss);
  } else {
    Logger.log('line_friends_cacheシートは既に存在します');
  }

  // --- line_sync_log シート（🆕 L2。プッシュ・タグ付与・metadata書き戻しの失敗ログ）---
  // 本番の既存シートには_ensureLineSyncLogSheet()が初回_logLineSync()実行時に自動作成される
  if (!ss.getSheetByName(SHEET_LINE_SYNC_LOG)) {
    _ensureLineSyncLogSheet(ss);
  } else {
    Logger.log('line_sync_logシートは既に存在します');
  }

  // --- spare_badges シート（🆕 名札印刷。予備名札プール）---
  // 本番の既存シートには_ensureSpareBadgesSheet()がsetupSpareBadges()初回実行時に自動作成される。
  // setupSheetsでの作成は「新規シートを最初から作る場合」の網羅目的（他の_ensure*系と同じ書き方）
  if (!ss.getSheetByName(SHEET_SPARE_BADGES)) {
    _ensureSpareBadgesSheet(ss);
  } else {
    Logger.log('spare_badgesシートは既に存在します');
  }

  // --- booth系4シート（🆕 v0.19.0。booths / booth_products / orders / order_items）---
  // 既存シートには触らない（_ensureBoothSheetsが冪等）。単独で作るなら setupBoothSheets() を実行する
  _ensureBoothSheets(ss);

  Logger.log('setupSheets完了');
}

// ============================================================
// 🆕 マイグレーション: 既存の applications シートに business_type 列（U列）を追加する
// 2026-08-05・設計書v3（§4-1-2）対応。
// 【本番シートに対して一度だけ手動実行すること】GASエディタの関数選択で
// migrateAddBusinessType を選び、▷実行する。setupSheetsとは別に必要（setupSheetsは
// シートが既に存在する場合は何もしないため、既存シートへの列追加はこの関数が担う）。
// 新規にsetupSheetsでシートを作る場合はヘッダーに最初から含まれるため実行不要。
// ============================================================
function migrateAddBusinessType() {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = _getSheet(ss, SHEET_APPLICATIONS);
  const header = sh.getRange(1, 21).getValue();
  if (header === 'business_type') {
    Logger.log('business_type列は既にU1に存在します。何もしませんでした。');
    return;
  }
  if (header) {
    throw new Error('U1に想定外の値が入っています（"' + header + '"）。手動で確認してください。');
  }
  sh.getRange(1, 21).setValue('business_type');
  Logger.log('business_type列をU1に追加しました。');
}

// ============================================================
// 🆕 マイグレーション: 既存の applications シートに request_id 列（V列）を追加する
// 2026-08-06・配送障害対策P1（GAS配送障害_対策計画.md §2-P1・§2-2）対応。
// 【本番シートに対して一度だけ手動実行すること】GASエディタの関数選択で
// migrateAddRequestIdColumn を選び、▷実行する。setupSheetsとは別に必要（既存シートには
// 自動で列が増えないため）。新規にsetupSheetsでシートを作る場合はヘッダーに
// 最初から含まれるため実行不要。2回実行しても冪等（既にV1にあれば何もしない）。
// ============================================================
function migrateAddRequestIdColumn() {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = _getSheet(ss, SHEET_APPLICATIONS);
  const header = sh.getRange(1, 22).getValue();
  if (header === 'request_id') {
    Logger.log('request_id列は既にV1に存在します。何もしませんでした。');
    return;
  }
  if (header) {
    throw new Error('V1に想定外の値が入っています（"' + header + '"）。手動で確認してください。');
  }
  sh.getRange(1, 22).setValue('request_id');
  Logger.log('request_id列をV1に追加しました。');
}

// ============================================================
// 🆕 マイグレーション: applications シートの電話番号列（I列=9列目）の先頭「0」落ちを修正する
// 2026-08-06発覚（pass.htmlの編集フォームで初めて可視化された）。
//
// 原因: 全て数字の文字列をAutomatic書式の列に書き込むと、Google Sheetsが数値型に
// 変換してしまい先頭の0が消える（appendRow・setValuesとも、UIでの手入力と同じ挙動）。
// applyApplication/updateApplication は本マイグレーション以降、書き込み直前に対象セルを
// Plain Text指定するよう修正済み（v0.8.1）。既存行の復旧はこの関数が担う。
//
// 【本番シートに対して一度だけ手動実行すること】GASエディタの関数選択で
// migrateFixPhoneColumn を選び、▷実行する。
// ============================================================
function migrateFixPhoneColumn() {
  _checkProps();
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh  = _getSheet(ss, SHEET_APPLICATIONS);
  const rows = sh.getDataRange().getValues();

  // 列全体をPlain Textにしておく（今後、手入力や別経路で書き込まれても先頭0が消えなくなる）
  sh.getRange(2, 9, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');

  let fixed = 0;
  for (let i = 1; i < rows.length; i++) {
    const v = rows[i][8];
    // セルの実体が数値型＝先頭の0が消えている状態（日本の電話番号は必ず0始まりなので復元できる）
    if (typeof v === 'number') {
      sh.getRange(i + 1, 9, 1, 1).setValue('0' + String(v));
      fixed++;
    }
  }
  Logger.log('電話番号列の書式修正が完了しました。復元した行数: ' + fixed + '（列全体もPlain Text化済み）');
}

// ============================================================
// 🆕 名札印刷badges.html用（S2・名札印刷_badges設計.md §5-5・§5-6）
// ============================================================

// spare_badges シートが無ければヘッダー付きで作成する（line_friends_cache等と同じ書き方）
function _ensureSpareBadgesSheet(ss) {
  let sh = ss.getSheetByName(SHEET_SPARE_BADGES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SPARE_BADGES);
    sh.getRange(1, 1, 1, 4).setValues([[
      'spare_no', 'ticket_token', 'assigned_app_id', 'assigned_at'
    ]]);
    sh.setFrozenRows(1);
    Logger.log('spare_badgesシート作成完了');
  }
  return sh;
}

// 予備名札を count 件（既定20）発行する。【本番シートに対して手動実行すること】
// 🔴 既存行は絶対に上書きしない。P-01が既にあればP-21から続けて追加する（冪等）。
function setupSpareBadges(count) {
  _checkProps();
  count = count || 20;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = _ensureSpareBadgesSheet(ss);

  const rows = sh.getDataRange().getValues();
  let maxSeq = 0;
  for (let i = 1; i < rows.length; i++) {
    const m = String(rows[i][0] || '').match(/^P-(\d+)$/);
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > maxSeq) maxSeq = n;
    }
  }

  const newRows = [];
  for (let i = 1; i <= count; i++) {
    const seq = maxSeq + i;
    newRows.push(['P-' + ('00' + seq).slice(-2), _genToken(), '', '']);
  }

  if (newRows.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 4).setValues(newRows);
  }

  Logger.log('setupSpareBadges完了: ' + newRows[0][0] + '〜' + newRows[newRows.length - 1][0] + ' を作成しました（' + newRows.length + '件）。');
}

// config シートに業態→色帯マッピングの既定値を投入する（無ければ追加・既存値は上書きしない）。
// 【本番シートに対して手動実行すること】badges.html §5-6・§14-3。
// 🔴 色はコードに直書きしない方針の実体化。A/B/Cの実際の色は用紙が決まってから
// たかしさんがconfigシートを書き換えるだけで変わる。ここに入れるのは仮置きの既定値。
function seedBadgeConfig() {
  _checkProps();
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh  = _getSheet(ss, SHEET_CONFIG);
  const cfg = _getConfig();

  const defaults = [
    ['badge_color_美容室',       'A'],
    ['badge_color_理容室',       'A'],
    ['badge_color_エステサロン', 'B'],
    ['badge_color_ネイルサロン', 'B'],
    ['badge_color_アイサロン',   'B'],
    ['badge_color_その他',       'C'],
    ['badge_color_default',     'C'],
    ['badge_band_A',  '#E8542F'],
    ['badge_band_B',  '#2FA8CC'],
    ['badge_band_C',  '#7A8B99'],
    ['badge_label_A', '理美容'],
    ['badge_label_B', '美容サロン'],
    ['badge_label_C', 'その他']
  ];

  let added = 0;
  defaults.forEach(function (pair) {
    if (cfg[pair[0]] !== undefined) return; // 既にある値は上書きしない
    sh.appendRow(pair);
    added++;
  });

  Logger.log('seedBadgeConfig完了: ' + added + '件を追加しました（既存の値は変更していません）。');
}

// ============================================================
// 本日のメール送信残数を確認する（診断用・手動実行）
// 2026-08-04実測: GAS実行アカウントはGoogle Workspace Business Standard契約にも
// かかわらず、Apps Script上では個人アカウント扱いで100通/日だった（原因未特定）
// ============================================================
function checkMailQuota() {
  Logger.log('本日の残りメール送信可能数: ' + MailApp.getRemainingDailyQuota());
}

// ============================================================
// 🆕 診断用アクション（diag.html から呼ぶ・v0.12.0）
//
// 目的: 2026-08-06に判明した「GASの実行ログは完了(1〜3秒)なのに、その結果が
// ブラウザに一度も届かない」現象の**実測**。実行ログ(実行数画面)の記録と、
// クライアント側で受け取れた件数を突き合わせて、配送層の失敗率を数える。
//
// 🔴 いずれも読み取りのみ。データを一切変更しない。
//   pingLight : スプレッドシートに触らず即座に返す → 配送層そのものの失敗率
//   pingHeavy : applicationsシートを読んでから返す（liffPrefillと同等の処理量）
//               → 処理時間が失敗率の引き金になっているかの切り分け
// ============================================================
function pingLight(data) {
  return _ok({
    mode: 'light',
    version: VERSION,
    seq: String((data && data.seq) || ''), // クライアント側の試行番号をそのまま返す（対応付け用）
    server_time: _now()
  });
}

function pingHeavy(data) {
  _checkProps();
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh   = _getSheet(ss, SHEET_APPLICATIONS);
  const rows = sh.getDataRange().getValues(); // liffPrefillと同じ「全行読み込み」を再現
  return _ok({
    mode: 'heavy',
    version: VERSION,
    seq: String((data && data.seq) || ''),
    row_count: rows.length - 1,
    server_time: _now()
  });
}

// ============================================================
// 🆕 S0（名札印刷badges.html・事前チェック専用）: beaufield-auth を開けるかだけを確かめる。
// 名札印刷_badges設計.md §0-2。認証基盤(validateSession等)を実装する前に、
// beaufesのGAS実行アカウントが beaufield-auth スプレッドシートを開けることを確認する。
// 確認できたら削除せず残してよい（診断用として無害）。
// 🔴 PIN・氏名・トークンは絶対にログへ出さない（行数だけ見る）
// ============================================================
function testAuthSheetAccess() {
  const id = PropertiesService.getScriptProperties().getProperty('AUTH_SHEET_ID');
  if (!id) { Logger.log('NG: AUTH_SHEET_ID が未設定'); return; }
  const ss = SpreadsheetApp.openById(id);
  ['users', 'user_app_roles', 'sessions'].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    Logger.log(n + ': ' + (sh ? sh.getLastRow() + '行' : '🔴 シートが無い'));
  });
}

// ============================================================
// 🆕 v0.17.0（2026-08-26）: 申込者一覧 admin.html 用（閲覧＋担当欄）
//
// ・閲覧（listApplications）と担当欄の書き込み（setTantou）だけを足した最小版。
//   申込フォーム側の経路（apply / applyLiff / updateApplication / getPass / resendPass）には
//   一切手を入れていない。担当は applications の W列（＝A〜V の 22 列の次）にのみ書く。
// ・🔴 setTantou は updated_at（C列）を触らない。updated_at は「申込内容が更新された時刻」で
//   badges.html の日付フィルタ（date_field=updated_at）が参照するため、担当の付け替えで
//   動かすと名札の抽出条件が壊れる。
// ・担当者の選択肢は config シートの `tantou_list`（例: `前島,佐藤,田中`）で管理する。
//   コードを触らずスプレッドシート側だけで増減できる（seedTantouList() でも投入可）。
// ・認証は badges.html と同じ beaufield-auth の共通セッション（_requireSession）。
// ============================================================

const COL_TANTOU = 23; // applications の W列（1始まり）

// applications シートに W列（tantou）を用意する。無ければ列とヘッダーを作る。
// 既存の本番シートは A〜V の 22 列で作られているため、この関数が実質のマイグレーションを兼ねる
// （他の _ensure*Sheet 系と同じ「呼ばれた時に自己修復する」方式・§0-1）。
function _ensureTantouColumn(sh) {
  if (sh.getMaxColumns() < COL_TANTOU) {
    sh.insertColumnsAfter(sh.getMaxColumns(), COL_TANTOU - sh.getMaxColumns());
  }
  const head = String(sh.getRange(1, COL_TANTOU).getValue() || '').trim();
  if (!head) sh.getRange(1, COL_TANTOU).setValue('tantou');
  return sh;
}

// config の tantou_list を配列にして返す。区切りはカンマ（全角・読点・改行も許容）。
function _parseTantouList(cfg) {
  const raw = String((cfg || {}).tantou_list || '');
  const normalized = raw
    .split('、').join(',')   // 、
    .split('，').join(',')   // ，
    .split('\r').join(',')
    .split('\n').join(',');
  const out = [];
  normalized.split(',').forEach(function (s) {
    const t = String(s).trim();
    if (t && out.indexOf(t) < 0) out.push(t);
  });
  return out;
}

// 申込一覧を返す（doPost: action=listApplications）。data: { session_token }
// 🔴 絞り込み・集計は画面側で行う（件数が数百なので全件返して即時フィルタするほうが速く、
//    GAS側の分岐も増えない）。ticket_token は返さない（入場パスの鍵そのものなので一覧に不要）。
function listApplications(data) {
  const auth = _requireSession(data);
  if (!auth.ok) return _err(auth.error);
  _checkProps();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = _ensureTantouColumn(_getSheet(ss, SHEET_APPLICATIONS));
  const rows = sh.getDataRange().getValues();
  const cfg  = _getConfig();

  const list = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    list.push({
      app_id:          String(rows[i][0]),
      created_at:      String(rows[i][1]),
      created_key:     _dateKey(rows[i][1]),
      updated_at:      String(rows[i][2]),
      source:          String(rows[i][3]),
      salon_name:      String(rows[i][4]),
      staff_name:      String(rows[i][5]),
      email:           String(rows[i][6]),
      phone:           String(rows[i][8]),
      has_transaction: String(rows[i][10]),
      address:         String(rows[i][11]),
      referrer:        String(rows[i][12]),
      has_line:        !!String(rows[i][14] || ''),
      status:          String(rows[i][17]),
      checked_in_at:   String(rows[i][18] || ''),
      note:            String(rows[i][19] || ''),
      business_type:   String(rows[i][20] || ''),
      tantou:          String(rows[i][COL_TANTOU - 1] || '')
    });
  }

  list.sort(function (a, b) { return a.app_id < b.app_id ? -1 : (a.app_id > b.app_id ? 1 : 0); });

  return _ok({
    total:            list.length,
    tantou_list:      _parseTantouList(cfg),
    business_types:   BUSINESS_TYPE_OPTIONS,
    server_version:   VERSION,
    viewer:           auth.session.name || auth.session.user_id || '',
    applications:     list
  });
}

// 担当をまとめて設定する（doPost: action=setTantou）。
// data: { session_token, app_ids: 'F2026-0001,F2026-0002', tantou: '前島' }
// tantou が空文字なら「未割当に戻す」。app_id は F2026-0001 形式でカンマを含まない。
function setTantou(data) {
  const auth = _requireSession(data);
  if (!auth.ok) return _err(auth.error);
  _checkProps();

  const rawIds = Array.isArray(data.app_ids) ? data.app_ids.join(',') : String(data.app_ids || '');
  const appIds = [];
  rawIds.split(',').forEach(function (s) {
    const t = String(s).trim();
    if (t && appIds.indexOf(t) < 0) appIds.push(t);
  });
  if (!appIds.length)     return _err('担当を設定する申込が選ばれていません');
  if (appIds.length > 500) return _err('一度に設定できるのは500件までです');

  const tantou = String(data.tantou || '').trim();
  const allowed = _parseTantouList(_getConfig());
  // 🔴 リストに無い名前は弾く。ここを緩めると表記ゆれ（「前島」「前島崇志」）で集計が割れる。
  if (tantou && allowed.indexOf(tantou) < 0) {
    return _err('config シートの tantou_list にない担当者です: ' + tantou);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = _ensureTantouColumn(_getSheet(ss, SHEET_APPLICATIONS));
    const rows = sh.getDataRange().getValues();
    if (rows.length < 2) return _err('申込データがありません');

    const want = {};
    appIds.forEach(function (id) { want[id] = true; });

    // 🔴 W列を丸ごと読み込み→書き戻す（1行ずつのsetValueだと件数分の書き込みが走るため）。
    //    対象外の行は今の値をそのまま書き戻すので内容は変わらない。
    //    ロック中に追記された行があっても、書き戻す範囲は読み込んだ行数までなので踏まない。
    const col = [];
    const seen = {};
    let updated = 0;
    for (let i = 1; i < rows.length; i++) {
      const id  = String(rows[i][0] || '');
      let value = String(rows[i][COL_TANTOU - 1] || '');
      if (id && want[id]) { value = tantou; seen[id] = true; updated++; }
      col.push([value]);
    }
    sh.getRange(2, COL_TANTOU, col.length, 1).setValues(col);

    const notFound = appIds.filter(function (id) { return !seen[id]; });
    return _ok({ updated: updated, tantou: tantou, not_found: notFound });
  } finally {
    lock.releaseLock();
  }
}

// 【手動実行用】config シートに tantou_list を投入する。
// GASエディタの関数選択で seedTantouList を選び ▷実行（引数なしなら既定値が入る）。
// 既に値が入っている場合は上書きしない（運用中の設定を潰さないため）。
// 担当者の増減はスプレッドシートの config シートを直接編集すればよい。
function seedTantouList(namesCsv) {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = _getSheet(ss, SHEET_CONFIG);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === 'tantou_list') {
      if (String(rows[i][1] || '').trim()) {
        Logger.log('tantou_list は既に設定済みです: ' + rows[i][1]);
        return;
      }
      sh.getRange(i + 1, 2).setValue(String(namesCsv || ''));
      Logger.log('tantou_list を更新しました: ' + namesCsv);
      return;
    }
  }
  sh.appendRow(['tantou_list', String(namesCsv || '')]);
  Logger.log('tantou_list を追加しました: ' + namesCsv);
}

// ============================================================
// 🆕 booth（ブース購買記録）v0.19.0
// 実装仕様の正本: 開発・自動化/beaufes/booth実装設計_確定版.md
//   §3 データモデル / §4 API契約 / §5 サーバー実装 / §2 各項目の根拠
// 上位の設計思想・当日運用は 当日運用_堅牢化設計.md §4 と正本§15。
//
// 🔴 3つの不変条件（迷ったらここへ戻る・§1）
//   I1 端末が申告した値（*_raw）とサーバーが解決した値（app_id 等）を列レベルで分ける
//   I2 write_state(PENDING→COMPLETE) と status(active→voided) は直交・どちらも前にしか進まない
//   I3 現実に発生した取引は消さない。原則は「拒否」ではなく「受理してフラグを立てる」
// ============================================================

const SHEET_BOOTHS         = 'booths';
const SHEET_BOOTH_PRODUCTS = 'booth_products';
const SHEET_ORDERS         = 'orders';
const SHEET_ORDER_ITEMS    = 'order_items';

const BOOTH_ORDER_COLS = 25; // orders は A〜Y の25列（§3-3）
const BOOTH_ITEM_COLS  = 10; // order_items は A〜J の10列（§3-4）

// 🔴 ロック内で1回だけ読む範囲。A〜F が冪等判定、G(expected)・H(resolve_status) が
// 再送時の分岐に要る（§5-3）。8列とも連続しているので getRange は1回のまま（§5-2）。
const BOOTH_KEY_COLS = 8;

// orders の列番号（1始まり）。🔴 A〜F の並びには意味がある。列を足すときは必ず末尾（Z以降）へ。
const BO = {
  order_id: 1, booth_id: 2, idempotency_key: 3, payload_hash: 4, write_state: 5, status: 6,
  expected_item_count: 7, resolve_status: 8, app_id: 9, ticket_token: 10, ticket_token_raw: 11,
  entry_code_raw: 12, input_method: 13, client_created_at: 14, received_at: 15, created_at: 16,
  client_instance_id: 17, master_version: 18, validation_state: 19, line_notify_state: 20,
  line_notified_at: 21, void_idempotency_key: 22, voided_at: 23, last_error: 24, note: 25
};

// 入力上限（§4-2）。超えたら INVALID_REQUEST（端末は恒久失敗として扱う）
const BOOTH_MAX_ITEMS         = 30;
const BOOTH_MAX_QTY           = 99;
const BOOTH_MAX_KEY_LEN       = 64;
const BOOTH_MAX_REQUEST_CHARS = 16384;
const BOOTH_LOCK_WAIT_MS      = 45000;
const BOOTH_CLOCK_SKEW_MS     = 10 * 60 * 1000;

// ------------------------------------------------------------
// シート
// ------------------------------------------------------------

// booth用シート4枚を冪等に用意して返す（既存があれば触らない・_ensureSpareBadgesSheetと同じ形）。
function _ensureBoothSheets(ss) {
  let bo = ss.getSheetByName(SHEET_BOOTHS);
  if (!bo) {
    bo = ss.insertSheet(SHEET_BOOTHS);
    bo.getRange(1, 1, 1, 5).setValues([[
      'booth_id', 'maker_name', 'booth_token', 'is_active', 'note'
    ]]);
    bo.setFrozenRows(1);
    // トークンは英数字なので指数表記に化けうる。列ごとPlain Textにしておく（電話番号列と同じ理由）
    bo.getRange(1, 3, bo.getMaxRows(), 1).setNumberFormat('@');
    Logger.log('boothsシート作成完了');
  }

  let bp = ss.getSheetByName(SHEET_BOOTH_PRODUCTS);
  if (!bp) {
    bp = ss.insertSheet(SHEET_BOOTH_PRODUCTS);
    bp.getRange(1, 1, 1, 7).setValues([[
      'product_id', 'booth_id', 'product_name', 'spec', 'unit_price', 'sort_order', 'is_active'
    ]]);
    bp.setFrozenRows(1);
    Logger.log('booth_productsシート作成完了');
  }

  let or = ss.getSheetByName(SHEET_ORDERS);
  if (!or) {
    or = ss.insertSheet(SHEET_ORDERS);
    or.getRange(1, 1, 1, BOOTH_ORDER_COLS).setValues([[
      'order_id', 'booth_id', 'idempotency_key', 'payload_hash', 'write_state', 'status',
      'expected_item_count', 'resolve_status', 'app_id', 'ticket_token', 'ticket_token_raw',
      'entry_code_raw', 'input_method', 'client_created_at', 'received_at', 'created_at',
      'client_instance_id', 'master_version', 'validation_state', 'line_notify_state',
      'line_notified_at', 'void_idempotency_key', 'voided_at', 'last_error', 'note'
    ]]);
    or.setFrozenRows(1);
    // 🔴 L列（entry_code_raw）は 'P-07' が日付に化けないようPlain Text（§3-5）。
    // K列（ticket_token_raw）・J列（ticket_token）も英数字トークンなので同じ扱いにする。
    or.getRange(1, BO.ticket_token,     or.getMaxRows(), 3).setNumberFormat('@');
    Logger.log('ordersシート作成完了');
  }

  let oi = ss.getSheetByName(SHEET_ORDER_ITEMS);
  if (!oi) {
    oi = ss.insertSheet(SHEET_ORDER_ITEMS);
    oi.getRange(1, 1, 1, BOOTH_ITEM_COLS).setValues([[
      'item_id', 'order_id', 'booth_id', 'product_id', 'product_name', 'spec',
      'qty', 'unit_price', 'delivery', 'created_at'
    ]]);
    oi.setFrozenRows(1);
    Logger.log('order_itemsシート作成完了');
  }

  return { booths: bo, products: bp, orders: or, items: oi };
}

// 【本番シートに対して手動実行する】booth用シート4枚を作る（STEP 1）。冪等。
function setupBoothSheets() {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  _ensureBoothSheets(ss);
  Logger.log('setupBoothSheets完了');
}

// 【手動実行】動作確認用のテストブース1件と商品3件を入れる（STEP 1）。
// 既に B99 があれば何もしない。🔴 本番のブースは B01〜 を手で作る。
function seedTestBooth() {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = _ensureBoothSheets(ss);

  const rows = sheets.booths.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === 'B99') {
      Logger.log('テストブース B99 は既に存在します。booth_token=' + rows[i][2]);
      return;
    }
  }

  const token = _genBoothToken();
  sheets.booths.getRange(sheets.booths.getLastRow() + 1, 1, 1, 5)
    .setValues([['B99', 'テスト商事（動作確認用）', token, true, '動作確認用。当日までに is_active=FALSE にする']]);

  const pRows = [
    ['B99-001', 'B99', 'テストシャンプー', '1000ml', 3000, 1, true],
    ['B99-002', 'B99', 'テストトリートメント', '1000g', 4200, 2, true],
    ['B99-003', 'B99', 'テストカラー剤', '80g',    900, 3, true]
  ];
  sheets.products.getRange(sheets.products.getLastRow() + 1, 1, pRows.length, 7).setValues(pRows);

  Logger.log('seedTestBooth完了: B99 / booth_token=' + token);
  Logger.log('確認URL例: <WebアプリURL>?action=boothInit&data=' +
    encodeURIComponent(JSON.stringify({ b: token })));
}

// ============================================================
// 🆕 v0.32.0 検証用: 予約まわりの動作確認に使うテスト申込を1行だけ作る。
//    【GASエディタから手動で実行する】
//    🔴 applyApplication を通らないので、全社員へのLINE WORKS通知は飛ばない。
//    🔴 2回実行しても増えない（既にあれば ticket_token をログに出すだけ）。
//    🔴 確認が終わったら必ず deleteTestApplication() で消すこと
//       （消さないと申込者一覧・名札の枚数に混ざる）。
//    控えメールの宛先はスクリプトプロパティ TEST_MAIL_TO、無ければ実行者のアドレス。
//    🔴 コードに個人のメールアドレスは書かない（公開リポジトリのため）。
// ============================================================
const TEST_APP_ID = 'TEST-RES';

function seedTestApplication() {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 🔴 reservations シートが無いと予約が一切書けない（_writeReservations が invalid を返す）
  let resSh = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!resSh) {
    resSh = ss.insertSheet(SHEET_RESERVATIONS);
    resSh.getRange(1, 1, 1, 6).setValues([[
      'res_id', 'app_id', 'session_id', 'created_at', 'status', 'attended_at'
    ]]);
    resSh.setFrozenRows(1);
    Logger.log('reservations シートを作成しました。');
  } else {
    Logger.log('reservations シートは既にあります。');
  }

  const sh   = _getSheet(ss, SHEET_APPLICATIONS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === TEST_APP_ID) {
      Logger.log('テスト申込は既にあります。ticket_token=' + rows[i][16]);
      return;
    }
  }

  // 🔴 Session.getActiveUser().getEmail() は使わない。
  //    OAuthスコープ（userinfo.email）が増えてプロジェクト全体の再承認ダイアログが出るため。
  //    本番Webアプリの認可に触りたくないので、宛先はスクリプトプロパティだけで決める。
  //    未設定なら空欄のまま作る（予約の検証はできる。控えメールを見たいときだけ設定する）。
  const email = PropertiesService.getScriptProperties().getProperty('TEST_MAIL_TO') || '';
  const token  = _genToken();
  const now    = _now();
  const newRow = rows.length + 1;
  sh.getRange(newRow, 9, 1, 1).setNumberFormat('@');   // 電話番号列はPlain Text（既存の作法に合わせる）
  sh.getRange(newRow, 1, 1, 23).setValues([[
    TEST_APP_ID, now, now, 'test',
    'テストサロン（動作確認用）', '検証 用', email, String(email).toLowerCase(), '', '',
    'なし', '', '', true,
    '', '',
    token, 'confirmed', '', '🔴 動作確認用。終わったら deleteTestApplication() で消すこと',
    '美容室', '', ''
  ]]);

  Logger.log('seedTestApplication 完了: app_id=' + TEST_APP_ID + ' / ticket_token=' + token);
  Logger.log(email
    ? '控えメールの宛先: ' + email
    : '控えメールの宛先は空欄です。メールも確認するなら、スクリプトプロパティ TEST_MAIL_TO に'
      + 'アドレスを入れてから deleteTestApplication() → seedTestApplication() の順で作り直すか、'
      + 'applications シートの ' + TEST_APP_ID + ' の行のG列にアドレスを直接入れてください。');
}

// テスト申込と、その予約行を消す。何度実行しても安全。
function deleteTestApplication() {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const resSh = ss.getSheetByName(SHEET_RESERVATIONS);
  if (resSh) {
    const rows = resSh.getDataRange().getValues();
    let n = 0;
    for (let i = rows.length - 1; i >= 1; i--) {   // 🔴 下から消す（先に上を消すと行番号がずれる）
      if (String(rows[i][RES_COL.appId]).trim() === TEST_APP_ID) { resSh.deleteRow(i + 1); n++; }
    }
    Logger.log('テスト予約を' + n + '行削除しました。');
  }

  const sh   = _getSheet(ss, SHEET_APPLICATIONS);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]).trim() === TEST_APP_ID) {
      sh.deleteRow(i + 1);
      Logger.log('テスト申込を削除しました。');
      return;
    }
  }
  Logger.log('テスト申込は見つかりませんでした（既に削除済み）。');
}

// ------------------------------------------------------------
// 小道具
// ------------------------------------------------------------

// ブース用トークン（ランダム24文字・§3-1）
function _genBoothToken() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 24);
}

function _boothPad(n, width) {
  let s = String(n);
  while (s.length < width) s = '0' + s;
  return s;
}

// 🔴 「明示的にFALSEでない限り有効」とする。空欄を無効扱いにすると、手作りのマスタで
// is_active を書き忘れただけで当日その商品が1件も出なくなる（黙って消える）。
// 無効化は運用ルールどおり FALSE を書く操作でのみ起きる（§7ルール1）。
function _boothIsActive(v) {
  if (v === false) return false;
  const s = String(v == null ? '' : v).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'NO' || s === 'N' || s === 'いいえ');
}

// 入力キーの正規化（§2-3 手順1）。NFKC＋trim。ecはさらに大文字化して使う。
function _boothNormKey(s) {
  let t = String(s == null ? '' : s).trim();
  if (t && t.normalize) t = t.normalize('NFKC').trim();
  return t;
}

// 'yyyy-MM-dd HH:mm:ss' をミリ秒に。読めなければ0（判定をスキップする）
function _boothParseTs(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})/);
  if (!m) return 0;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]),
                  Number(m[4]), Number(m[5]), Number(m[6])).getTime();
}

function _boothClockSkew(clientCreatedAt, receivedAt) {
  const c = _boothParseTs(clientCreatedAt);
  const r = _boothParseTs(receivedAt);
  if (!c || !r) return false;
  return Math.abs(r - c) > BOOTH_CLOCK_SKEW_MS;
}

function _boothFlags(list) {
  const out = [];
  (list || []).forEach(function (f) {
    if (f && out.indexOf(f) < 0) out.push(f);
  });
  return out.join(',');
}

// 🔴 冪等判定の正準ハッシュ（§2-3）。対象は booth_id ＋ 生の入力キー ＋ 明細 ＋ delivery のみ。
// master_version・時刻・端末ID・解決後のapp_id・商品名/単価は**入れない**
// （入れると、再送の合間に状況が変わっただけで正常な再送が CONFLICT になる）。
function _boothPayloadHash(boothId, tt, ec, delivery, items) {
  const body = boothId + '|' + tt + '|' + ec + '|' + delivery + '|' +
    items.map(function (i) { return i.product_id + ':' + i.qty; }).join(',');
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, body, Utilities.Charset.UTF_8);
  return raw.map(function (b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

// 明細の正規化（§2-3 手順2）: qty整数化 → 0以下を捨てる → 同一product_idを合算 → product_id昇順。
// 🔴 並べ替えは localeCompare を使わない（ロケール依存でハッシュがぶれる）。
function _boothCanonicalizeItems(items) {
  if (!Array.isArray(items)) return { error: 'INVALID_REQUEST' };
  if (items.length > BOOTH_MAX_ITEMS) return { error: 'INVALID_REQUEST' };

  const merged = {};
  for (let i = 0; i < items.length; i++) {
    const it = items[i] || {};
    const pid = String(it.product_id == null ? '' : it.product_id).trim();
    if (!pid || pid.length > BOOTH_MAX_KEY_LEN) return { error: 'INVALID_REQUEST' };
    const qty = Number(it.qty);
    if (!isFinite(qty) || Math.floor(qty) !== qty) return { error: 'INVALID_REQUEST' };
    if (qty > BOOTH_MAX_QTY) return { error: 'INVALID_REQUEST' };
    if (qty <= 0) continue; // 正規化で捨てる（エラーにはしない）
    merged[pid] = (merged[pid] || 0) + qty;
  }

  const ids = Object.keys(merged).sort(function (a, b) {
    return a < b ? -1 : (a > b ? 1 : 0);
  });
  if (ids.length === 0) return { error: 'NO_ITEMS' };

  return { items: ids.map(function (pid) { return { product_id: pid, qty: merged[pid] }; }) };
}

// booth_token → ブース。🔴 他ブースの情報は一切返さない。
function _boothAuth(sheets, boothToken) {
  const token = String(boothToken || '').trim();
  if (!token) return { ok: false, error: 'BOOTH_NOT_FOUND' };

  const rows = sheets.booths.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).trim() !== token) continue;
    return {
      ok: true,
      booth_id: String(rows[i][0]).trim(),
      maker_name: String(rows[i][1] || ''),
      is_active: _boothIsActive(rows[i][3])
    };
  }
  return { ok: false, error: 'BOOTH_NOT_FOUND' };
}

// 商品の検証＋スナップショット（§2-6）。🔴 商品名・規格・単価はサーバーがマスタから取る。
function _boothSnapshotProducts(sheets, boothId, canonItems) {
  const rows = sheets.products.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    const pid = String(rows[i][0]).trim();
    if (!pid) continue;
    map[pid] = {
      booth_id: String(rows[i][1]).trim(),
      product_name: String(rows[i][2] || ''),
      spec: String(rows[i][3] || ''),
      unit_price: (rows[i][4] === '' || rows[i][4] == null) ? '' : rows[i][4],
      is_active: _boothIsActive(rows[i][6])
    };
  }

  const out = [];
  let stale = false;
  for (let k = 0; k < canonItems.length; k++) {
    const c = canonItems[k];
    const p = map[c.product_id];
    if (!p) return { error: 'PRODUCT_NOT_FOUND' };
    if (p.booth_id !== boothId) return { error: 'PRODUCT_OTHER_BOOTH' };
    if (!p.is_active) stale = true; // 受理してフラグ（I3）
    out.push({
      product_id: c.product_id, qty: c.qty,
      product_name: p.product_name, spec: p.spec, unit_price: p.unit_price
    });
  }
  return { items: out, stale: stale };
}

// ------------------------------------------------------------
// 本人解決（§5-4）
// ------------------------------------------------------------

// applications と spare_badges を1回ずつだけ読んで索引を作る。
// 🔴 ロックの外で呼ぶこと（一番重い読み取り・§5-2）。
function _boothLoadSubjectIndex(ss) {
  const appRows = _getSheet(ss, SHEET_APPLICATIONS).getDataRange().getValues();
  const appsById = {}, appsByToken = {};
  for (let i = 1; i < appRows.length; i++) {
    const appId = String(appRows[i][0] || '').trim();
    if (!appId) continue;
    const rec = {
      app_id: appId,
      ticket_token: String(appRows[i][16] || '').trim(),
      status: String(appRows[i][17] || '').trim(),
      line_friend_id: String(appRows[i][14] || '').trim()
    };
    appsById[appId] = rec;
    // 🔴 予備名札の割当で applications.ticket_token は上書きされる（§2-7）。
    // だからここは「いまの値」の索引でしかない。集計の正キーは app_id。
    if (rec.ticket_token) appsByToken[rec.ticket_token] = rec;
  }

  const spRows = _ensureSpareBadgesSheet(ss).getDataRange().getValues();
  const sparesByNo = {}, sparesByToken = {};
  for (let i = 1; i < spRows.length; i++) {
    const no = String(spRows[i][0] || '').trim();
    if (!no) continue;
    const rec = {
      spare_no: no,
      ticket_token: String(spRows[i][1] || '').trim(),
      assigned_app_id: String(spRows[i][2] || '').trim()
    };
    sparesByNo[no] = rec;
    if (rec.ticket_token) sparesByToken[rec.ticket_token] = rec;
  }

  return { appsById: appsById, appsByToken: appsByToken,
           sparesByNo: sparesByNo, sparesByToken: sparesByToken };
}

// 'F2026-123' → 'F2026-0123' / 'P-7' → 'P-07' のゆらぎを吸収する。
// 🔴 ハッシュは生値（正規化前の ec）で取るので、ここでの補正は冪等性に影響しない。
function _boothPadCode(ec) {
  let m = ec.match(/^(F\d{4}-)(\d{1,4})$/);
  if (m) return m[1] + _boothPad(m[2], 4);
  m = ec.match(/^(P-)(\d{1,2})$/);
  if (m) return m[1] + _boothPad(m[2], 2);
  return ec;
}

function _boothResolveByToken(tt, idx) {
  const app = idx.appsByToken[tt];
  if (app) return { status: 'ok', app_id: app.app_id, flag: '' };

  const spare = idx.sparesByToken[tt];
  if (spare) {
    if (!spare.assigned_app_id) return { status: 'unresolved', app_id: '', flag: 'spare_unassigned' };
    if (!idx.appsById[spare.assigned_app_id]) {
      // 割当先の申込行が見つからない（データ不整合）。推測せず未解決のままにする。
      return { status: 'unresolved', app_id: '', flag: 'unknown_code' };
    }
    return { status: 'ok', app_id: spare.assigned_app_id, flag: '' };
  }
  return { status: 'unresolved', app_id: '', flag: 'unknown_token' };
}

function _boothResolveByCode(ec, idx) {
  const code = _boothPadCode(ec);

  if (/^F\d{4}-/.test(code)) {
    const app = idx.appsById[code];
    if (app) return { status: 'ok', app_id: app.app_id, flag: '' };
    return { status: 'unresolved', app_id: '', flag: 'unknown_code' };
  }

  if (code.indexOf('P-') === 0) {
    const spare = idx.sparesByNo[code];
    if (!spare) return { status: 'unresolved', app_id: '', flag: 'unknown_code' };
    if (!spare.assigned_app_id) return { status: 'unresolved', app_id: '', flag: 'spare_unassigned' };
    if (!idx.appsById[spare.assigned_app_id]) return { status: 'unresolved', app_id: '', flag: 'unknown_code' };
    return { status: 'ok', app_id: spare.assigned_app_id, flag: '' };
  }

  return { status: 'unresolved', app_id: '', flag: 'unknown_code' };
}

// tt/ec から本人を解決する（§5-4）。🔴 推測して片方を採らない。
// 返り値: { resolve_status, app_id, ticket_token, flags[] }
function _boothResolveSubject(tt, ec, idx) {
  const flags = [];
  const a = tt ? _boothResolveByToken(tt, idx) : null;
  const b = ec ? _boothResolveByCode(ec, idx) : null;

  let picked = null;
  if (a && b) {
    if (a.status === 'ok' && b.status === 'ok' && a.app_id === b.app_id) {
      picked = a;
    } else if (a.status === 'ok' || b.status === 'ok') {
      // 片方だけ解決 / 別人に解決 → 採らない（§5-4）
      flags.push('subject_mismatch');
    } else {
      if (a.flag) flags.push(a.flag);
      if (b.flag) flags.push(b.flag);
    }
  } else {
    picked = (a || b);
    if (picked && picked.status !== 'ok') {
      if (picked.flag) flags.push(picked.flag);
      picked = null;
    }
  }

  if (!picked || picked.status !== 'ok') {
    return { resolve_status: 'unresolved', app_id: '', ticket_token: '', flags: flags };
  }

  const app = idx.appsById[picked.app_id];
  // 🔴 キャンセル済みでも受理する。控えの母集団から外すためのフラグを立てるだけ（§2-5）
  if (app && app.status === 'cancelled') flags.push('subject_cancelled');

  return {
    resolve_status: 'ok',
    app_id: picked.app_id,
    ticket_token: app ? app.ticket_token : '',
    flags: flags
  };
}

// ------------------------------------------------------------
// 明細の書き込み（item_id が決定的なので何度呼んでも重複しない・§5-3）
// ------------------------------------------------------------
function _boothSyncItems(sheets, orderId, boothId, snapItems, delivery) {
  const sh = sheets.items;
  const lastRow = sh.getLastRow();
  const prefix = orderId + '-';
  const existing = {};
  let count = 0;

  if (lastRow >= 2) {
    const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      const id = String(ids[i][0] || '');
      if (id.indexOf(prefix) === 0) { existing[id] = true; count++; }
    }
  }

  const now = _now();
  const rows = [];
  for (let i = 0; i < snapItems.length; i++) {
    const itemId = orderId + '-' + (i + 1);
    if (existing[itemId]) continue;
    const p = snapItems[i];
    rows.push([itemId, orderId, boothId, p.product_id, p.product_name, p.spec,
               p.qty, p.unit_price, delivery, now]);
  }
  if (rows.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, BOOTH_ITEM_COLS).setValues(rows);
    count += rows.length;
  }
  return count;
}

// ------------------------------------------------------------
// boothInit（GET・認証なし・booth_token で認可・§4-1）
// ------------------------------------------------------------
function boothInit(data) {
  _checkProps();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = _ensureBoothSheets(ss);

  const auth = _boothAuth(sheets, (data || {}).b);
  if (!auth.ok) return _err(auth.error);
  // 🔴 新規セッションだけを止める。キュー済みの送信（boothSubmit）は受理し続ける（§4-6）
  if (!auth.is_active) return _err('BOOTH_INACTIVE');

  const rows = sheets.products.getDataRange().getValues();
  const products = [];
  for (let i = 1; i < rows.length; i++) {
    const pid = String(rows[i][0]).trim();
    if (!pid) continue;
    if (String(rows[i][1]).trim() !== auth.booth_id) continue; // 🔴 他ブースは1バイトも返さない
    if (!_boothIsActive(rows[i][6])) continue;
    products.push({
      product_id: pid,
      product_name: String(rows[i][2] || ''),
      spec: String(rows[i][3] || ''),
      unit_price: (rows[i][4] === '' || rows[i][4] == null) ? '' : Number(rows[i][4]),
      sort_order: Number(rows[i][5] || 0)
    });
  }
  products.sort(function (a, b) {
    if (a.sort_order !== b.sort_order) return a.sort_order - b.sort_order;
    return a.product_id < b.product_id ? -1 : (a.product_id > b.product_id ? 1 : 0);
  });

  // master_version はサーバー時刻の文字列。🔴 サーバーは判定に一切使わない（監査専用・§2-6）
  const now = _now();
  return _ok({
    booth_id: auth.booth_id,
    maker_name: auth.maker_name,
    master_version: now,
    server_time: now,
    server_version: VERSION,
    products: products
  });
}

// ------------------------------------------------------------
// boothSubmit（POST・認証なし・booth_token で認可・§4-2/§5-3）
// ------------------------------------------------------------
function _boothParseSubmit(data) {
  const d = data || {};
  // 文字数での近似（GASにバイト長の安価な計測がないため）。上限を超える要求は恒久失敗にする
  if (JSON.stringify(d).length > BOOTH_MAX_REQUEST_CHARS) return { error: 'INVALID_REQUEST' };

  const boothToken = String(d.b || '').trim();
  if (!boothToken || boothToken.length > BOOTH_MAX_KEY_LEN) return { error: 'BOOTH_NOT_FOUND' };

  const idemKey = String(d.idempotency_key || '').trim();
  if (!idemKey || idemKey.length > BOOTH_MAX_KEY_LEN) return { error: 'INVALID_REQUEST' };

  const tt = _boothNormKey(d.ticket_token);
  const ec = _boothNormKey(d.entry_code).toUpperCase();
  if (tt.length > BOOTH_MAX_KEY_LEN || ec.length > BOOTH_MAX_KEY_LEN) return { error: 'INVALID_REQUEST' };
  if (!tt && !ec) return { error: 'INVALID_REQUEST' }; // 本人キーなしは記録として成立しない（§5-4）

  const delivery = String(d.delivery || '').trim();
  if (delivery !== 'handed' && delivery !== 'later') return { error: 'INVALID_REQUEST' };

  return {
    boothToken: boothToken,
    idemKey: idemKey,
    tt: tt,
    ec: ec,
    delivery: delivery,
    items: d.items,
    inputMethod: tt ? 'qr' : 'manual',
    clientCreatedAt: String(d.client_created_at || '').trim().slice(0, 32),
    clientInstanceId: String(d.client_instance_id || '').trim().slice(0, BOOTH_MAX_KEY_LEN),
    masterVersion: String(d.master_version || '').trim().slice(0, 32)
  };
}

function boothSubmit(data) {
  _checkProps();

  // ===== ロックの外（重い読み取りは全部ここ・§5-2）=====
  const req = _boothParseSubmit(data);
  if (req.error) return _err(req.error);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = _ensureBoothSheets(ss);

  const auth = _boothAuth(sheets, req.boothToken);
  if (!auth.ok) return _err(auth.error);

  const flags = [];
  // 🔴 boothSubmit は BOOTH_INACTIVE を返さない。無効化後に届いたキューも受理する（§2-6/§4-6）
  if (!auth.is_active) flags.push('stale_master');

  const canon = _boothCanonicalizeItems(req.items);
  if (canon.error) return _err(canon.error);

  const snap = _boothSnapshotProducts(sheets, auth.booth_id, canon.items);
  if (snap.error) return _err(snap.error);
  if (snap.stale) flags.push('stale_master');

  const hash = _boothPayloadHash(auth.booth_id, req.tt, req.ec, req.delivery, canon.items);

  const idx  = _boothLoadSubjectIndex(ss);
  const subj = _boothResolveSubject(req.tt, req.ec, idx);
  subj.flags.forEach(function (f) { flags.push(f); });

  const receivedAt = _now();
  if (_boothClockSkew(req.clientCreatedAt, receivedAt)) flags.push('clock_skew');
  const validationState = _boothFlags(flags);
  const expected = canon.items.length;

  // ===== ロックの中（シート書き込みだけ・狙いは0.5秒以内）=====
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(BOOTH_LOCK_WAIT_MS)) return _err('LOCK_BUSY');

  let result, rowNum, notifyOrderId;
  try {
    const or = sheets.orders;
    const lastRow = or.getLastRow();
    const keys = (lastRow >= 2) ? or.getRange(2, 1, lastRow - 1, BOOTH_KEY_COLS).getValues() : [];

    // 🔴 採番は「いま読んだA列の最大連番+1」。getLastRow() は使わない（手で1行消されただけで重複する）
    let maxSeq = 0, found = -1;
    for (let i = 0; i < keys.length; i++) {
      const m = String(keys[i][0] || '').match(/^O-(\d+)$/);
      if (m) {
        const n = parseInt(m[1], 10);
        if (n > maxSeq) maxSeq = n;
      }
      if (found < 0 &&
          String(keys[i][1]).trim() === auth.booth_id &&
          String(keys[i][2]).trim() === req.idemKey) found = i;
    }

    if (found < 0) {
      // ---- 新規 ----
      const orderId = 'O-' + _boothPad(maxSeq + 1, 4);
      rowNum = lastRow + 1;

      const values = [];
      for (let i = 0; i < BOOTH_ORDER_COLS; i++) values.push('');
      values[BO.order_id - 1]            = orderId;
      values[BO.booth_id - 1]            = auth.booth_id;
      values[BO.idempotency_key - 1]     = req.idemKey;
      values[BO.payload_hash - 1]        = hash;
      values[BO.write_state - 1]         = 'PENDING';
      values[BO.status - 1]              = 'active';
      values[BO.expected_item_count - 1] = expected;
      values[BO.resolve_status - 1]      = subj.resolve_status;
      values[BO.app_id - 1]              = subj.app_id;       // 🔴 ok のときだけ値が入る
      values[BO.ticket_token - 1]        = subj.ticket_token;
      values[BO.ticket_token_raw - 1]    = req.tt;
      values[BO.entry_code_raw - 1]      = req.ec;
      values[BO.input_method - 1]        = req.inputMethod;
      values[BO.client_created_at - 1]   = req.clientCreatedAt;
      values[BO.received_at - 1]         = receivedAt;
      values[BO.created_at - 1]          = receivedAt;
      values[BO.client_instance_id - 1]  = req.clientInstanceId;
      values[BO.master_version - 1]      = req.masterVersion;
      values[BO.validation_state - 1]    = validationState;

      // 🔴 PENDING行の追加は1回のsetValuesで全25列を書く。列を分けて書くと
      //    「expected_item_count だけ空のPENDING行」という復旧不能な中間状態ができる（§5-2）
      or.getRange(rowNum, 1, 1, BOOTH_ORDER_COLS).setValues([values]);

      const cnt = _boothSyncItems(sheets, orderId, auth.booth_id, snap.items, req.delivery);
      let writeState = 'PENDING';
      if (cnt === expected) {
        or.getRange(rowNum, BO.write_state).setValue('COMPLETE');
        writeState = 'COMPLETE';
      }

      notifyOrderId = orderId;
      result = _ok({
        order_id: orderId, write_state: writeState, status: 'active',
        resolve_status: subj.resolve_status, validation_state: validationState, duplicate: false
      });

    } else {
      rowNum = 2 + found;
      const row = keys[found];
      const orderId = String(row[0]);
      const storedHash = String(row[3] || '');

      // 🔴 墓標の判定を CONFLICT より先に置く（§5-3）。墓標は payload_hash が空という1点で見分ける
      if (storedHash === '') {
        // ---- 墓標行を埋める（取消が先に届いていた・§2-2）----
        // 🔴 payload_hash（D列）は最後に書く。途中で落ちても「payload_hash が空＝まだ墓標」の
        //    ままなので、再送が同じ経路をもう一度やり直せる（順序を変えると復旧不能になる）
        or.getRange(rowNum, BO.expected_item_count, 1, 9).setValues([[
          expected, subj.resolve_status, subj.app_id, subj.ticket_token,
          req.tt, req.ec, req.inputMethod, req.clientCreatedAt, receivedAt
        ]]); // G〜O
        or.getRange(rowNum, BO.client_instance_id, 1, 3).setValues([[
          req.clientInstanceId, req.masterVersion, validationState
        ]]); // Q〜S
        or.getRange(rowNum, BO.payload_hash).setValue(hash);

        const cnt = _boothSyncItems(sheets, orderId, auth.booth_id, snap.items, req.delivery);
        let writeState = 'PENDING';
        if (cnt === expected) {
          or.getRange(rowNum, BO.write_state).setValue('COMPLETE');
          writeState = 'COMPLETE';
        }

        // 🔴 status は voided のまま（絶対に active に戻さない）。通知もしない
        result = _ok({
          order_id: orderId, write_state: writeState, status: 'voided',
          resolve_status: subj.resolve_status, validation_state: validationState, duplicate: false
        });

      } else if (storedHash !== hash) {
        result = _err('IDEMPOTENCY_CONFLICT');

      } else {
        // ---- 同一キー・同一内容（＝応答喪失後の再送）----
        // 🔴 ロック内で applications を読み直さない。解決はロック外で済んでいる（§5-3）
        let writeState    = String(row[4] || '');
        const rowStatus   = String(row[5] || '');
        let resolveStatus = String(row[7] || '');
        let vState = String(or.getRange(rowNum, BO.validation_state).getValue() || '');

        if (resolveStatus !== 'ok' && subj.resolve_status === 'ok') {
          or.getRange(rowNum, BO.resolve_status, 1, 3)
            .setValues([[ 'ok', subj.app_id, subj.ticket_token ]]); // H〜J
          or.getRange(rowNum, BO.validation_state).setValue(validationState);
          resolveStatus = 'ok';
          vState = validationState;
        }

        if (writeState === 'PENDING') {
          // hashが一致している＝明細も同一なので、期待件数は今回の正規化結果と必ず一致する
          const cnt = _boothSyncItems(sheets, orderId, auth.booth_id, snap.items, req.delivery);
          if (cnt === expected) {
            or.getRange(rowNum, BO.write_state).setValue('COMPLETE');
            writeState = 'COMPLETE';
          }
        }
        or.getRange(rowNum, BO.received_at).setValue(receivedAt);

        if (writeState === 'COMPLETE' && rowStatus === 'active') notifyOrderId = orderId;

        // 🔴 duplicate:true でも「いまの」resolve_status / validation_state を返す。
        //    端末の⚠表示はこの2つだけを見て更新する
        result = _ok({
          order_id: orderId, write_state: writeState, status: rowStatus,
          resolve_status: resolveStatus, validation_state: vState, duplicate: true
        });
      }
    }
  } catch (err) {
    Logger.log('boothSubmit error: ' + err);
    return _err('INTERNAL_ERROR');
  } finally {
    lock.releaseLock();
  }

  // ===== ロックの外（通知の失敗は注文の失敗にしない・§5-5）=====
  // 🔴 subj.resolve_status も見る。再送で「行はok・今回の解決はunresolved」のときに
  //    appsById[''] を引いて not_applicable を書き込んでしまうのを防ぐ
  if (notifyOrderId && result && result.success && result.data.write_state === 'COMPLETE' &&
      result.data.status === 'active' && result.data.resolve_status === 'ok' &&
      subj.resolve_status === 'ok') {
    try {
      _boothNotifyLine(sheets, rowNum, idx.appsById[subj.app_id], auth.maker_name,
                       snap.items, result.data.validation_state);
    } catch (e) {
      Logger.log('boothSubmit LINE通知の失敗（注文は成立）: ' + e);
    }
  }
  return result;
}

// 友だちにだけ即時プッシュする（§5-5・正本§15-6）。
// 🔴 line_notify_state が 'sent' の行には二度と送らない（冪等再送での重複防止）
function _boothNotifyLine(sheets, rowNum, appRec, makerName, snapItems, validationState) {
  const or = sheets.orders;
  const current = String(or.getRange(rowNum, BO.line_notify_state).getValue() || '');
  if (current === 'sent') return;

  const friendId = appRec ? appRec.line_friend_id : '';
  if (!friendId || String(validationState).indexOf('subject_cancelled') >= 0) {
    if (current !== 'not_applicable') or.getRange(rowNum, BO.line_notify_state).setValue('not_applicable');
    return;
  }

  const lines = snapItems.map(function (p) {
    return '・' + p.product_name + (p.spec ? '（' + p.spec + '）' : '') + ' × ' + p.qty;
  }).join('\n');
  const text = 'ビューフェス2026\n' + makerName + 'ブースでのご注文を承りました。\n\n' +
               lines + '\n\n※控えは当日終了後にメールでお送りします。';

  try {
    _lhSendMessage(friendId, text);
    or.getRange(rowNum, BO.line_notify_state, 1, 2).setValues([['sent', _now()]]);
  } catch (e) {
    or.getRange(rowNum, BO.line_notify_state).setValue('failed');
    or.getRange(rowNum, BO.last_error).setValue(String(e).substring(0, 300));
  }
}

// ------------------------------------------------------------
// boothVoid（POST・認証なし・§4-3/§2-2）
// ------------------------------------------------------------
function boothVoid(data) {
  _checkProps();
  const d = data || {};

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = _ensureBoothSheets(ss);

  const auth = _boothAuth(sheets, d.b);
  if (!auth.ok) return _err(auth.error);

  const voidKey   = String(d.void_idempotency_key || '').trim().slice(0, BOOTH_MAX_KEY_LEN);
  const targetKey = String(d.target_idempotency_key || '').trim().slice(0, BOOTH_MAX_KEY_LEN);
  const targetId  = String(d.target_order_id || '').trim().slice(0, BOOTH_MAX_KEY_LEN);
  if (!targetKey && !targetId) return _err('INVALID_REQUEST');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(BOOTH_LOCK_WAIT_MS)) return _err('LOCK_BUSY');

  try {
    const or = sheets.orders;
    const lastRow = or.getLastRow();
    const keys = (lastRow >= 2) ? or.getRange(2, 1, lastRow - 1, BOOTH_KEY_COLS).getValues() : [];

    let maxSeq = 0, found = -1;
    for (let i = 0; i < keys.length; i++) {
      const m = String(keys[i][0] || '').match(/^O-(\d+)$/);
      if (m) {
        const n = parseInt(m[1], 10);
        if (n > maxSeq) maxSeq = n;
      }
      // 🔴 他ブースの注文は取り消せない（画面制御に頼らずここで照合する）
      if (found >= 0 || String(keys[i][1]).trim() !== auth.booth_id) continue;
      if (targetKey && String(keys[i][2]).trim() === targetKey) found = i;
      else if (!targetKey && targetId && String(keys[i][0]).trim() === targetId) found = i;
    }

    const now = _now();

    if (found >= 0) {
      const rowNum  = 2 + found;
      const orderId = String(keys[found][0]);
      if (String(keys[found][5]) === 'voided') {
        return _ok({ order_id: orderId, status: 'voided', tombstone: false }); // 冪等
      }
      or.getRange(rowNum, BO.status).setValue('voided');
      or.getRange(rowNum, BO.void_idempotency_key, 1, 2).setValues([[voidKey, now]]); // V,W
      return _ok({ order_id: orderId, status: 'voided', tombstone: false });
    }

    if (!targetKey) return _err('ORDER_NOT_FOUND'); // order_id 指定のみで見つからない場合

    // ---- 墓標行を作る（注文がまだ届いていない・§2-2）----
    const orderId = 'O-' + _boothPad(maxSeq + 1, 4);
    const values = [];
    for (let i = 0; i < BOOTH_ORDER_COLS; i++) values.push('');
    values[BO.order_id - 1]             = orderId;
    values[BO.booth_id - 1]             = auth.booth_id;
    values[BO.idempotency_key - 1]      = targetKey; // 後から届く注文がこの行に当たる鍵
    values[BO.payload_hash - 1]         = '';        // 🔴 空＝墓標の目印
    values[BO.write_state - 1]          = 'PENDING';
    values[BO.status - 1]               = 'voided';
    values[BO.received_at - 1]          = now;
    values[BO.created_at - 1]           = now;
    values[BO.void_idempotency_key - 1] = voidKey;
    values[BO.voided_at - 1]            = now;
    values[BO.note - 1]                 = 'tombstone: 取消が注文より先に到着';
    or.getRange(lastRow + 1, 1, 1, BOOTH_ORDER_COLS).setValues([values]);

    return _ok({ order_id: orderId, status: 'voided', tombstone: true });

  } catch (err) {
    Logger.log('boothVoid error: ' + err);
    return _err('INTERNAL_ERROR');
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------
// boothRecent（POST・認証なし・§4-4）
// 🔴 氏名・サロン名は返さない。entry_label は「解決済みなら app_id、未解決なら生の入力値」
// ------------------------------------------------------------
function boothRecent(data) {
  _checkProps();
  const d = data || {};
  const limit = Math.min(Math.max(parseInt(d.limit, 10) || 10, 1), 50);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = _ensureBoothSheets(ss);

  const auth = _boothAuth(sheets, d.b);
  if (!auth.ok) return _err(auth.error);

  const rows = sheets.orders.getDataRange().getValues();
  const picked = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[BO.booth_id - 1]).trim() !== auth.booth_id) continue;
    if (String(r[BO.write_state - 1]) !== 'COMPLETE') continue;
    if (String(r[BO.status - 1]) !== 'active') continue;
    picked.push({
      order_id: String(r[BO.order_id - 1]),
      idempotency_key: String(r[BO.idempotency_key - 1]),
      entry_label: String(r[BO.resolve_status - 1]) === 'ok'
        ? String(r[BO.app_id - 1])
        : (String(r[BO.entry_code_raw - 1] || '') || String(r[BO.ticket_token_raw - 1] || '')),
      client_created_at: String(r[BO.client_created_at - 1] || ''),
      items: []
    });
  }

  picked.sort(function (a, b) { return a.order_id < b.order_id ? 1 : (a.order_id > b.order_id ? -1 : 0); });
  const out = picked.slice(0, limit);

  const wanted = {};
  out.forEach(function (o) { wanted[o.order_id] = o; });

  const itemRows = sheets.items.getDataRange().getValues();
  for (let i = 1; i < itemRows.length; i++) {
    const o = wanted[String(itemRows[i][1])];
    if (!o) continue;
    o.items.push({ product_name: String(itemRows[i][4] || ''), qty: Number(itemRows[i][6] || 0) });
  }

  return _ok({ booth_id: auth.booth_id, orders: out, server_time: _now() });
}

// ------------------------------------------------------------
// 社員用（認証必須・§4-5）
// ------------------------------------------------------------

// 撤収照合用。ブース別に COMPLETE&active / PENDING / unresolved の件数を返す（§15-5-4）
function boothSummary(data) {
  const authz = _requireSession(data);
  if (!authz.ok) return _err(authz.error);
  _checkProps();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = _ensureBoothSheets(ss);

  const boothRows = sheets.booths.getDataRange().getValues();
  const summary = {}, order = [];
  for (let i = 1; i < boothRows.length; i++) {
    const bid = String(boothRows[i][0]).trim();
    if (!bid) continue;
    summary[bid] = {
      booth_id: bid, maker_name: String(boothRows[i][1] || ''),
      is_active: _boothIsActive(boothRows[i][3]),
      complete_active: 0, pending: 0, unresolved: 0, voided: 0
    };
    order.push(bid);
  }

  const lastRow = sheets.orders.getLastRow();
  const rows = (lastRow >= 2) ? sheets.orders.getRange(2, 1, lastRow - 1, BOOTH_KEY_COLS).getValues() : [];
  for (let i = 0; i < rows.length; i++) {
    const bid = String(rows[i][1]).trim();
    if (!bid) continue;
    if (!summary[bid]) {
      summary[bid] = { booth_id: bid, maker_name: '(booths未登録)', is_active: false,
                       complete_active: 0, pending: 0, unresolved: 0, voided: 0 };
      order.push(bid);
    }
    const s = summary[bid];
    const writeState = String(rows[i][4]);
    const status     = String(rows[i][5]);
    if (status === 'voided') s.voided++;
    if (writeState === 'COMPLETE' && status === 'active') s.complete_active++;
    if (writeState === 'PENDING') s.pending++;
    if (String(rows[i][7]) === 'unresolved') s.unresolved++;
  }

  const list = order.map(function (bid) { return summary[bid]; });
  const total = { complete_active: 0, pending: 0, unresolved: 0, voided: 0 };
  list.forEach(function (s) {
    total.complete_active += s.complete_active;
    total.pending += s.pending;
    total.unresolved += s.unresolved;
    total.voided += s.voided;
  });

  return _ok({ booths: list, total: total, server_time: _now(), server_version: VERSION });
}

// unresolved 行の一括再解決（§4-5）。🔴 mail_merge 生成の直前に必ず実行する（§7ルール6）
// ロックは取らない: 書き込むのは該当行の H〜J・S だけで、boothSubmit の再送が同時に書いても
// 同じ値になる（解決は生の入力キーから決まる）。当日終了後に社員が回す想定。
function boothResolveUnresolved(data) {
  const authz = _requireSession(data);
  if (!authz.ok) return _err(authz.error);
  _checkProps();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = _ensureBoothSheets(ss);
  const or = sheets.orders;

  const lastRow = or.getLastRow();
  if (lastRow < 2) return _ok({ checked: 0, resolved: 0, still_unresolved: 0, details: [] });

  const rows = or.getRange(2, 1, lastRow - 1, BO.validation_state).getValues();
  const idx = _boothLoadSubjectIndex(ss);

  let checked = 0, resolved = 0, still = 0;
  const details = [];
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][BO.resolve_status - 1]) !== 'unresolved') continue;
    checked++;

    const tt = String(rows[i][BO.ticket_token_raw - 1] || '');
    const ec = String(rows[i][BO.entry_code_raw - 1] || '');
    const subj = _boothResolveSubject(tt, ec, idx);
    const orderId = String(rows[i][BO.order_id - 1]);

    if (subj.resolve_status !== 'ok') {
      still++;
      details.push({ order_id: orderId, resolve_status: 'unresolved',
                     validation_state: _boothFlags(subj.flags) });
      continue;
    }

    const rowNum = 2 + i;
    or.getRange(rowNum, BO.resolve_status, 1, 3).setValues([['ok', subj.app_id, subj.ticket_token]]);

    // 既存フラグから未解決系だけを落とし、今回の判定結果を足す
    const keep = String(rows[i][BO.validation_state - 1] || '').split(',').filter(function (f) {
      return f && ['spare_unassigned', 'unknown_token', 'unknown_code', 'subject_mismatch'].indexOf(f) < 0;
    });
    const merged = _boothFlags(keep.concat(subj.flags));
    or.getRange(rowNum, BO.validation_state).setValue(merged);

    resolved++;
    details.push({ order_id: orderId, resolve_status: 'ok', app_id: subj.app_id,
                   validation_state: merged });
  }

  return _ok({ checked: checked, resolved: resolved, still_unresolved: still, details: details });
}

// メーカー別・商品別の集計CSV（§4-5）。母集団は COMPLETE かつ active。
// 🔴 subject_cancelled は**含める**（実際に発生した注文なので・§2-5）
function boothExportCsv(data) {
  const authz = _requireSession(data);
  if (!authz.ok) return _err(authz.error);
  _checkProps();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = _ensureBoothSheets(ss);
  const filterBooth = String((data || {}).booth_id || '').trim();

  const makerName = {};
  const boothRows = sheets.booths.getDataRange().getValues();
  for (let i = 1; i < boothRows.length; i++) {
    makerName[String(boothRows[i][0]).trim()] = String(boothRows[i][1] || '');
  }

  const lastRow = sheets.orders.getLastRow();
  const orderRows = (lastRow >= 2) ? sheets.orders.getRange(2, 1, lastRow - 1, BOOTH_KEY_COLS).getValues() : [];
  const target = {};
  let orderCount = 0;
  for (let i = 0; i < orderRows.length; i++) {
    if (String(orderRows[i][4]) !== 'COMPLETE' || String(orderRows[i][5]) !== 'active') continue;
    const bid = String(orderRows[i][1]).trim();
    if (filterBooth && bid !== filterBooth) continue;
    target[String(orderRows[i][0])] = true;
    orderCount++;
  }

  const agg = {}, keys = [];
  const itemRows = sheets.items.getDataRange().getValues();
  for (let i = 1; i < itemRows.length; i++) {
    const orderId = String(itemRows[i][1]);
    if (!target[orderId]) continue;
    const bid = String(itemRows[i][2]).trim();
    const pid = String(itemRows[i][3]).trim();
    const key = bid + ' ' + pid;
    if (!agg[key]) {
      agg[key] = {
        booth_id: bid, maker_name: makerName[bid] || '', product_id: pid,
        product_name: String(itemRows[i][4] || ''), spec: String(itemRows[i][5] || ''),
        unit_price: (itemRows[i][7] === '' || itemRows[i][7] == null) ? '' : Number(itemRows[i][7]),
        qty: 0, orders: 0, amount: ''
      };
      keys.push(key);
    }
    agg[key].qty += Number(itemRows[i][6] || 0);
    agg[key].orders++;
  }
  keys.sort();

  const header = ['booth_id', 'maker_name', 'product_id', 'product_name', 'spec',
                  'unit_price', 'qty_total', 'order_count', 'amount_total'];
  const lines = [header.map(_boothCsvCell).join(',')];
  const rows = [];
  keys.forEach(function (k) {
    const a = agg[k];
    a.amount = (a.unit_price === '') ? '' : a.unit_price * a.qty;
    rows.push(a);
    lines.push([a.booth_id, a.maker_name, a.product_id, a.product_name, a.spec,
                a.unit_price, a.qty, a.orders, a.amount].map(_boothCsvCell).join(','));
  });

  return _ok({
    generated_at: _now(), order_count: orderCount, row_count: rows.length,
    rows: rows, csv: lines.join('\r\n')
  });
}

function _boothCsvCell(v) {
  const s = String(v == null ? '' : v);
  return '"' + s.split('"').join('""') + '"';
}

// 退避したキューJSONをシートへ戻す（§4-5）。
// 端末の書き出しJSONは { idempotency_key, type:'submit'|'void', payload } の配列。
// 🔴 payload をそのまま boothSubmit / boothVoid へ流すので、投入は冪等
//   （同じものを端末が後から送っても二重にならない）
function boothImportQueue(data) {
  const authz = _requireSession(data);
  if (!authz.ok) return _err(authz.error);
  _checkProps();

  const d = data || {};
  const entries = Array.isArray(d.entries) ? d.entries : null;
  if (!entries) return _err('INVALID_REQUEST');
  if (entries.length > 200) return _err('INVALID_REQUEST'); // 1回の投入は200件まで

  const results = [];
  let okCount = 0, ngCount = 0;
  for (let i = 0; i < entries.length; i++) {
    const e = entries[i] || {};
    const type = String(e.type || 'submit');
    const payload = e.payload || {};
    let res;
    try {
      res = (type === 'void') ? boothVoid(payload) : boothSubmit(payload);
    } catch (err) {
      res = _err('INTERNAL_ERROR');
      Logger.log('boothImportQueue error: ' + err);
    }
    if (res && res.success) okCount++; else ngCount++;
    results.push({
      idempotency_key: String(e.idempotency_key || ''),
      type: type,
      success: !!(res && res.success),
      error: (res && res.success) ? '' : String((res || {}).error || ''),
      data: (res && res.success) ? res.data : null
    });
  }

  return _ok({ total: entries.length, ok: okCount, ng: ngCount, results: results });
}

// ============================================================
// 🆕 v0.26.0（2026-09-05）受付 scan.html 用（`当日運用_堅牢化設計.md` §3）
//
// 設計の要点（この3本を実装する理由）:
//   R1 表示に必要なものを通信から切り離す → 朝1回 scanRoster で名簿を端末に落とす
//   R2 書き込みはローカルキューに積み、裏で送る → scanCheckin は端末のキューから届く
//   R3 同期通信は待たされても誰も困らない場面にだけ置く → scanCheckedIn は30秒ごとの裏更新
//
// 🔒 3本とも認証必須（_requireSession）。名簿は氏名・サロン名を含むため公開しない。
// 🔴 端末に渡すのは受付に必要な項目だけ（§3-3）。メール・電話・住所・紹介者は返さない。
//    端末紛失時の被害を減らすための設計であり、「ついでに返す」をしないこと。
// 🔴 本人解決は booth と同じ _boothResolveByToken / _boothResolveByCode を使う。
//    受付とブースで解決規則がずれると、同じ名札が画面ごとに違う人になる。
// ============================================================

const SCAN_LOCK_WAIT_MS = 25000; // boothと同じ待ち方（実測: 30件同時でもLOCK_BUSYゼロ）
const SCAN_MAX_BATCH    = 50;    // 1リクエストで受け取るチェックインの上限

// checkins シートの列（1始まり・setupSheets が作るヘッダーと1対1）
const CK = {
  checkin_id: 1, app_id: 2, ticket_token: 3, session_id: 4,
  checked_at: 5, staff_user_id: 6, device: 7, note: 8
};
const CHECKIN_COLS = 8;

// checkins シートが無ければヘッダー付きで作成する（_ensureSpareBadgesSheet と同じ形）。
// 🔴 setupSheets() と同じヘッダーにすること。片方だけ直すと列がずれる。
function _ensureCheckinsSheet(ss) {
  let sh = ss.getSheetByName(SHEET_CHECKINS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_CHECKINS);
    sh.getRange(1, 1, 1, CHECKIN_COLS).setValues([[
      'checkin_id', 'app_id', 'ticket_token', 'session_id',
      'checked_at', 'staff_user_id', 'device', 'note'
    ]]);
    sh.setFrozenRows(1);
    Logger.log('checkinsシート作成完了');
  }
  return sh;
}

// ------------------------------------------------------------
// 名簿取得（doPost: action=scanRoster）。当日朝に各端末で1回だけ叩く。
// data: { session_token }
// ------------------------------------------------------------
function scanRoster(data) {
  const auth = _requireSession(data);
  if (!auth.ok) return _err(auth.error);
  _checkProps();

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rows = _getSheet(ss, SHEET_APPLICATIONS).getDataRange().getValues();
  const cfg  = _getConfig();

  let skippedNoToken = 0;
  const roster = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][17]) !== 'confirmed') continue; // status（listBadges と同じ絞り込み）

    const token = String(rows[i][16] || '').trim();
    if (!token) { skippedNoToken++; continue; } // ticket_tokenが空＝データ異常。QRで来られない

    const businessType = String(rows[i][20] || '');
    roster.push({
      app_id:        String(rows[i][0]),
      ticket_token:  token,
      staff_name:    String(rows[i][5]),
      salon_name:    String(rows[i][4]),
      business_type: businessType,
      band:          _resolveBand(cfg, businessType)
    });
    // 🔴 メール・電話・住所・紹介者は入れない（§3-3）。
    //    追加要望が出たら「端末紛失時の被害を減らす」という理由ごと再確認すること
  }
  roster.sort(function (a, b) { return a.app_id < b.app_id ? -1 : (a.app_id > b.app_id ? 1 : 0); });

  // 予備名札は「割当済み」のものだけ渡す。未割当は誰でもないので端末では解決できない
  // （その場で割り当てた予備名札は朝の名簿には載らない。受付番号の手入力で通す運用・§3-8）
  const spRows = _ensureSpareBadgesSheet(ss).getDataRange().getValues();
  const spares = [];
  for (let i = 1; i < spRows.length; i++) {
    const no    = String(spRows[i][0] || '').trim();
    const token = String(spRows[i][1] || '').trim();
    const appId = String(spRows[i][2] || '').trim();
    if (!no || !appId) continue;
    spares.push({ spare_no: no, ticket_token: token, app_id: appId });
  }

  return _ok({
    fetched_at:       _now(),
    server_version:   VERSION,
    total:            roster.length,
    skipped_no_token: skippedNoToken,
    bands:            _buildBandsFromConfig(cfg),
    roster:           roster,
    spares:           spares
  });
}

// ------------------------------------------------------------
// チェックイン記録（doPost: action=scanCheckin）。端末のキューからまとめて届く。
// data: { session_token, checkins: [{ checkin_id, ticket_token, entry_code,
//                                     checked_in_at, device_id, note }] }
//
// 🔴 冪等キーは端末が生成する checkin_id。配送が失われて再送されても行は増えない
//    （GASの故障モードは「書けたのに返事が来ない」・調査レポート §3-2）。
// 🔴 同じ人が2回スキャンされたら checkin_id が違うので2行とも残す。
//    checkins は監査用の生ログであり、重複入場の判断材料そのものを消してはいけない。
//    画面へは already_checked_in で知らせる。
// ------------------------------------------------------------
function scanCheckin(data) {
  const auth = _requireSession(data);
  if (!auth.ok) return _err(auth.error);
  _checkProps();

  const d    = data || {};
  const list = Array.isArray(d.checkins) ? d.checkins : null;
  if (!list || list.length === 0) return _err('NO_CHECKINS');
  if (list.length > SCAN_MAX_BATCH) return _err('INVALID_REQUEST');

  // ===== ロックの外（重い読み取りは全部ここ）=====
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh  = _ensureCheckinsSheet(ss);
  const idx = _boothLoadSubjectIndex(ss);

  const parsed = [];
  for (let i = 0; i < list.length; i++) {
    const c  = list[i] || {};
    const id = String(c.checkin_id || '').trim();
    const tt = String(c.ticket_token || '').trim();
    const ec = String(c.entry_code || '').trim();
    if (!id) return _err('INVALID_REQUEST');
    if (!tt && !ec) return _err('INVALID_REQUEST');

    // 本人解決は booth と同じ規則。QRのtokenを優先し、無ければ受付番号で引く
    const r = tt ? _boothResolveByToken(tt, idx) : _boothResolveByCode(ec, idx);
    parsed.push({
      checkin_id:   id,
      ticket_token: tt,
      entry_code:   ec,
      checked_at:   String(c.checked_in_at || '').trim() || _now(),
      device:       String(c.device_id || '').trim(),
      note:         String(c.note || '').trim(),
      resolve:      r
    });
  }

  // ===== ロックの中（シート書き込みだけ）=====
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(SCAN_LOCK_WAIT_MS)) return _err('LOCK_BUSY');

  const results = [];
  try {
    const lastRow  = sh.getLastRow();
    const existing = (lastRow >= 2)
      ? sh.getRange(2, 1, lastRow - 1, CHECKIN_COLS).getValues() : [];

    const seenId  = {};   // checkin_id → true（既に書かれている）
    const firstBy = {};   // 人のキー → 最初のチェックイン { checkin_id, checked_at }
    for (let i = 0; i < existing.length; i++) {
      const id = String(existing[i][CK.checkin_id - 1] || '').trim();
      if (id) seenId[id] = true;
      const key = _scanSubjectKey(String(existing[i][CK.app_id - 1] || ''),
                                  String(existing[i][CK.ticket_token - 1] || ''));
      if (key && !firstBy[key]) {
        firstBy[key] = { checkin_id: id, checked_at: String(existing[i][CK.checked_at - 1] || '') };
      }
    }

    const toAppend = [];
    for (let i = 0; i < parsed.length; i++) {
      const p     = parsed[i];
      const key   = _scanSubjectKey(p.resolve.app_id, p.ticket_token || p.entry_code);
      const prior = firstBy[key] || null;

      if (seenId[p.checkin_id]) {
        // 再送。行は増やさない
        results.push({
          checkin_id: p.checkin_id, status: 'duplicate',
          app_id: p.resolve.app_id, resolve_status: p.resolve.status, flag: p.resolve.flag,
          already_checked_in: !!(prior && prior.checkin_id !== p.checkin_id),
          first_checked_at: prior ? prior.checked_at : ''
        });
        continue;
      }

      const values = [];
      for (let k = 0; k < CHECKIN_COLS; k++) values.push('');
      values[CK.checkin_id - 1]    = p.checkin_id;
      values[CK.app_id - 1]        = p.resolve.app_id;   // 🔴 ok のときだけ値が入る
      values[CK.ticket_token - 1]  = p.ticket_token || p.entry_code;
      values[CK.session_id - 1]    = '';                 // 受付では使わない（セミナー出欠用の列）
      values[CK.checked_at - 1]    = p.checked_at;       // 🔴 端末側の時刻。送信時刻ではない（§3-5）
      values[CK.staff_user_id - 1] = String(auth.session.user_id || '');
      values[CK.device - 1]        = p.device;
      values[CK.note - 1]          = [p.resolve.flag, p.note].filter(String).join(' ');
      toAppend.push(values);

      seenId[p.checkin_id] = true;
      if (key && !firstBy[key]) firstBy[key] = { checkin_id: p.checkin_id, checked_at: p.checked_at };

      results.push({
        checkin_id: p.checkin_id, status: 'recorded',
        app_id: p.resolve.app_id, resolve_status: p.resolve.status, flag: p.resolve.flag,
        already_checked_in: !!(prior && prior.checkin_id !== p.checkin_id),
        first_checked_at: prior ? prior.checked_at : ''
      });
    }

    // 🔴 まとめて1回のsetValuesで書く。1件ずつ書くとロック保持時間が件数に比例して伸びる
    if (toAppend.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, CHECKIN_COLS).setValues(toAppend);
    }
  } catch (err) {
    Logger.log('scanCheckin error: ' + err);
    return _err('INTERNAL_ERROR');
  } finally {
    lock.releaseLock();
  }

  return _ok({ received: parsed.length, results: results, server_time: _now() });
}

// 重複入場の判定キー。app_id が取れたらそれ、取れなければ生のtoken/受付番号で見る。
// 🔴 app_id を正キーにする理由: 予備名札の割当で applications.ticket_token は上書きされうる
//    （booth実装設計_確定版.md §2-7）。tokenだけで見ると同一人物が2人に割れる。
function _scanSubjectKey(appId, raw) {
  const a = String(appId || '').trim();
  if (a) return 'a:' + a;
  const r = String(raw || '').trim();
  return r ? 'r:' + r : '';
}

// ------------------------------------------------------------
// 入場済み一覧（doPost: action=scanCheckedIn）。30秒ごとに裏で取り、他端末の入場を反映する。
// data: { session_token }
// 🔴 これが取れなくても受付は止まらない。失われるのは「重複入場の検知」だけ（§3-2）。
// ------------------------------------------------------------
function scanCheckedIn(data) {
  const auth = _requireSession(data);
  if (!auth.ok) return _err(auth.error);
  _checkProps();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = _ensureCheckinsSheet(ss);
  const lastRow = sh.getLastRow();
  const rows = (lastRow >= 2) ? sh.getRange(2, 1, lastRow - 1, CHECKIN_COLS).getValues() : [];

  const seen = {};
  const entries = [];
  for (let i = 0; i < rows.length; i++) {
    const appId = String(rows[i][CK.app_id - 1] || '').trim();
    const raw   = String(rows[i][CK.ticket_token - 1] || '').trim();
    const key   = _scanSubjectKey(appId, raw);
    if (!key || seen[key]) continue;   // 最初の1件だけ返す（入場時刻は最初のスキャンが正）
    seen[key] = true;
    entries.push({
      key: key, app_id: appId, ticket_token: raw,
      checked_at: String(rows[i][CK.checked_at - 1] || ''),
      device: String(rows[i][CK.device - 1] || '')
    });
  }

  return _ok({ as_of: _now(), total: entries.length, entries: entries });
}

// ============================================================
// 🆕 v0.33.0（2026-09-06）admin.html の当日機能
//   ・adminCreateApplication : 飛び込み客の代理登録
//   ・assignSpare            : 予備名札の割当（🔴 再割当を拒否するガード付き）
//   ・adminResendPass        : 一覧の行から入場パスを再送
//
// 🔒 3本とも認証必須（_requireSession）。
// 🔴 公開フォームの経路（applyApplication / applyLiff / updateApplication）には一切触らない。
//    §15-0 の公開フォーム専用ガードは applyApplication からしか呼ばない約束なので、
//    代理登録はそれを通らない別経路として書く（受付は社員が本人を目視しているため、
//    公開フォーム向けの機械的なガードは不要であり、かけると当日の受付が止まる）。
// ============================================================

const ADMIN_LOCK_WAIT_MS = 25000;

// 代理登録で必須にする項目（2026-09-06 Takashiさん確定）。
// 🔴 メール・電話は任意。受付で聞き出すと列が詰まるため。
//    名札は予備名札（事前に刷ってある「予備 P-07」）を手渡すので、ここで集めた情報は名札に出ない。
//    この3つは台帳（誰が来たか）とブース購買記録の突合のために取る。
function _validateAdminApplication(data) {
  const f = {
    salonName:    String(data.salon_name || '').trim(),
    staffName:    String(data.staff_name || '').trim(),
    businessType: String(data.business_type || '').trim(),
    email:        String(data.email || '').trim(),
    phone:        String(data.phone || '').trim(),
    note:         String(data.note || '').trim(),
    hasTransaction: String(data.has_transaction || '').trim(),
    address:      '',
    referrer:     '',
    agree:        ''   // 🔴 本人がチェックした値ではないので空のまま。TRUEを書かない
  };
  if (!f.staffName)    return { error: 'お名前を入力してください' };
  if (!f.salonName)    return { error: 'サロン名を入力してください' };
  if (!f.businessType) return { error: '業態を選択してください' };
  if (BUSINESS_TYPE_OPTIONS.indexOf(f.businessType) < 0) return { error: '業態の値が不正です' };

  // メールは任意だが、入れるなら形式は見る（打ち間違いをそのまま台帳に残さない）
  if (f.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(f.email)) {
    return { error: 'メールアドレスの形式が正しくありません（空欄でも登録できます）' };
  }
  f.emailNorm = f.email ? f.email.toLowerCase() : '';
  return { fields: f };
}

// 飛び込み客の代理登録（doPost: action=adminCreateApplication）。
// data: { session_token, request_id, staff_name, salon_name, business_type, email?, phone?, note? }
//
// 🔴 request_id は端末が作る冪等キー。受付で二度押ししても、GASが応答を落としても
//    行は増えない（配送障害の実測: 書けたのに返事が来ない・調査レポート §3-2）。
// 🔴 控えメールは送らない（2026-09-06 確定）。目の前にいる人にパスURLを送っても使わず、
//    メール枠100通/日を当日の他の用途のために残す。
function adminCreateApplication(data) {
  const auth = _requireSession(data);
  if (!auth.ok) return _err(auth.error);
  _checkProps();

  const d = data || {};
  const requestId = String(d.request_id || '').trim();
  if (!requestId) return _err('INVALID_REQUEST');   // 冪等キーなしは受け付けない

  const v = _validateAdminApplication(d);
  if (v.error) return _err(v.error);
  const f = v.fields;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = _ensureTantouColumn(_getSheet(ss, SHEET_APPLICATIONS));

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(ADMIN_LOCK_WAIT_MS)) return _err('LOCK_BUSY');

  try {
    const rows = sh.getDataRange().getValues();

    // 冪等: 同じ request_id が既にあれば、その行をそのまま返す
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][21] || '').trim() === requestId) {   // V列 = request_id
        return _ok({
          app_id: String(rows[i][0]), ticket_token: String(rows[i][16]),
          staff_name: String(rows[i][5]), salon_name: String(rows[i][4]),
          duplicate: true
        });
      }
    }

    // 🔴 行の組み立ては公開フォームと同じ _appendApplicationRow を使う。
    //    列レイアウトを知っている場所を1つに保つため（別に書くと列がずれても気づけない）
    const created = _appendApplicationRow(sh, rows, f, 'admin', '', '', requestId);

    return _ok({
      app_id: created.appId, ticket_token: created.ticketToken,
      staff_name: f.staffName, salon_name: f.salonName,
      duplicate: false
    });
  } catch (err) {
    Logger.log('adminCreateApplication error: ' + err);
    return _err('INTERNAL_ERROR');
  } finally {
    lock.releaseLock();
  }
}

// 予備名札の割当（doPost: action=assignSpare）。
// data: { session_token, spare_no, app_id }
//
// 🔴 割当済みの予備名札を別人に付け替えることは禁止（booth実装設計_確定版.md §2-7）。
//    未解決注文の自己修復は「生トークン → applications.ticket_token を検索」で行うため、
//    同じ spare_no を別人に割り当て直せると、先に記録された未解決注文が後の別人に紐づく。
//    誤紐づけの唯一の経路なので入口で塞ぐ。
function assignSpare(data) {
  const auth = _requireSession(data);
  if (!auth.ok) return _err(auth.error);
  _checkProps();

  const d      = data || {};
  const spareNo = _boothPadCode(String(d.spare_no || '').trim().normalize('NFKC').toUpperCase());
  const appId   = String(d.app_id || '').trim();
  if (!spareNo || !appId) return _err('INVALID_REQUEST');

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const spSh  = _ensureSpareBadgesSheet(ss);
  const appSh = _getSheet(ss, SHEET_APPLICATIONS);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(ADMIN_LOCK_WAIT_MS)) return _err('LOCK_BUSY');

  try {
    const spRows = spSh.getDataRange().getValues();
    let spRow = -1, spareToken = '', assignedTo = '';
    for (let i = 1; i < spRows.length; i++) {
      if (String(spRows[i][0] || '').trim() === spareNo) {
        spRow = i + 1;
        spareToken = String(spRows[i][1] || '').trim();
        assignedTo = String(spRows[i][2] || '').trim();
        break;
      }
    }
    if (spRow < 0)   return _err('SPARE_NOT_FOUND');
    if (!spareToken) return _err('SPARE_NO_TOKEN');   // 印刷前のデータ不整合。推測で発番しない

    if (assignedTo) {
      // 🔴 同じ人への再実行は成功として返す（受付の二度押し・再送で止まらないように）
      if (assignedTo === appId) {
        return _ok({ spare_no: spareNo, app_id: appId, ticket_token: spareToken, duplicate: true });
      }
      return { success: false, error: 'SPARE_ALREADY_ASSIGNED', assigned_app_id: assignedTo };
    }

    const appRows = appSh.getDataRange().getValues();
    let appRow = -1, appName = '', appSalon = '', prevToken = '';
    for (let i = 1; i < appRows.length; i++) {
      if (String(appRows[i][0] || '').trim() === appId) {
        appRow = i + 1;
        appSalon = String(appRows[i][4] || '');
        appName  = String(appRows[i][5] || '');
        prevToken = String(appRows[i][16] || '').trim();
        break;
      }
    }
    if (appRow < 0) return _err('APP_NOT_FOUND');

    // 🔴 書く順番が効く。spare_badges を先に、applications を後に書くこと。
    //    途中で落ちた場合:
    //      この順（spare先）  → 予備は「割当済み」なので他人へ付け替えられない。
    //                          booth/scan は spare_badges 経由で正しく本人へ解決できる。安全側
    //      逆順（app先）      → applications.ticket_token だけ書き換わり、予備は「未割当」のまま。
    //                          同じ予備名札を別人にも割り当てられてしまい、2人が同じトークンを持つ
    spSh.getRange(spRow, 3, 1, 2).setValues([[appId, _now()]]);   // assigned_app_id, assigned_at
    appSh.getRange(appRow, 17).setValue(spareToken);              // Q列 ticket_token を上書き
    appSh.getRange(appRow, 3).setValue(_now());                   // updated_at

    return _ok({
      spare_no: spareNo, app_id: appId, ticket_token: spareToken,
      staff_name: appName, salon_name: appSalon,
      previous_ticket_token: prevToken,   // 監査用。古いQRは以後どの申込にも当たらなくなる
      duplicate: false
    });
  } catch (err) {
    Logger.log('assignSpare error: ' + err);
    return _err('INTERNAL_ERROR');
  } finally {
    lock.releaseLock();
  }
}

// 一覧の行から入場パスを再送する（doPost: action=adminResendPass）。
// data: { session_token, app_id }
//
// 公開の resendPass（メールアドレス指定・該当の有無を隠して常に同じ応答を返す）とは別物。
// 社員用なので「送ったのか、抑止されたのか、メールが無いのか」を正直に返す。
function adminResendPass(data) {
  const auth = _requireSession(data);
  if (!auth.ok) return _err(auth.error);
  _checkProps();

  const appId = String((data || {}).app_id || '').trim();
  if (!appId) return _err('INVALID_REQUEST');

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rows = _getSheet(ss, SHEET_APPLICATIONS).getDataRange().getValues();

  let target = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() === appId) {
      target = {
        email:     String(rows[i][6] || '').trim(),
        salonName: String(rows[i][4] || ''),
        staffName: String(rows[i][5] || ''),
        token:     String(rows[i][16] || ''),
        status:    String(rows[i][17] || '')
      };
      break;
    }
  }
  if (!target)                     return _err('APP_NOT_FOUND');
  if (target.status === 'cancelled') return _err('APP_CANCELLED');
  if (!target.email)               return _err('NO_EMAIL');   // 代理登録でメール未取得の方
  if (!target.token)               return _err('NO_TICKET_TOKEN');

  // 🔴 直近10分に送っていれば重ねて送らない（v0.13.0 の再送抑止をそのまま使う）。
  //    公開経路と違い、抑止されたことを画面に返す
  if (_recentResendMailSent(target.email)) {
    return _ok({ sent: false, reason: 'RECENTLY_SENT', to: _maskEmail(target.email) });
  }

  // 🔴 その行1件だけを送る。公開の resendPass は同じメールの全申込を並べて送るが、
  //    社員が一覧の1行を選んで押している以上、送る相手はその1人が予測どおり
  _sendPassResendMail(target.email, [{
    salonName: target.salonName,
    staffName: target.staffName,
    passUrl:   SITE_BASE_URL + 'pass.html?t=' + target.token
  }]);

  return _ok({
    sent: true, to: _maskEmail(target.email),
    remaining_quota: MailApp.getRemainingDailyQuota()   // 🔴 当日は100通/日が効いてくる
  });
}

// 画面に出す用のメール伏せ字（社員用画面だが、肩越しに見えるので全部は出さない）
function _maskEmail(email) {
  const s = String(email || '');
  const at = s.indexOf('@');
  if (at <= 0) return s;
  const head = s.slice(0, at);
  const shown = head.length <= 2 ? head.slice(0, 1) : head.slice(0, 2);
  return shown + '***' + s.slice(at);
}
