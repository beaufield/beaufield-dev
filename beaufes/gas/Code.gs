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
// ============================================================

const VERSION  = '0.16.0-wip1';
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
  if (!token) return { valid: false };
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
          const r = { valid: false };
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
          return { valid: false };
        }

        const roles = rolesSh.getDataRange().getValues();
        let role = '';
        for (let j = 1; j < roles.length; j++) {
          if (String(roles[j][0]) === rowUserId && String(roles[j][1]) === APP_NAME) {
            role = String(roles[j][2] || '').trim().toLowerCase();
            break;
          }
        }
        if (!role || role === 'none') return { valid: false };

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
  const r = { valid: false };
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
  // ここを一緒くたにすると、Google側の一時不調のたびに社員がログアウトさせられる。
  return { ok: false, error: v.transient ? 'AUTH_TRANSIENT' : 'SESSION_INVALID' };
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
      // 🆕 診断用（diag.html）。読み取りのみ・データを一切変更しない
      case 'ping':              return _jsonResponse(pingLight(data));
      case 'pingHeavy':         return _jsonResponse(pingHeavy(data));
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
    _sendConfirmationMail(f.email, f.salonName, f.staffName, passUrl, false);
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
  } finally {
    lock.releaseLock();
  }

  return _ok({ app_id: appId, pass_url: SITE_BASE_URL + 'pass.html?t=' + token });
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
  } finally {
    lock.releaseLock();
  }

  const passUrl = SITE_BASE_URL + 'pass.html?t=' + ticketToken;
  // 更新時はメールを送らない（本人が画面上でパスURLをそのまま受け取れるため。updateApplicationと同じ考え方）
  if (!isUpdate) {
    _sendConfirmationMail(f.email, f.salonName, f.staffName, passUrl, false);
    try {
      _notifyNewApplicationLineWorks(appId, f, 'liff');
    } catch (e) {
      Logger.log('LINE WORKS通知に失敗（申込自体は成立済み）: ' + e);
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

  return _ok({ app_id: appId, pass_url: passUrl, is_update: isUpdate });
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
function _sendConfirmationMail(email, salonName, staffName, passUrl, isUpdate) {
  const cfg       = _getConfig();
  const eventDate = cfg.event_date || '2026-10-26';
  const eventTime = cfg.event_time || '10:00〜16:00';
  const venueName = cfg.venue_name || '青島屋（AOSHIMAYA）';
  const venueAddr = cfg.venue_addr || '宮崎市青島2丁目12-11';

  const subject = isUpdate
    ? '【ビューフェス2026】お申し込み内容を更新しました'
    : '【ビューフェス2026】お申し込みありがとうございます（入場パス）';

  const textBody =
    salonName + '\n' +
    staffName + ' 様\n\n' +
    'ビューフェス2026へのお申し込みを' + (isUpdate ? '更新' : '受付') + 'いたしました。\n\n' +
    '  日時 : ' + eventDate + ' ' + eventTime + '\n' +
    '  会場 : ' + venueName + ' ' + venueAddr + '\n\n' +
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
    '<p><a href="' + passUrl + '">▼ 入場パスを開く</a></p>' +
    '<p style="color:#666;font-size:13px;">' +
    '※このメールを保存いただくか、上のリンクをスマホのホーム画面に追加しておくと当日スムーズです</p>' +
    '<p style="color:#999;font-size:12px;">ビューフェス事務局（' + MAIL_FROM_ADDR + '）</p>';

  _sendMail(email, subject, textBody, htmlBody, isUpdate ? 'resend' : 'apply');
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
    appSh.getRange(1, 1, 1, 22).setValues([[
      'app_id', 'created_at', 'updated_at', 'source',
      'salon_name', 'staff_name', 'email', 'email_norm', 'phone', 'area', // areaは2026-08-06にフォームから削除・列は維持（空文字のみ）
      'has_transaction', 'address', 'referrer', 'agree_capability',
      'line_friend_id', 'line_user_id',
      'ticket_token', 'status', 'checked_in_at', 'note',
      'business_type',          // 🆕 U列（§4-1-2・v0.5.0で追加）
      'request_id'              // 🆕 V列（v0.13.0・§2-2 冪等キー）
    ]]);
    appSh.setFrozenRows(1);
    appSh.setColumnWidth(1, 110);
    Logger.log('applicationsシート作成完了');
  } else {
    Logger.log('applicationsシートは既に存在します');
  }

  // --- sessions シート（セミナー枠マスタ。今回は保留のため空のまま用意）---
  let sesSh = ss.getSheetByName(SHEET_SESSIONS);
  if (!sesSh) {
    sesSh = ss.insertSheet(SHEET_SESSIONS);
    sesSh.getRange(1, 1, 1, 9).setValues([[
      'session_id', 'slot', 'title', 'speaker', 'room',
      'starts_at', 'ends_at', 'capacity', 'is_active'
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
