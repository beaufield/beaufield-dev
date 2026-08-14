// ============================================================
// Beaufield 発注アプリ - Google Apps Script バックエンド
// ============================================================
// [重要] コードにIDを直書きしない。以下の手順でスクリプトプロパティに設定すること。
//
// GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」→「プロパティを追加」
//   SPREADSHEET_ID    : 発注管理データのスプレッドシートID
//   AUTH_SHEET_ID     : beaufield-auth スプレッドシートID（共通）
//   REORDER_API_KEY   : 発注点・発注提案更新用APIキー（Pythonスクリプトと共有）
//   LINEWORKS_WEBHOOK : LINE WORKS Incoming Webhook URL（任意・発注提案の通知用）
//
// 発注先マスターの拡張列（シートに直接入力・空欄なら既定値7日）:
//   F列 = リードタイム(日)   発注してから入荷するまでの日数
//   G列 = 発注サイクル(日)   そのメーカーへ発注する間隔（週1なら7）
//
// ============================================================

// スクリプトプロパティから機密値を取得（コードへの直書き禁止）
const _PROPS          = PropertiesService.getScriptProperties();
const SPREADSHEET_ID  = _PROPS.getProperty('SPREADSHEET_ID');
const AUTH_SHEET_ID   = _PROPS.getProperty('AUTH_SHEET_ID');
const UPDATE_SECRET   = _PROPS.getProperty('UPDATE_SECRET');   // 商品マスター更新用（Power Automate連携）
const VERSION         = 'v1.33.0';
const APP_NAME        = 'order-app';
const CACHE_TTL_SESSION = 60; // 権限変更・ログアウトを最大1分で反映
const PROP_STUCK_NOTIFY_DAYS = 14; // 提案滞留の通知・「要対応」表示の閾値（日）。Phase M, v1.31.0〜

// Google Drive上の商品マスターCSVファイル名
// ※ 同名ファイルが複数ある場合はファイルIDで指定（下記コメント参照）
const PRODUCT_CSV_NAME = '商品.CSV';

// ファイルIDで直接指定する場合はこちらを使う（より確実）
// Google DriveでファイルをID確認後に設定: setProductFileId() を実行
const PRODUCT_FILE_ID_KEY = 'PRODUCT_FILE_ID'; // ScriptPropertiesのキー名

// シート名定数
const SHEET_HISTORY   = '発注履歴';
const SHEET_ITEMS     = '発注明細';
const SHEET_SUPPLIERS = '発注先マスター';
const SHEET_PRODUCTS  = '商品マスター';
const SHEET_REORDER   = '発注点マスター';
const SHEET_PROPOSALS = '発注提案';       // analyze_demand.py が週次で書き込む発注提案リスト
const SHEET_PROPOSAL_EXCL = '提案除外設定'; // 発注提案の対象外にする商品（カタログ等のノイズ除外用）
const SHEET_EXCESS     = '過剰在庫';       // analyze_demand.py が週次で書き込む過剰在庫リスト
const SHEET_EXCESS_ACK = '過剰在庫確認済み'; // 持ちすぎを承知の上で確認済みにした商品（意図的なまとめ仕入等）
const SHEET_KPI_HISTORY = '在庫KPI履歴';    // analyze_demand.py が実行日ごとに1行追記（同日再実行は上書き）
const SHEET_EOL         = '終売商品設定';   // 在庫はあるが再発注できない商品（キャンペーン終了等）。提案除外設定とは別枠
const SHEET_LOT_OVERRIDE = '最低発注数設定';  // 過去発注実績からの自動推定より優先する手動の最低発注数設定
const SHEET_DEAD        = '死蔵在庫';       // analyze_demand.py が毎回全面書き換えする死蔵在庫リスト（Phase E, v1.10.0〜）
const SHEET_DEAD_ACK    = '死蔵在庫確認済み'; // 死蔵と承知の上で確認済みにした商品（季節品・サンプル等）
const SHEET_RECEIVED    = '入荷済み記録';    // 入荷待ちを手動で「✓入荷済み」にした発注（Phase F, v1.24.0〜）
const SHEET_RECEIPT_AUTO = '入荷実績（自動）'; // analyze_demand.pyが仕入データと突合して入荷確認できた発注（Phase G, v1.25.0〜）
const SHEET_POSTING_LAG  = '計上ラグ（自動）'; // 仕入先ごとの「仕入日→仕入入力日」の実績中央値（Phase G, v1.25.0〜）
const SHEET_ORDER_GROUPS = '発注グループ設定'; // 系列合計が一定本数の倍数でしか発注できないメーカーの設定（Phase J, v1.27.0〜）
const SHEET_GROUP_STATUS = '発注グループ状況'; // analyze_demand.py がグループ別の発注時期判定を毎回全面書き換え（Phase J, v1.27.0〜）
const SHEET_PROP_STUCK   = '提案滞留';       // 「いつから不足しているか」を持つ滞留トラッキング（Phase M, v1.31.0〜）

// 入荷待ち判定パラメータ（Phase F, v1.24.0〜）
// 提案の抑制（on_order加算）はリードタイムまでで変更しない。入荷待ちリストへの表示だけ
// 予定日超過後もPENDING_GRACE_DAYS日は残し、「遅延」として警告する（§3.4参照）
const PENDING_GRACE_DAYS = 14;

// 発注先マスターの拡張列（F=リードタイム(日), G=発注サイクル(日)）
// 未入力の場合のフォールバック値。analyze_demand.py の計算に使う
const DEFAULT_LEAD_TIME_DAYS  = 7;
const DEFAULT_ORDER_CYCLE_DAYS = 7;

// メーカー発注書テンプレート定義（Drive配信用）
// スクリプトプロパティ ORDER_TEMPLATE_FOLDER_ID に Drive フォルダIDを設定すること
const ORDER_TEMPLATES = {
  'grandex':           '54.pdf',
  'chiyoda':           '48.pdf',
  'alpenrose':         '57.pdf',
  'melos':             '2.pdf',
  'melos_2025':        'メロス発注書2025年価格改定後.pdf',
  'adelans':           '82.pdf',
  'rhythm':            'リズム注文書2023冬～.pdf',
  'hokkaido_natural':  '北海道ナチュラルバイオ.pdf'
};

// ============================================================
// セッション検証
// beaufield-auth の sessions シートでトークンを照合する
// CacheService で 15 分間キャッシュしてシート読み込みを削減する
// （ログアウト即時反映が必要な運用ではない前提。他アプリと同一パターン）
// 戻り値: { valid: true, user_id } または { valid: false }
// ============================================================
function validateSession(token) {
  if (!token) return { valid: false };

  const cache    = CacheService.getScriptCache();
  const cacheKey = 'sess_order_v2_' + token.slice(-32);
  const cached   = cache.get(cacheKey);
  if (cached !== null) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  try {
    const ss   = SpreadsheetApp.openById(AUTH_SHEET_ID);
    const sh   = ss.getSheetByName('sessions');
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
  } catch(e) {
    // 認証シートを読めなかった＝一時障害。ここも負キャッシュしない（上記と同じ理由）
    Logger.log('セッション検証エラー: ' + e);
    return { valid: false, transient: true };
  }
  // ここに到達＝シートは読めたがトークンが見つからなかった＝本物の無効
  const r = { valid: false };
  cache.put(cacheKey, JSON.stringify(r), 60);
  return r;
}

// ============================================================
// エントリーポイント（GET）
// ============================================================
function doGet(e) {
  return jsonResponse({ success: false, error: 'USE_POST' });
}

// ============================================================
// エントリーポイント（POST）
// application/x-www-form-urlencoded または application/json を受け付ける
// ============================================================
function doPost(e) {
  // JSON ボディを優先して解析（Power AutomateからのPOST対応）
  let p = {};
  let action = '';
  if (e && e.postData && e.postData.contents) {
    try {
      const body = JSON.parse(e.postData.contents);
      p = body;
      action = body.action || '';
    } catch(ex) {
      // JSON解析失敗 → form-encodedとして処理
    }
  }
  // form-encoded フォールバック
  if (!action && e && e.parameter) {
    p = e.parameter;
    action = p.action || '';
  }

  // updateProductMaster: シークレットキー認証（Power Automate用・セッション不要）
  if (action === 'updateProductMaster') {
    if (p.secret !== UPDATE_SECRET) {
      return jsonResponse({ success: false, error: 'UNAUTHORIZED' });
    }
    try {
      return jsonResponse(updateProductMaster(p.data || ''));
    } catch(err) {
      Logger.log('updateProductMaster error: ' + err);
      return jsonResponse({ success: false, error: 'INTERNAL_ERROR' });
    }
  }

  // APIキー認証アクション（Pythonスクリプト用・セッション不要）
  const API_KEY_ACTIONS = ['updateReorderPoints', 'getReorderConfig', 'updateOrderProposals', 'updateProposalExplanations', 'updateReceiptMatches', 'testNotify'];
  if (API_KEY_ACTIONS.indexOf(action) !== -1) {
    const apiKey = p.api_key || '';
    if (!apiKey || apiKey !== _PROPS.getProperty('REORDER_API_KEY')) {
      return jsonResponse({ success: false, error: 'UNAUTHORIZED' });
    }
    try {
      switch (action) {
        case 'updateReorderPoints':        return jsonResponse(updateReorderPoints(p.products || []));
        case 'getReorderConfig':           return jsonResponse(getReorderConfig());
        case 'updateOrderProposals':       return jsonResponse(updateOrderProposals(p));
        case 'updateProposalExplanations': return jsonResponse(updateProposalExplanations(p));
        case 'updateReceiptMatches':       return jsonResponse(updateReceiptMatches(p));
        case 'testNotify':                 return jsonResponse(testNotify());
      }
    } catch(err) {
      Logger.log(action + ' error: ' + err);
      return jsonResponse({ success: false, error: 'INTERNAL_ERROR' });
    }
  }

  // 通常のセッション検証
  const token = p.session_token || '';
  const auth = validateSession(token);
  if (!auth.valid) {
    // 認証シートを一時的に読めなかっただけの場合はSESSION_INVALIDにしない。
    // クライアント側はこれを見てログアウトさせず、リトライ対象として扱う
    if (auth.transient) {
      return jsonResponse({ success: false, error: 'AUTH_UNAVAILABLE', message: '認証確認に失敗しました。もう一度お試しください。' });
    }
    return jsonResponse({ success: false, error: 'SESSION_INVALID', message: '認証が必要です。ポータルからログインし直してください。' });
  }

  try {
    switch (action) {
      case 'saveOrder':    return jsonResponse(saveOrder(p, auth.user_id));
      case 'checkOrderByRequestId': return jsonResponse(checkOrderByRequestId(p, auth.user_id));
      case 'deleteOrder':  return jsonResponse(deleteOrder(p, auth.user_id));
      case 'saveSupplier': return jsonResponse(saveSupplier(p, auth.user_id));
      case 'saveProposalExclusion': return jsonResponse(saveProposalExclusion(p, auth.user_id));
      case 'saveExcessAck': return jsonResponse(saveExcessAck(p, auth.user_id));
      case 'saveDeadAck': return jsonResponse(saveDeadAck(p, auth.user_id));
      case 'saveReceived': return jsonResponse(saveReceived(p, auth.user_id));
      case 'saveEolFlag': return jsonResponse(saveEolFlag(p, auth.user_id));
      case 'saveLotOverride': return jsonResponse(saveLotOverride(p, auth.user_id));
      case 'getMasters':        return jsonResponse(getMasters());
      case 'getProductMaster':  return jsonResponse(getProductMaster());
      case 'getOrders':         return jsonResponse(getOrders(p.supplierCode || ''));
      case 'getOrderDetail':    return jsonResponse(getOrderDetail(p.orderNo));
      case 'getOrderTemplate':  return jsonResponse(getOrderTemplate(p.makerKey || ''));
      case 'getOrderProposals': return jsonResponse(getOrderProposals());
      default:             return jsonResponse({ success: false, error: '不明なアクション: ' + action });
    }
  } catch(err) {
    Logger.log('doPost error: ' + err);
    return jsonResponse({ success: false, error: 'INTERNAL_ERROR' });
  }
}

// ============================================================
// ヘルパー: JSONレスポンス生成
// ============================================================
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// ヘルパー: シート取得
// SpreadsheetApp.openById() は1リクエスト内で複数回呼ぶとオーバーヘッドになるため
// 実行コンテキスト内でキャッシュして使い回す
// ============================================================
let _ssCache = null;
function getSS() {
  if (!_ssCache) _ssCache = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ssCache;
}
function getSheet(name) {
  const sh = getSS().getSheetByName(name);
  if (!sh) throw new Error('シートが見つかりません: ' + name);
  return sh;
}

// ============================================================
// ヘルパー: 末尾限定読み込み（パフォーマンス改善_設計プラン.md 対策1）
//
// 発注履歴・発注明細は追記のみ（削除は稀）で、A列（発注No。YYYYMMDD-NNN形式）が
// 日付を含むため、行の物理的な並び順＝時系列の昇順になっている。アプリが実際に
// 使うのは常に「直近」のデータのため、シート全件ではなく末尾から必要な分だけを
// 読むことで、データが何年蓄積しても速度が変わらないようにする。
//
// 取りこぼしを防ぐため、必ず「十分読めたか」を確認しながら段階的に遡る。
// 十分と判定できないままシート先頭に達したら、結果的に全件を返す（＝従来と同じ動作を保証）。
// ============================================================

// A列の値だけを頼りに、「A列が key 以下になる行」まで遡った開始行番号を返す。
// 1セルずつの軽い読み取りで探索するため、対象範囲が広くても低コスト。
// （発注No・日付キーのように昇順に並んでいる列に対してのみ使えること）
function findTailStartRow_(sheet, key) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  let chunk = 300;
  while (true) {
    const startRow = Math.max(2, lastRow - chunk + 1);
    if (startRow === 2) return 2;
    const firstVal = String(sheet.getRange(startRow, 1).getValue() || '').trim();
    if (firstVal <= key) return startRow;
    chunk *= 4; // 見つかるまで大きく倍加し、往復回数（＝実行時間）を抑える
  }
}

// A列の値が orderNo と一致する行を末尾から探して読み取る（1発注分の明細取得等に使用）。
// 戻り値: { startRow, rows }（rows はヘッダーを含まない・シート上の並び順のまま）
function readRowsForKey_(sheet, key, numCols) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { startRow: 2, rows: [] };
  const startRow = findTailStartRow_(sheet, key);
  const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, numCols).getValues();
  return { startRow, rows };
}

// 末尾から段階的に広げて読み込み、isEnough(rows) が true になった時点（＝もう遡らなくて
// 良いと呼び出し元が判断できた時点）で打ち切る。count系の判定（「直近20件揃った」等）向け。
function readTailRowsUntil_(sheet, numCols, isEnough, initialChunk) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  let chunk = initialChunk || 300;
  while (true) {
    const startRow = Math.max(2, lastRow - chunk + 1);
    const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, numCols).getValues();
    if (startRow === 2 || isEnough(rows)) return rows;
    chunk *= 4;
  }
}

// ============================================================
// ヘルパー: セル値を安全に文字列化
// スプレッドシートが日付・日時型として認識したセルは getValues() で
// Dateオブジェクトとして返ってくる。そのまま String() すると
// "Thu Mar 26 2026 00:00:00 GMT+0900..." のGMT形式になってしまうため、
// Utilities.formatDate() を使って明示的にフォーマットする。
//
// fmt: 省略時は 'yyyy/MM/dd HH:mm'（日時用）
//       日付のみの場合は 'yyyy/MM/dd' を指定すること
// ============================================================
function cellToStr(val, fmt) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Tokyo', fmt || 'yyyy/MM/dd HH:mm');
  }
  return String(val || '');
}

// ============================================================
// GET: マスターデータ一括取得
// レスポンス: { success: true, suppliers: [...] }
// ============================================================
function getMasters() {
  const suppSheet  = getSheet(SHEET_SUPPLIERS);

  const suppData  = suppSheet.getDataRange().getValues();
  const suppliers = suppData.slice(1)
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      code:          String(r[0]).trim(),
      name:          String(r[1]).trim(),
      fax:           String(r[2] || '').trim(),
      // 発注方法: カンマ区切り文字列 → 配列。空欄は空配列（アプリ側で全ボタン表示）
      // 'Web' = 仕入先独自サイトで発注する仕入先（このアプリからは出力しない・在庫/数量確認用途）
      outputMethods: String(r[4] || '').trim()
                       .split(',')
                       .map(s => s.trim())
                       .filter(Boolean),
      // 備考: 最低発注金額等。発注入力画面の上部に常時表示する
      note:          String(r[7] || '').trim(),
      // 発注限時刻: 'HH:mm'形式。全発注先で必須に近い項目のため備考と別枠で管理
      // スプレッドシートが'11:00'等の文字列を時刻として自動認識しDateオブジェクト化することがあるため
      // cellToStrで両方のケースに対応する（生文字列ならそのまま、Dateなら'HH:mm'に整形）
      deadline:      cellToStr(r[8], 'HH:mm').trim()
    }));

  return { success: true, suppliers };
}

// ============================================================
// GET: 商品マスター取得
// Google Sheetsの「商品マスター」シートから読み込んで返す
// 「発注点マスター」シートのデータをJOINして reorderPoint フィールドを付与する
// レスポンス: { success: true, products: [...], updatedAt: '...' }
// ============================================================
function getProductMaster() {
  const ss = getSS();
  const sh = ss.getSheetByName(SHEET_PRODUCTS);
  if (!sh || sh.getLastRow() === 0) {
    return { success: true, products: [], updatedAt: '' };
  }

  const data    = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  const colMap = {
    code:            findColIdxGAS(headers, 'コード'),
    name:            findColIdxGAS(headers, '商品名'),
    kana:            findColIdxGAS(headers, 'かな'),
    unit:            findColIdxGAS(headers, '単位名'),
    supplierCD:      findColIdxGAS(headers, '仕入先CD'),
    supplierName:    findColIdxGAS(headers, '仕入先名'),
    makerCode:       findColIdxGAS(headers, '相手商品CD'),
    jan:             findColIdxGAS(headers, 'JANCD'),
    purchasePrice:   findColIdxGAS(headers, '仕入単価'),
    discontinued:    findColIdxGAS(headers, '廃番'),
    stockManagement: findColIdxGAS(headers, '在庫有無'),
    lastSaleDate:    findColIdxGAS(headers, '最終売上日'),
    stock:           findColIdxGAS(headers, '在庫数'),
    shelfMain:       findColIdxGAS(headers, '棚番１(本)'),
    shelfSub:        findColIdxGAS(headers, '棚番１(枝)')
  };

  const products = [];
  for (let i = 1; i < data.length; i++) {
    const r    = data[i];
    const name = colMap.name !== -1 ? String(r[colMap.name] || '').trim() : '';
    if (!name) continue;
    // 棚番: 「本」（英字の棚区画）＋「枝」（棚内の位置番号）を連結。
    // 「本」が未入力（空欄）の商品は棚番未設定として扱う（「枝」だけ入っていても既定値000のため無視）
    const shelfMainRaw = colMap.shelfMain !== -1 ? String(r[colMap.shelfMain] || '').trim() : '';
    const shelfSubRaw  = colMap.shelfSub  !== -1 ? String(r[colMap.shelfSub]  || '').trim() : '';
    const shelfNo       = shelfMainRaw ? (shelfMainRaw + shelfSubRaw) : '';
    products.push({
      code:            colMap.code !== -1            ? String(r[colMap.code]            || '').trim()                                        : '',
      name:            name,
      kana:            colMap.kana !== -1            ? String(r[colMap.kana]            || '').trim()                                        : '',
      unit:            colMap.unit !== -1            ? String(r[colMap.unit]            || '').trim()                                        : '',
      supplierCD:      colMap.supplierCD !== -1      ? String(r[colMap.supplierCD]      || '').trim()                                        : '',
      supplierName:    colMap.supplierName !== -1    ? String(r[colMap.supplierName]    || '').trim()                                        : '',
      makerCode:       colMap.makerCode !== -1       ? String(r[colMap.makerCode]       || '').trim()                                        : '',
      jan:             colMap.jan !== -1             ? String(r[colMap.jan]             || '').trim()                                        : '',
      purchasePrice:   colMap.purchasePrice !== -1   ? (parseFloat(String(r[colMap.purchasePrice] || '0').replace(/,/g, '')) || 0)          : 0,
      discontinued:    colMap.discontinued !== -1    ? String(r[colMap.discontinued]    || '').trim()                                        : '',
      stockManagement: colMap.stockManagement !== -1 ? String(r[colMap.stockManagement] || '').trim()                                        : '',
      lastSaleDate:    colMap.lastSaleDate !== -1    ? String(r[colMap.lastSaleDate]    || '').trim()                                        : '',
      stock:           colMap.stock !== -1           ? String(r[colMap.stock]           || '').trim()                                        : '',
      shelfNo:         shelfNo,
      reorderPoint:    null,
      reorderUpdatedAt: ''
    });
  }

  // 発注点マスターをJOIN
  try {
    const rsh = ss.getSheetByName(SHEET_REORDER);
    if (rsh && rsh.getLastRow() > 1) {
      const rdata = rsh.getDataRange().getValues();
      // ヘッダー: [0]=商品コード [1]=適正在庫（需要分析ベース） [2]=更新日時
      const reorderMap = {};
      for (let i = 1; i < rdata.length; i++) {
        const code = String(rdata[i][0] || '').trim();
        if (code) {
          reorderMap[code] = {
            reorderPoint:     parseFloat(rdata[i][1]) || 0,
            reorderUpdatedAt: cellToStr(rdata[i][2], 'yyyy/MM/dd')
          };
        }
      }
      products.forEach(p => {
        if (p.code && reorderMap[p.code]) {
          p.reorderPoint     = reorderMap[p.code].reorderPoint;
          p.reorderUpdatedAt = reorderMap[p.code].reorderUpdatedAt;
        }
      });
    }
  } catch(e) {
    Logger.log('発注点マスター読込エラー（無視）: ' + e);
  }

  const updatedAt = PropertiesService.getScriptProperties().getProperty('PM_UPDATED_AT') || '';
  return { success: true, products, updatedAt };
}

// ============================================================
// POST: 発注点マスター（適正在庫）更新（Pythonスクリプトからの自動実行用）
// v1.21.0〜: analyze_demand.py（需要分析）が全分析対象商品の適正在庫を送信する。
//           旧calc_reorder_point.py（6ヶ月月平均のみ）は基準統一のため使用停止。
// APIキー認証のみ（セッション不要）
// リクエスト: { action: 'updateReorderPoints', api_key: '...', products: [{code, reorderPoint, updatedAt}] }
// レスポンス: { success: true, count: N }
// ============================================================
function updateReorderPoints(products) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_REORDER);
  if (!sh) sh = ss.insertSheet(SHEET_REORDER);

  sh.clearContents();
  sh.getRange(1, 1, 1, 3).setValues([['商品コード', '適正在庫', '更新日時']]);

  if (!products || products.length === 0) {
    return { success: true, count: 0 };
  }

  const rows = products.map(p => [
    String(p.code       || '').trim(),
    parseFloat(p.reorderPoint) || 0,
    String(p.updatedAt  || '')
  ]);
  sh.getRange(2, 1, rows.length, 3).setValues(rows);

  Logger.log('✅ 発注点マスター更新完了: ' + rows.length + '件');
  return { success: true, count: rows.length };
}

// ============================================================
// タイマートリガー: Google Driveから商品マスターCSVを読み込んでSheets更新
// GASエディタのトリガー設定、または setDailyTrigger() で毎朝6時に自動実行される
// ============================================================
function updateProductMasterFromDrive() {
  let file = null;

  // まずScriptPropertiesにファイルIDが保存されていればそちらを優先
  const savedId = PropertiesService.getScriptProperties().getProperty(PRODUCT_FILE_ID_KEY);
  if (savedId) {
    try {
      file = DriveApp.getFileById(savedId);
    } catch(e) {
      Logger.log('保存済みファイルID無効。名前検索にフォールバック: ' + e);
    }
  }

  // ファイルIDがなければファイル名で検索（更新日時が最新のものを使用）
  if (!file) {
    const files = DriveApp.getFilesByName(PRODUCT_CSV_NAME);
    let latest = null;
    while (files.hasNext()) {
      const f = files.next();
      if (!latest || f.getLastUpdated() > latest.getLastUpdated()) latest = f;
    }
    if (!latest) throw new Error('Google Driveに「' + PRODUCT_CSV_NAME + '」が見つかりません');
    file = latest;
    // 次回以降はファイルIDで直接取得するよう保存
    PropertiesService.getScriptProperties().setProperty(PRODUCT_FILE_ID_KEY, file.getId());
  }

  const csvText = file.getBlob().getDataAsString('Shift-JIS');
  const rows    = parseCSVText(csvText);
  if (rows.length < 2) throw new Error('CSVが空か1行のみです');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_PRODUCTS);
  if (!sh) sh = ss.insertSheet(SHEET_PRODUCTS);

  sh.clearContents();
  // 大量データは一括書き込みで高速化
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  const updatedAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  PropertiesService.getScriptProperties().setProperty('PM_UPDATED_AT', updatedAt);
  Logger.log('✅ 商品マスター更新完了: ' + (rows.length - 1) + '件 (' + updatedAt + ')');
}

// ============================================================
// 毎朝6時の自動トリガーを設定する（GASエディタから【1回だけ】手動実行）
// 既存の同名トリガーがあれば先に削除してから再登録する
// ============================================================
function setDailyTrigger() {
  // 既存トリガーを削除（重複防止）
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'updateProductMasterFromDrive') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 毎朝6時（日本時間）に登録
  ScriptApp.newTrigger('updateProductMasterFromDrive')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .inTimezone('Asia/Tokyo')
    .create();
  Logger.log('✅ 毎朝6時トリガーを設定しました');
}

// ============================================================
// POST: 商品マスターCSVをSheets更新（base64経由・旧Power Automate用）
// ※ Google Drive直接読み込みに移行したため通常は使わない
// ============================================================
function updateProductMaster(base64Data) {
  if (!base64Data) return { success: false, error: 'dataが空です' };

  const bytes   = Utilities.base64Decode(base64Data);
  const blob    = Utilities.newBlob(bytes, 'text/plain', 'products.csv');
  const csvText = blob.getDataAsString('Shift-JIS');

  const rows = parseCSVText(csvText);
  if (rows.length < 2) return { success: false, error: 'CSVが空か1行のみです' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_PRODUCTS);
  if (!sh) sh = ss.insertSheet(SHEET_PRODUCTS);

  sh.clearContents();
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  const updatedAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  PropertiesService.getScriptProperties().setProperty('PM_UPDATED_AT', updatedAt);

  return { success: true, rows: rows.length - 1, updatedAt };
}

// ============================================================
// ヘルパー: CSVテキストを行×列の2次元配列に変換
// ============================================================
function parseCSVText(text) {
  const rows = [];
  const lines = text.split('\n');
  lines.forEach(line => {
    const clean = line.replace(/\r/g, '');
    if (clean.trim() === '') return;
    rows.push(splitCSVLineGAS(clean));
  });
  return rows;
}

function splitCSVLineGAS(line) {
  const result = []; let cur = '', inQ = false;
  for (const ch of line) {
    if (ch === '"') { inQ = !inQ; }
    else if (ch === ',' && !inQ) { result.push(cur); cur = ''; }
    else { cur += ch; }
  }
  result.push(cur);
  return result;
}

function findColIdxGAS(headers, name) {
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === name || headers[i].indexOf(name) !== -1) return i;
  }
  return -1;
}

// ============================================================
// GET: 発注履歴取得（直近20件・新しい順）
//      supplierCode を指定するとそのメーカーの直近20件を返す
//      + 取得した20件の明細から商品ごとの最新注文情報（productHistory）を返す
// レスポンス: { success: true, orders: [...], productHistory: {...} }
// ============================================================
function getOrders(filterSupplierCode) {
  filterSupplierCode = String(filterSupplierCode || '').trim();

  const sh = getSheet(SHEET_HISTORY);
  const matchesFilter = r => {
    const state = String(r[13] || '').trim();
    const visible = state === '' || state === 'COMPLETE'; // 移行前行は表示、保存途中は隠す
    return visible && r[0] !== '' && r[0] !== null &&
      (!filterSupplierCode || String(r[2] || '').trim() === filterSupplierCode);
  };

  // N列saveStateまで読み、PENDINGを通常履歴へ露出させない。
  // 「マッチする行が20件揃うまで」だけ末尾から遡って読む（対策1）
  const histRows = readTailRowsUntil_(sh, 14, rows => {
    let n = 0;
    for (const r of rows) { if (matchesFilter(r)) { n++; if (n >= 20) return true; } }
    return false;
  });

  if (histRows.length === 0) return { success: true, orders: [], productHistory: {} };

  // メーカー指定がある場合は先にフィルターしてから直近20件を取得する
  const orders = histRows
    .filter(matchesFilter)
    .map(r => ({
      orderNo:      String(r[0] || ''),
      // r[1] は発注日。スプレッドシートが日付型として認識するため cellToStr で変換
      date:         cellToStr(r[1], 'yyyy/MM/dd'),
      supplierCode: String(r[2] || ''),
      supplierName: String(r[3] || ''),
      fax:          String(r[4] || ''),
      staff:        String(r[5] || ''),
      itemCount:    r[6] || 0,
      outputType:   String(r[7] || ''),
      // r[8] は登録日時。同じく cellToStr で変換（デフォルト: 日時フォーマット）
      createdAt:    cellToStr(r[8])
    }))
    .reverse()
    .slice(0, 20);

  if (orders.length === 0) return { success: true, orders: [], productHistory: {} };

  // 直近20件の発注明細から商品ごとの最新注文情報を構築
  // 追加のAPI通信なし（発注明細シートをここで一括読込）
  const orderNos     = new Set(orders.map(o => o.orderNo));
  const orderDateMap = {};
  orders.forEach(o => { orderDateMap[o.orderNo] = o.date; });
  // 発注明細は発注No順（＝時系列順）に追記される前提。20件の中で最も古い発注Noより
  // 前まで読み終えたら、それより古い行にこの20件の明細は存在しないので打ち切る（対策1）
  const minOrderNo = orders.reduce((min, o) => (o.orderNo < min ? o.orderNo : min), orders[0].orderNo);

  const itemsSh   = getSheet(SHEET_ITEMS);
  const itemsData = readRowsForKey_(itemsSh, minOrderNo, 6).rows; // jan/code/qty/unit までの6列で足りる
  const productHistory = {}; // キー: 商品コードまたはJANコード → { date, qty, unit }

  itemsData.forEach(r => {
    const orderNo = String(r[0] || '').trim();
    if (!orderNos.has(orderNo)) return; // 直近20件以外はスキップ

    const jan  = String(r[1] || '').trim();
    const code = String(r[2] || '').trim();
    const qty  = r[4] || 0;
    const unit = String(r[5] || '').trim();
    const date = orderDateMap[orderNo] || '';

    // 商品コードとJANコードの両方をキーとして登録（より新しい日付で上書き）
    [code, jan].filter(Boolean).forEach(key => {
      if (!productHistory[key] || date > productHistory[key].date) {
        productHistory[key] = { date, qty, unit };
      }
    });
  });

  return { success: true, orders, productHistory };
}

// ============================================================
// GET: 発注明細取得
// レスポンス: { success: true, items: [...] }
// ============================================================
function getOrderDetail(orderNo) {
  if (!orderNo) return { success: false, error: 'orderNoが未指定です' };
  const sh     = getSheet(SHEET_ITEMS);
  const target = String(orderNo).trim();
  // 発注明細は発注No順に追記される前提で、対象の発注Noより前まで読み終えたら打ち切る（対策1）
  const data = readRowsForKey_(sh, target, 8).rows;

  const items = data
    .filter(r => String(r[0]).trim() === target)
    .map(r => ({
      jan:           String(r[1] || ''),
      code:          String(r[2] || ''),
      name:          String(r[3] || ''),
      qty:           r[4] || 0,
      unit:          String(r[5] || ''),
      memo:          String(r[6] || ''),
      isHandwritten: r[7] === 'TRUE'
    }));

  return { success: true, items };
}

// ============================================================
// POST: 発注保存
// user_id: validateSession() から取得した実際のユーザーID（改ざん不可）
// ============================================================
const ORDER_HISTORY_HEADERS = ['発注No','発注日','発注先コード','発注先名','FAX番号','担当者','品目数','出力方法','登録日時','user_id','requestId','requestHash','revisionBaseOrderNo','saveState'];

function ensureOrderHistorySchema_(histSh) {
  const current = histSh.getRange(1, 1, 1, ORDER_HISTORY_HEADERS.length).getValues()[0];
  let differs = false;
  for (let i = 0; i < ORDER_HISTORY_HEADERS.length; i++) {
    if (String(current[i] || '') !== ORDER_HISTORY_HEADERS[i]) { differs = true; break; }
  }
  if (differs) histSh.getRange(1, 1, 1, ORDER_HISTORY_HEADERS.length).setValues([ORDER_HISTORY_HEADERS]);
}

function findExactRows_(sheet, column, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  return sheet.getRange(2, column, lastRow - 1, 1)
    .createTextFinder(String(value))
    .matchEntireCell(true)
    .matchCase(true)
    .useRegularExpression(false)
    .findAll()
    .map(r => r.getRow())
    .sort((a, b) => a - b);
}

function normalizeOrderItem_(item) {
  const qty = Number(item && item.qty);
  return {
    janCode: String(item && item.janCode || '').trim(),
    code: String(item && item.code || '').trim(),
    name: String(item && item.name || ''),
    qty: Number.isFinite(qty) ? qty : 0,
    unit: String(item && item.unit || ''),
    memo: String(item && item.memo || ''),
    isHandwritten: item && (item.isHandwritten === true || String(item.isHandwritten).toUpperCase() === 'TRUE')
  };
}

function canonicalOrderPayload_(p, user_id, items) {
  return JSON.stringify([
    'order-v1', String(user_id || ''), String(p.date || '').trim(),
    String(p.supplierCode || '').trim(), String(p.supplierName || '').trim(),
    String(p.fax || '').trim(), String(p.staff || '').trim(), String(p.outputType || '').trim(),
    String(p.revisionBaseOrderNo || '').trim(),
    items.map(item => [item.janCode, item.code, item.name, item.qty, item.unit, item.memo, item.isHandwritten])
  ]);
}

function sha256Hex_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

function deleteExactRows_(sheet, column, value) {
  const rows = findExactRows_(sheet, column, value);
  for (let i = rows.length - 1; i >= 0; i--) sheet.deleteRow(rows[i]);
  return rows.length;
}

function replaceAndVerifyOrderItems_(itemsSh, orderNo, items, now) {
  deleteExactRows_(itemsSh, 1, orderNo);
  if (items.length > 0) {
    const rows = items.map(item => [
      orderNo, item.janCode, item.code, item.name, item.qty, item.unit, item.memo,
      item.isHandwritten ? 'TRUE' : 'FALSE', now
    ]);
    const startRow = itemsSh.getLastRow() + 1;
    // 自動数値変換で先頭ゼロのJANが欠落しないよう、文字列列だけを先にプレーンテキストへ固定する。
    // 数量(E列)・手書きフラグ(H列)・登録日時(I列)の書式は変更しない。
    itemsSh.getRange(startRow, 2, rows.length, 3).setNumberFormat('@');
    itemsSh.getRange(startRow, 6, rows.length, 2).setNumberFormat('@');
    itemsSh.getRange(startRow, 1, rows.length, 9).setValues(rows);
  }
  SpreadsheetApp.flush();
  const itemRows = findExactRows_(itemsSh, 1, orderNo);
  if (itemRows.length !== items.length) throw new Error('ITEM_WRITE_COUNT_MISMATCH');
  for (let i = 1; i < itemRows.length; i++) {
    if (itemRows[i] !== itemRows[0] + i) throw new Error('ITEM_WRITE_ROWS_NOT_CONTIGUOUS');
  }
  const values = items.length ? itemsSh.getRange(itemRows[0], 2, items.length, 7).getValues() : [];
  const actual = values.map(r => normalizeOrderItem_({
    janCode: r[0], code: r[1], name: r[2], qty: r[3], unit: r[4], memo: r[5], isHandwritten: r[6]
  }));
  if (JSON.stringify(actual) !== JSON.stringify(items)) throw new Error('ITEM_WRITE_CONTENT_MISMATCH');
}

function deleteOrderUnlocked_(orderNo, histSh, itemsSh) {
  const histRows = findExactRows_(histSh, 1, orderNo);
  // 通常の履歴取得が削除途中を完成済みとして読まないよう、先にPENDINGへ戻す。
  histRows.forEach(row => histSh.getRange(row, 14).setValue('PENDING'));
  const deletedItems = deleteExactRows_(itemsSh, 1, orderNo);
  for (let i = histRows.length - 1; i >= 0; i--) histSh.deleteRow(histRows[i]);
  return { deletedHist: histRows.length, deletedItems };
}

function saveOrder(p, user_id) {
  const requestId = String(p.requestId || '').trim();
  if (!requestId) return { success: false, error: 'REQUEST_ID_REQUIRED', message: 'requestIdが必要です' };
  if (!/^[A-Za-z0-9-]{8,100}$/.test(requestId)) {
    return { success: false, error: 'REQUEST_ID_INVALID', message: 'requestIdの形式が不正です' };
  }

  let rawItems;
  try { rawItems = JSON.parse(p.items || '[]'); }
  catch(e) { return { success: false, error: 'ITEMS_INVALID', message: '明細JSONを読み取れません' }; }
  if (!Array.isArray(rawItems) || rawItems.length === 0 || rawItems.length > 500) {
    return { success: false, error: 'ITEMS_INVALID', message: '明細は1〜500件で指定してください' };
  }
  const items = rawItems.map(normalizeOrderItem_);
  const date = String(p.date || '').trim();
  const supplierCode = String(p.supplierCode || '').trim();
  const supplierName = String(p.supplierName || '').trim();
  const fax = String(p.fax || '').trim();
  const staff = String(p.staff || '').trim();
  const outputType = String(p.outputType || '').trim();
  const revisionBaseOrderNo = String(p.revisionBaseOrderNo || '').trim();
  if (!date || !supplierCode || !supplierName || !staff) {
    return { success: false, error: 'REQUIRED_FIELDS_MISSING', message: '必須項目が不足しています' };
  }

  const requestHash = sha256Hex_(canonicalOrderPayload_(p, user_id, items));
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, error: 'LOCK_BUSY', message: '現在別の保存処理が実行中です' }; }

  try {
    const histSh = getSheet(SHEET_HISTORY);
    const itemsSh = getSheet(SHEET_ITEMS);
    ensureOrderHistorySchema_(histSh);
    const matches = findExactRows_(histSh, 11, requestId);
    if (matches.length > 1) {
      return { success: false, error: 'REQUEST_ID_CONFLICT', message: '同じrequestIdの履歴が複数あります' };
    }

    let histRow;
    let orderNo;
    if (matches.length === 1) {
      histRow = matches[0];
      const existing = histSh.getRange(histRow, 1, 1, 14).getValues()[0];
      orderNo = String(existing[0] || '').trim();
      const existingUser = String(existing[9] || '').trim();
      const existingHash = String(existing[11] || '').trim();
      const existingState = String(existing[13] || '').trim();
      if (existingUser && existingUser !== String(user_id || '').trim()) {
        return { success: false, error: 'REQUEST_ID_CONFLICT', message: 'requestIdの利用者が一致しません' };
      }
      if (existingHash && existingHash !== requestHash) {
        return { success: false, error: 'REQUEST_ID_CONFLICT', message: '同じrequestIdに異なる内容が送信されました' };
      }
      if (existingState === 'COMPLETE' && existingHash === requestHash) {
        return { success: true, orderNo, alreadyComplete: true };
      }
      // 移行前行またはPENDINGは、存在確認だけで済ませず同じorderNoへ全明細を書き直す。
    } else {
      orderNo = generateOrderNo(date);
      if (!orderNo || revisionBaseOrderNo === orderNo) {
        return { success: false, error: 'ORDER_STATE_INVALID', message: '発注Noまたは修正元の状態が不正です' };
      }
      // 空行を先にappendせず、次のsetValues 1回でPENDING行を作る。
      histRow = histSh.getLastRow() + 1;
    }
    if (!orderNo || revisionBaseOrderNo === orderNo) {
      return { success: false, error: 'ORDER_STATE_INVALID', message: '発注Noまたは修正元の状態が不正です' };
    }

    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    histSh.getRange(histRow, 1, 1, 14).setValues([[
      orderNo, date, supplierCode, supplierName, fax, staff, items.length, outputType,
      now, user_id, requestId, requestHash, revisionBaseOrderNo, 'PENDING'
    ]]);
    replaceAndVerifyOrderItems_(itemsSh, orderNo, items, now);
    if (revisionBaseOrderNo) deleteOrderUnlocked_(revisionBaseOrderNo, histSh, itemsSh);
    // 修正元行の削除で行番号が詰まるため、requestIdから現在行を取り直して完了にする。
    const finalRows = findExactRows_(histSh, 11, requestId);
    if (finalRows.length !== 1) throw new Error('REQUEST_ROW_LOST_DURING_SAVE');
    histSh.getRange(finalRows[0], 14).setValue('COMPLETE');
    SpreadsheetApp.flush();
    return { success: true, orderNo };
  } finally {
    lock.releaseLock();
  }
}

function checkOrderByRequestId(p, user_id) {
  const requestId = String(p.requestId || '').trim();
  if (!requestId) return { success: false, state: 'UNKNOWN', error: 'REQUEST_ID_REQUIRED' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return { success: true, state: 'UNKNOWN', reason: 'BUSY' };
  try {
    const histSh = getSheet(SHEET_HISTORY);
    ensureOrderHistorySchema_(histSh);
    const matches = findExactRows_(histSh, 11, requestId);
    if (matches.length === 0) return { success: true, state: 'NOT_FOUND_NOW' };
    if (matches.length > 1) return { success: true, state: 'CONFLICT' };
    const row = histSh.getRange(matches[0], 1, 1, 14).getValues()[0];
    const owner = String(row[9] || '').trim();
    if (owner && owner !== String(user_id || '').trim()) {
      return { success: false, state: 'UNKNOWN', error: 'FORBIDDEN' };
    }
    const state = String(row[13] || '').trim();
    return {
      success: true,
      state: state === 'COMPLETE' ? 'COMPLETE' : 'PARTIAL',
      orderNo: String(row[0] || '').trim()
    };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// POST: 発注履歴削除（管理者のみ）
// 発注履歴シートと発注明細シートから該当orderNoの行を物理削除する
// ============================================================
function deleteOrder(p, user_id) {
  const orderNo = String(p.orderNo || '').trim();
  if (!orderNo) return { success: false, error: '発注Noが未指定です' };

  // 管理者チェック
  if (!getIsAdmin(user_id)) {
    return { success: false, error: '削除は管理者のみ実行できます' };
  }

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, error: 'LOCK_BUSY', message: '現在別の保存処理が実行中です' }; }
  try {
    const histSh = getSheet(SHEET_HISTORY);
    const itemsSh = getSheet(SHEET_ITEMS);
    ensureOrderHistorySchema_(histSh);
    const result = deleteOrderUnlocked_(orderNo, histSh, itemsSh);
    if (result.deletedHist === 0) return { success: true, notFound: true, deletedItems: result.deletedItems };
    return { success: true, deletedItems: result.deletedItems };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// ヘルパー: ユーザーが管理者かどうかを確認
// beaufield-auth の users シートの列F（is_admin）を参照する
// ============================================================
function getIsAdmin(user_id) {
  try {
    const ss = SpreadsheetApp.openById(AUTH_SHEET_ID);
    const sh = ss.getSheetByName('users');
    if (!sh) return false;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(user_id).trim()) {
        // 列F（index 5）が TRUE または 'TRUE' の場合に管理者と判定
        return data[i][5] === true || String(data[i][5]).toUpperCase() === 'TRUE';
      }
    }
  } catch(e) {
    Logger.log('管理者チェックエラー: ' + e);
  }
  return false;
}

// 発注No採番（YYYYMMDD-NNN）
function generateOrderNo(dateStr) {
  const dateKey = dateStr.replace(/-/g, '');
  const sh = getSheet(SHEET_HISTORY);
  // 対象日より前の行まで読み終えたら打ち切る（発注Noは日付を含み昇順に並ぶ前提。対策1）。
  // A列（発注No）だけで判定できるので1列のみ読む（対策2）
  const rows = readTailRowsUntil_(sh, 1, chunkRows =>
    chunkRows.length > 0 && String(chunkRows[0][0] || '').trim() < dateKey
  );
  let maxSeq = 0;
  rows.forEach(r => {
    const no = String(r[0] || '');
    if (no.startsWith(dateKey + '-')) {
      const seq = parseInt(no.split('-')[1]) || 0;
      if (seq > maxSeq) maxSeq = seq;
    }
  });
  return dateKey + '-' + String(maxSeq + 1).padStart(3, '0');
}

// ============================================================
// POST: 発注先マスター操作
// ============================================================
function saveSupplier(p, user_id) {
  if (!getIsAdmin(user_id)) return { success: false, error: 'FORBIDDEN', message: '管理者権限が必要です' };
  const mode          = p.mode          || '';
  const code          = String(p.code          || '').trim();
  const name          = String(p.name          || '').trim();
  const fax           = String(p.fax           || '').trim();
  const outputMethods = String(p.outputMethods || '').trim(); // カンマ区切り文字列で受け取る
  const deadline      = String(p.deadline      || '').trim(); // 発注限時刻('HH:mm')
  const note          = String(p.note          || '').trim(); // 最低発注金額等

  if (!code) return { success: false, error: 'コードが未入力です' };

  const sh   = getSheet(SHEET_SUPPLIERS);
  const data = sh.getDataRange().getValues();
  const now  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  // 列: A=コード B=発注先名 C=FAX D=登録日時 E=発注方法 F=リードタイム(日) G=発注サイクル(日) H=備考 I=発注限時刻
  // F・G はシート直接入力のみ（このフォームからは触らない）
  // I列(発注限時刻)は'11:00'等の文字列を書き込むと、スプレッドシートが時刻として自動認識し
  // Dateシリアル値に変換してしまう（読み戻すと'Sat Dec 30 1899...'のような値になる不具合の原因）。
  // これを防ぐため、書き込み前にセルの表示形式を強制的にプレーンテキスト('@')にしておく
  if (mode === 'add') {
    const exists = data.slice(1).some(r => String(r[0]).trim() === code);
    if (exists) return { success: false, error: 'コード「' + code + '」はすでに登録されています' };
    sh.appendRow([code, name, fax, now, outputMethods, '', '', note, '']);
    const newRow = sh.getLastRow();
    sh.getRange(newRow, 9, 1, 1).setNumberFormat('@').setValue(deadline);
    return { success: true };
  } else if (mode === 'update') {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === code) {
        sh.getRange(i + 1, 1, 1, 5).setValues([[code, name, fax, now, outputMethods]]);
        sh.getRange(i + 1, 8, 1, 1).setValue(note);
        sh.getRange(i + 1, 9, 1, 1).setNumberFormat('@').setValue(deadline);
        return { success: true };
      }
    }
    return { success: false, error: 'コード「' + code + '」が見つかりません' };
  } else if (mode === 'delete') {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === code) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'コード「' + code + '」が見つかりません' };
  } else {
    return { success: false, error: '不明なmode: ' + mode };
  }
}

// ============================================================
// 発注提案機能（analyze_demand.py 連携）
// ============================================================
// フロー:
//   1. analyze_demand.py が getReorderConfig で設定取得
//      （発注先別リードタイム・除外商品・終売商品・過去発注からのロット推定材料）
//   2. 売上データを分析して updateOrderProposals で「発注提案」「過剰在庫」「死蔵在庫」
//      「在庫KPI履歴」シートを更新。提案・過剰在庫・死蔵在庫があれば
//      LINEWORKS_WEBHOOK（スクリプトプロパティ・任意）へ通知
//   3. アプリは getOrderProposals で提案・過剰在庫・死蔵在庫・経営KPI（直近2回分）を表示
//   4. ノイズ商品は saveProposalExclusion で除外登録（以後提案されない）
//   5. 持ちすぎを承知の上の商品は saveExcessAck で確認済み登録（次回分析後もリストから消える）
//   6. 在庫はあるが再発注できない商品（キャンペーン終了等）は saveEolFlag で終売登録
//      （除外設定とは別枠。以後提案されない）
//   7. ロット（最低注文数）が明確な商品は saveLotOverride で手動設定（自動推定より優先）
//   8. Claude Code が updateProposalExplanations でAI説明を追記（任意）
//   9. 死蔵と承知の上の商品は saveDeadAck で確認済み登録（次回分析後もリストから消える）
//  10. 入荷待ち（buildPendingOrders）はGAS内でリアルタイム集計。入荷確認できた発注は
//      saveReceived で登録すると recentOrders・入荷待ちリストの両方から即座に外れる
//  11. analyze_demand.py が仕入データ明細表.CSV と発注明細を突き合わせ、入荷確認できた発注を
//      updateReceiptMatches で「入荷実績（自動）」へ書き戻す（Phase G, v1.25.0〜）。
//      これにより手動✓なしで入荷待ちが自動消し込みされる（実測 3,039行→112行）。
//      あわせて仕入先別の「計上ラグ」も書き戻し、入荷待ちの遅延判定に使う
// ============================================================

// 発注提案シートの列定義（updateOrderProposals / getOrderProposals で共有）
// 区分列は v1.26.0 で追加。'提案'=通常の発注提案 / '参考'=自動条件では提案対象外だが
// 在庫が推奨在庫を下回っている商品（アプリではチェックOFF・グレー表示で金額合計に入れない）
// グループID・配分枠は v1.27.0（Phase J）で追加。まとめ発注グループに属する商品の行に入る
// 直近90日実需要は v1.31.0（Phase M）で追加。アプリの「要対応」表示の掲載判定に使う
const PROPOSAL_HEADERS = ['商品コード','商品名','仕入先コード','仕入先名','パターン',
                          '現在庫','発注済','推奨在庫','提案数量','最低発注数','月平均',
                          '注文P95','最大注文','根拠メモ','AI説明','分析日時','仕入単価','提案金額','ABCランク',
                          '区分','グループID','配分枠','直近90日実需要'];

// 提案滞留シートの列定義（updateOrderProposals / getOrderProposals で共有。Phase M, v1.31.0〜）
// 「いつから不足しているか」を持つ。発注提案シートは毎回全面書き換えなので、この情報だけは
// 別シートで前回との差分を追跡する必要がある（詳細: 発注提案精度改善_設計プラン.md M-1）
const PROP_STUCK_HEADERS = ['商品コード','商品名','初回不足日','最終確認日','直近在庫','直近不足数','区分'];

// 過剰在庫シートの列定義（updateOrderProposals / getOrderProposals で共有）
const EXCESS_HEADERS = ['商品コード','商品名','仕入先コード','仕入先名','現在庫','推奨在庫',
                        '過剰数量','仕入単価','過剰金額','在庫月数','ABCランク','パターン','分析日時'];

// 在庫KPI履歴シートの列定義（updateOrderProposals / getOrderProposals で共有）
const KPI_HEADERS = ['日付','在庫金額','月次売上原価','回転日数','過剰在庫額','過剰件数',
                     '提案額','提案件数','年間保有コスト概算','死蔵在庫額','死蔵件数'];

// 死蔵在庫シートの列定義（updateOrderProposals / getOrderProposals で共有。Phase E, v1.10.0〜）
const DEAD_HEADERS = ['商品コード','商品名','仕入先コード','仕入先名','現在庫','仕入単価',
                      '在庫金額','最終売上日','経過月数','区分','理由','分析日時'];

// 入荷済み記録シートの列定義（saveReceived / buildPendingOrders で共有。Phase F, v1.24.0〜）
const RECEIVED_HEADERS = ['発注No','商品コード','商品名','数量','確認者','確認日時'];

// 入荷実績（自動）シートの列定義（updateReceiptMatches / getReceivedKeys で共有。Phase G, v1.25.0〜）
// analyze_demand.py が仕入データ明細表.CSV と発注明細を突き合わせた結果を毎日全面書き換えする。
// 判定='入荷済み' … 仕入計上を確認できた（＝商品マスターの在庫数にも反映済み）
// 判定='打ち切り' … 発注から30日超えても仕入計上されない（欠品・キャンセル扱い。再提案される）
// どちらも入荷待ちリストからは除外する
const RECEIPT_AUTO_HEADERS = ['発注No','商品コード','判定','発注日','発注数量','分析日時'];

// 計上ラグ（自動）シートの列定義（Phase G, v1.25.0〜）
// 仕入日→仕入入力日の実績中央値。納品書の到着待ちで仕入入力が遅れる日数は仕入先ごとに
// 大きく異なる（実測: デミ2日 / 千代田化学7日 / ナプラ10日）ため、入荷待ちの遅延判定に使う
const POSTING_LAG_HEADERS = ['仕入先コード','計上ラグ日数','分析日時'];

// 計上ラグの実績が無い仕入先に使う既定値（analyze_demand.py の DEFAULT_POSTING_LAG_DAYS と揃える）
const DEFAULT_POSTING_LAG_DAYS = 6;

// ============================================================
// まとめ発注グループ（Phase J, v1.27.0〜。設計原本: まとめ発注グループ_設計プラン.md）
//
// 一部のメーカーは「系列合計が○本の倍数」でないと発注できない（1商品あたりも○本単位）。
// 商品単位の提案だけでは発注書が作れないため、系列でまとめて発注時期を判定し、
// 発注単位ぴったりになる組み合わせを analyze_demand.py が自動で組む。
//
// 対象メーカー・発注単位・所属判定ルールは「発注グループ設定」シート側で管理する
// （下の ORDER_GROUP_DEFAULTS は初回アクセス時の投入値。以後はシートが正）。
// 仕様と実データの検証内容は まとめ発注グループ_設計プラン.md 参照
// ============================================================
const ORDER_GROUP_HEADERS = ['グループID','グループ名','仕入先コード','発注単位（系列）','発注単位（商品）',
                             '名前に含む','名前に含まない','個別除外コード',
                             'トリガー割合(%)','必須枠日数','積み上げ上限日数','最低月需要','有効'];

// 初回アクセス時に投入する既定グループ（2026-07-30 時点の運用設定）
// ⚠️ 「名前に含む」のグラム表記は analyze_demand.py 側で全角ｇを半角gに正規化してから判定する。
//   商品名に全角ｇを使っているものが実際にあり、半角だけで判定すると取りこぼす
const ORDER_GROUP_DEFAULTS = [
  ['MILFY',       'ミルフィシリーズ',        '48', 120, 6, 'ミルフィ,120g',    'オキシ,OX', '', 50, 7, 60, 0.5, true],
  ['WAKAN18',     '和漢彩染 十八番',        '54',  72, 6, '十八番,120g',      'LUC',       '', 50, 7, 60, 0.5, true],
  ['WAKAN18_LUC', '和漢彩染 十八番 LUC',    '54',  72, 6, '十八番,120g,LUC',  '',          '', 50, 7, 60, 0.5, true]
];

// 発注グループ状況シートの列定義（updateOrderProposals / getOrderProposals で共有）
// analyze_demand.py が毎回全面書き換えする。1グループ1行
const GROUP_STATUS_HEADERS = ['グループID','グループ名','仕入先コード','仕入先名','発注単位（系列）','発注単位（商品）',
                              '商品数','月需要','系列在庫','系列発注済','系列適正在庫','不足合計','トリガー閾値',
                              '発注時期','提案本数','ロット数','系列在庫日数','発注時期までの目安日数','分析日時'];

// 発注グループ設定シートを取得（無ければ既定グループ入りで作成する）
function ensureOrderGroupSheet_() {
  const ss = getSS();
  let sh = ss.getSheetByName(SHEET_ORDER_GROUPS);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_ORDER_GROUPS);
  sh.getRange(1, 1, 1, ORDER_GROUP_HEADERS.length).setValues([ORDER_GROUP_HEADERS]);
  sh.getRange(2, 1, ORDER_GROUP_DEFAULTS.length, ORDER_GROUP_HEADERS.length)
    .setValues(ORDER_GROUP_DEFAULTS);
  // 仕入先コード・名前パターンは「48」「120g」等が数値・日付に化けないようテキスト固定にする
  // （v1.20.0 で発注限時刻の '11:00' がDateシリアル値に変換された事案と同種の予防）
  sh.getRange(2, 3, ORDER_GROUP_DEFAULTS.length, 1).setNumberFormat('@');
  sh.getRange(2, 6, ORDER_GROUP_DEFAULTS.length, 3).setNumberFormat('@');
  sh.setFrozenRows(1);
  Logger.log('✅ 発注グループ設定シートを既定値で作成しました');
  return sh;
}

// 発注グループ設定を読み出す（getReorderConfig から Python へ返す）
function readOrderGroups_() {
  const sh = ensureOrderGroupSheet_();
  if (sh.getLastRow() < 2) return [];
  const csv = s => String(s == null ? '' : s).split(',').map(x => x.trim()).filter(Boolean);
  const num = (v, d) => (parseFloat(v) > 0 ? parseFloat(v) : d);
  return sh.getDataRange().getValues().slice(1)
    .filter(r => String(r[0] || '').trim() !== '')
    .filter(r => r[12] !== false && String(r[12]).toUpperCase() !== 'FALSE')
    .map(r => ({
      groupId:      String(r[0]).trim(),
      groupName:    String(r[1] || '').trim(),
      supplierCode: String(r[2] || '').trim(),
      groupUnit:    num(r[3], 0),          // 系列の発注単位（120本 / 72本）
      itemUnit:     num(r[4], 6),          // 1商品の発注単位（6本）
      nameIncludes: csv(r[5]),
      nameExcludes: csv(r[6]),
      excludeCodes: csv(r[7]),
      triggerPct:   num(r[8], 50),
      mustDays:     num(r[9], 7),
      capDays:      num(r[10], 60),
      minMean:      parseFloat(r[11]) >= 0 ? parseFloat(r[11]) : 0.5
    }))
    .filter(g => g.groupUnit > 0 && g.itemUnit > 0 && g.nameIncludes.length > 0);
}

// POST(APIキー): 分析に必要な設定を返す
// レスポンス: { success, suppliers: [{code,name,leadTimeDays,orderCycleDays}],
//              exclusions: [商品コード], lotStats: {code: {orderCount,minQty,gcdQty}} }
function getReorderConfig() {
  // 発注先マスター（F列=リードタイム(日), G列=発注サイクル(日)。未入力はnull→Python側で既定値）
  const suppData = getSheet(SHEET_SUPPLIERS).getDataRange().getValues();
  const suppliers = suppData.slice(1)
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      code:           String(r[0]).trim(),
      name:           String(r[1]).trim(),
      leadTimeDays:   parseFloat(r[5]) > 0 ? parseFloat(r[5]) : null,
      orderCycleDays: parseFloat(r[6]) > 0 ? parseFloat(r[6]) : null
    }));

  // 提案除外設定（シート未作成なら空）
  let exclusions = [];
  const exSh = getSS().getSheetByName(SHEET_PROPOSAL_EXCL);
  if (exSh && exSh.getLastRow() > 1) {
    exclusions = exSh.getDataRange().getValues().slice(1)
      .map(r => String(r[0]).trim()).filter(Boolean);
  }

  // 終売商品設定（シート未作成なら空。提案除外設定とは別枠で管理）
  let eolCodes = [];
  const eolSh = getSS().getSheetByName(SHEET_EOL);
  if (eolSh && eolSh.getLastRow() > 1) {
    eolCodes = eolSh.getDataRange().getValues().slice(1)
      .map(r => String(r[0]).trim()).filter(Boolean);
  }

  // 手動ロット設定（シート未作成なら空。過去発注実績からの自動推定より優先して使う）
  let lotOverrides = {};
  const lotSh = getSS().getSheetByName(SHEET_LOT_OVERRIDE);
  if (lotSh && lotSh.getLastRow() > 1) {
    lotSh.getDataRange().getValues().slice(1).forEach(r => {
      const code = String(r[0] || '').trim();
      const lot  = parseInt(r[2], 10);
      if (code && lot > 0) lotOverrides[code] = lot;
    });
  }

  // 過去の発注明細から商品別のロット推定材料を集計
  // gcdQty: 全発注数量の最大公約数（ケース単位の推定に使う）
  // あわせて直近90日の発注（発注済み・未入荷の可能性がある分）も商品別に返す
  // （Phase F, v1.24.0〜: 60日→90日に延長。長めのリードタイムの仕入先でも取りこぼさないため）
  // 発注日は発注No（YYYYMMDD-NNN）の先頭8桁から取得
  // 手動で「✓入荷済み」登録済みの発注（発注No+商品コード）は在庫に反映済みとみなし、
  // ここで除外する（在庫と二重加算しないため）。
  // ⚠️ 自動判定分（入荷実績（自動））は**ここでは除外しない**（getReceivedKeys の第1引数=false）。
  //   recentOrders は analyze_demand.py が突合の入力として使うため、自動判定済みの発注を
  //   ここで落とすと翌日の突合対象から消え、「入荷実績（自動）」を再生成できなくなる
  //   （毎回全面書き換えのため、消えた発注が翌日また未入荷に戻って発振する）
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 90);
  const cutoffStr = Utilities.formatDate(cutoff, 'Asia/Tokyo', 'yyyyMMdd');
  const receivedKeys = getReceivedKeys(false);

  const itemsData = getSheet(SHEET_ITEMS).getDataRange().getValues();
  const lotStats = {};
  const recentOrders = {};
  itemsData.slice(1).forEach(r => {
    const orderNo = String(r[0] || '').trim();
    const code = String(r[2] || '').trim();
    const qty  = Math.round(parseFloat(r[4]) || 0);
    if (!code || qty <= 0) return;
    if (!lotStats[code]) lotStats[code] = { orderCount: 0, minQty: qty, gcdQty: 0 };
    const s = lotStats[code];
    s.orderCount++;
    if (qty < s.minQty) s.minQty = qty;
    s.gcdQty = gcdInt(s.gcdQty, qty);

    const orderDateKey = orderNo.split('-')[0];  // 発注No先頭のYYYYMMDD
    if (orderDateKey.length === 8 && orderDateKey >= cutoffStr && !receivedKeys.has(orderNo + '|' + code)) {
      if (!recentOrders[code]) recentOrders[code] = [];
      // orderNo は Phase G(v1.25.0)で追加。analyze_demand.py が発注明細行ごとの
      // 入荷判定結果を「入荷実績（自動）」へ書き戻すためのキーとして使う
      recentOrders[code].push({ orderNo, date: orderDateKey, qty });
    }
  });

  // まとめ発注グループ設定（Phase J, v1.27.0〜。シート未作成なら既定3グループで自動作成）
  // 所属商品の判定は商品名で行うため Python 側（商品.CSVを持っている）で実施する。
  // ここでは設定を返すだけにして、判定ロジックを1箇所（analyze_demand.py）に集約する
  const orderGroups = readOrderGroups_();

  return {
    success: true,
    suppliers,
    exclusions,
    eolCodes,
    lotStats,
    lotOverrides,
    recentOrders,
    orderGroups,
    defaults: { leadTimeDays: DEFAULT_LEAD_TIME_DAYS, orderCycleDays: DEFAULT_ORDER_CYCLE_DAYS }
  };
}

function gcdInt(a, b) {
  a = Math.abs(a); b = Math.abs(b);
  while (b) { const t = b; b = a % b; a = t; }
  return a;
}

// 入荷確認済みの発注を (発注No + '|' + 商品コード) のSetで返す
// getReorderConfig（recentOrders除外）と buildPendingOrders（入荷待ちリスト除外）の両方で使う
//
// v1.25.0(Phase G)〜: 手動の「✓入荷済み」に加えて、analyze_demand.py が仕入データと
// 突き合わせて自動判定した「入荷実績（自動）」も合わせて除外対象にする。
// 自動判定が大半を処理し、手動✓は自動で拾えないケースのオーバーライドとして残る。
//
// ⚠️ getReorderConfig から呼ぶ場合は includeAuto=false にすること。
//   recentOrders は analyze_demand.py が突合の入力として使うため、ここで自動判定分を
//   先に除外してしまうと、突合対象から消えて「入荷実績（自動）」を再生成できなくなる
//   （＝一度入荷済みになった発注が翌日以降リストから永久に落ちる）。
function getReceivedKeys(includeAuto) {
  const keys = new Set();
  const ss = getSS();

  const sh = ss.getSheetByName(SHEET_RECEIVED);
  if (sh && sh.getLastRow() > 1) {
    sh.getDataRange().getValues().slice(1).forEach(r => {
      const orderNo = String(r[0] || '').trim();
      const code    = String(r[1] || '').trim();
      if (orderNo && code) keys.add(orderNo + '|' + code);
    });
  }

  if (includeAuto) {
    const autoSh = ss.getSheetByName(SHEET_RECEIPT_AUTO);
    if (autoSh && autoSh.getLastRow() > 1) {
      autoSh.getDataRange().getValues().slice(1).forEach(r => {
        const orderNo = String(r[0] || '').trim();
        const code    = String(r[1] || '').trim();
        if (orderNo && code) keys.add(orderNo + '|' + code);
      });
    }
  }
  return keys;
}

// 仕入先コード → 計上ラグ日数（「計上ラグ（自動）」シート。未登録は既定値）
function getPostingLagByCode() {
  const map = {};
  const sh = getSS().getSheetByName(SHEET_POSTING_LAG);
  if (sh && sh.getLastRow() > 1) {
    sh.getDataRange().getValues().slice(1).forEach(r => {
      const code = String(r[0] || '').trim();
      const lag  = parseFloat(r[1]);
      if (code && lag >= 0) map[code] = lag;
    });
  }
  return map;
}

// カレンダー日付(y,m,d)をUTC起点の通し日数(epoch day)に変換する。
// タイムゾーン・DST（日本には無いが念のため）の影響を受けずに日数差分を計算するため、
// Dateオブジェクトの時刻部分ではなく整数の「日数」だけで加減算する
function ymdToEpochDay(y, m, d) {
  return Math.floor(Date.UTC(y, m - 1, d) / 86400000);
}
function epochDayToDateStr(epochDay) {
  return Utilities.formatDate(new Date(epochDay * 86400000), 'Asia/Tokyo', 'yyyy-MM-dd');
}

// 入荷待ちリストを組み立てる（Phase F, v1.24.0〜 / Phase G, v1.25.0で判定根拠を変更）。
// GAS内でリアルタイム集計するため、アプリから発注した直後にgetOrderProposalsを呼べば即座に反映される。
//
// 対象: 発注明細（アプリ経由の発注のみ・電話/FAX等は対象外）のうち、
//   「✓入荷済み」（手動）にも「入荷実績（自動）」にも入っていない行。
//
// v1.25.0(Phase G)の変更点:
//   1. 除外判定に analyze_demand.py の突合結果（入荷実績（自動））を加えた。
//      これにより仕入計上を確認できた発注は手動✓なしで自動的にリストから消える
//      （実測: 3,039行 → 112行）
//   2. 遅延判定を「発注日+リードタイム」から「発注日+リードタイム+計上ラグ」に変更した。
//      納品書の到着待ちで仕入入力が数日遅れるのは正常な状態であり、これを遅延として
//      表示すると実質すべてが「予定日超過」になってリストとして機能しないため
//      （実測: 千代田化学は計上ラグ7日・リードタイム1日）
//
// ⚠️ ここでの「遅延」はあくまで警告表示用で、提案そのものは止めない（Phase Fからの原則）
function buildPendingOrders() {
  const ss = getSS();
  const itemsSh = ss.getSheetByName(SHEET_ITEMS);
  if (!itemsSh || itemsSh.getLastRow() < 2) return [];

  // 発注先コード → リードタイム（発注先マスターF列。未入力は既定値）
  // ※ 発注先マスターは数十行程度で増え続けないシートなので全件読みのままでよい
  const suppData = getSheet(SHEET_SUPPLIERS).getDataRange().getValues();
  const leadTimeByCode = {};
  suppData.slice(1).forEach(r => {
    const code = String(r[0] || '').trim();
    if (code) leadTimeByCode[code] = parseFloat(r[5]) > 0 ? parseFloat(r[5]) : DEFAULT_LEAD_TIME_DAYS;
  });

  // 手動✓と自動突合の両方を除外対象にする（Phase G）
  const receivedKeys = getReceivedKeys(true);
  const postingLagByCode = getPostingLagByCode();

  // 発注明細・発注履歴は「入荷待ちとして意味を持つ期間」だけに絞って読む（対策1）。
  // 猶予期限（graceDeadline）を過ぎた発注はどのみち対象外になるので、
  // 「最大リードタイム＋最大計上ラグ＋猶予日数」より前の発注は読む必要が無い
  const leadTimeValues = Object.values(leadTimeByCode);
  const maxLeadTime    = leadTimeValues.length > 0 ? Math.max(DEFAULT_LEAD_TIME_DAYS, ...leadTimeValues) : DEFAULT_LEAD_TIME_DAYS;
  const lagValues      = Object.values(postingLagByCode);
  const maxPostingLag  = lagValues.length > 0 ? Math.max(DEFAULT_POSTING_LAG_DAYS, ...lagValues) : DEFAULT_POSTING_LAG_DAYS;
  const scanWindowDays = Math.ceil(maxLeadTime) + Math.ceil(maxPostingLag) + PENDING_GRACE_DAYS + 5; // +5日は安全マージン
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - scanWindowDays);
  const cutoffKey = Utilities.formatDate(cutoffDate, 'Asia/Tokyo', 'yyyyMMdd');

  // 発注No/JAN/コード/商品名/数量の5列で足りる
  const itemsData = readTailRowsUntil_(itemsSh, 5, rows =>
    rows.length > 0 && String(rows[0][0] || '').trim() < cutoffKey
  );

  // 発注No → 仕入先情報（発注履歴シートから。PENDINGは入荷待ち集計へ含めない）
  const histSh = ss.getSheetByName(SHEET_HISTORY);
  const supplierByOrderNo = {};
  if (histSh && histSh.getLastRow() > 1) {
    readTailRowsUntil_(histSh, 14, rows =>
      rows.length > 0 && String(rows[0][0] || '').trim() < cutoffKey
    ).forEach(r => {
      const orderNo = String(r[0] || '').trim();
      const state = String(r[13] || '').trim();
      if (orderNo && (state === '' || state === 'COMPLETE')) {
        supplierByOrderNo[orderNo] = { code: String(r[2] || '').trim(), name: String(r[3] || '') };
      }
    });
  }

  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const todayParts = todayStr.split('-').map(Number);
  const todayEpochDay = ymdToEpochDay(todayParts[0], todayParts[1], todayParts[2]);

  const candidates = [];
  itemsData.forEach(r => {
    const orderNo = String(r[0] || '').trim();
    const code    = String(r[2] || '').trim();
    const name    = String(r[3] || '');
    const qty     = Math.round(parseFloat(r[4]) || 0);
    if (!orderNo || !code || qty <= 0) return;
    if (receivedKeys.has(orderNo + '|' + code)) return;

    const dateKey = orderNo.split('-')[0]; // 発注No先頭のYYYYMMDD
    if (dateKey.length !== 8) return;
    const y = parseInt(dateKey.slice(0, 4), 10), m = parseInt(dateKey.slice(4, 6), 10), d = parseInt(dateKey.slice(6, 8), 10);
    const orderEpochDay = ymdToEpochDay(y, m, d);

    if (!supplierByOrderNo[orderNo]) return; // PENDINGや履歴削除済みの明細は集計しない
    const supp = supplierByOrderNo[orderNo];
    const leadTimeDays = leadTimeByCode[supp.code] || DEFAULT_LEAD_TIME_DAYS;
    // 物理的な到着予定日（従来どおり）
    const expectedEpochDay = orderEpochDay + Math.round(leadTimeDays);
    // 仕入計上まで見込んだ日（納品書の到着待ち分を上乗せ）。遅延判定はこちらを基準にする
    const postingLagDays = postingLagByCode[supp.code] !== undefined
      ? postingLagByCode[supp.code] : DEFAULT_POSTING_LAG_DAYS;
    const postingDueEpochDay = expectedEpochDay + Math.round(postingLagDays);
    const graceDeadlineEpochDay = postingDueEpochDay + PENDING_GRACE_DAYS;
    if (graceDeadlineEpochDay < todayEpochDay) return; // 猶予期間も過ぎたら自動的に対象外

    const overdueDays = todayEpochDay - postingDueEpochDay;
    candidates.push({
      code, name,
      supplierName: supp.name,
      orderNo,
      orderDate: dateKey.slice(0, 4) + '-' + dateKey.slice(4, 6) + '-' + dateKey.slice(6, 8),
      qty,
      leadTimeDays,
      postingLagDays,
      expectedDate: epochDayToDateStr(expectedEpochDay),
      postingDueDate: epochDayToDateStr(postingDueEpochDay),
      elapsedDays: todayEpochDay - orderEpochDay,
      overdueDays,
      isDelayed: overdueDays > 0
    });
  });

  if (candidates.length === 0) return [];

  // 仕入単価を商品マスターから引く（対象コードのみに絞って走査）。
  // 商品マスターは9,000件超×約20列あるが、必要なのは「コード」「仕入単価」の2列だけなので、
  // その2列だけを読み込む（対策2）。ヘッダー行だけ先に読んで列位置を特定してから読む
  const codeSet = new Set(candidates.map(c => c.code));
  const unitCostByCode = {};
  const prodSh = ss.getSheetByName(SHEET_PRODUCTS);
  if (prodSh && prodSh.getLastRow() > 1) {
    const prodLastRow = prodSh.getLastRow();
    const headers = prodSh.getRange(1, 1, 1, prodSh.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const colCode = findColIdxGAS(headers, 'コード');
    const colCost = findColIdxGAS(headers, '仕入単価');
    if (colCode !== -1 && colCost !== -1) {
      const codeCol = prodSh.getRange(2, colCode + 1, prodLastRow - 1, 1).getValues();
      const costCol = prodSh.getRange(2, colCost + 1, prodLastRow - 1, 1).getValues();
      for (let i = 0; i < codeCol.length; i++) {
        const code = String(codeCol[i][0] || '').trim();
        if (codeSet.has(code) && !(code in unitCostByCode)) {
          unitCostByCode[code] = parseFloat(String(costCol[i][0] || '0').replace(/,/g, '')) || 0;
        }
      }
    }
  }

  candidates.forEach(c => {
    c.unitCost = unitCostByCode[c.code] || 0;
    c.amount = Math.round(c.unitCost * c.qty);
  });

  // 遅延を最上位、その中で発注日の古い順
  candidates.sort((a, b) => {
    if (a.isDelayed !== b.isDelayed) return a.isDelayed ? -1 : 1;
    return a.orderDate < b.orderDate ? -1 : (a.orderDate > b.orderDate ? 1 : 0);
  });

  return candidates;
}

// 金額を「◯◯.◯万円」表記にする（LINE WORKS通知用）
function formatManYen(yen) {
  return (yen / 10000).toFixed(1) + '万円';
}

// 'YYYY-MM-DD' 文字列同士の日数差（UTC正午基準で計算しタイムゾーンの影響を避ける）
function daysBetweenIso_(fromIso, toIso) {
  const toMs = s => new Date(s + 'T12:00:00Z').getTime();
  return Math.round((toMs(toIso) - toMs(fromIso)) / 86400000);
}

// 提案滞留の更新（Phase M, v1.31.0。設計原本: 発注提案精度改善_設計プラン.md M-1）。
// 「発注提案」シートは毎回全面書き換えなので前回との差分が消える。ここだけ別シートで
// 「いつから不足しているか」を追跡する。滞留日数は「今日−初回不足日」の日数差で持つ
// （連続実行回数ではない）ため、分析バッチが飛んだ日があってもズレない。
// まとめ発注グループ所属商品（groupId有り）は「発注グループ状況」シートで発注時期を
// 別管理しているため対象外（refOnly=trueが「まだ発注時期でない」正常な待機状態であり、
// 個別商品の「見落とし」とは意味が異なるため）
function updateProposalStuckTracking_(ss, proposals, analyzedAt) {
  const today = (String(analyzedAt || '').match(/^\d{4}-\d{2}-\d{2}/) || [])[0]
                || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  // 今回時点で「不足」している商品（推奨在庫 − 在庫 − 発注済 > 0）を code をキーに集約
  const shortNow = {};
  proposals.forEach(x => {
    if (x.groupId) return;
    const code = String(x.code || '').trim();
    if (!code) return;
    const stock = parseFloat(x.stock) || 0;
    const onOrder = parseFloat(x.onOrder) || 0;
    const recommended = parseFloat(x.recommended) || 0;
    const shortQty = recommended - stock - onOrder;
    if (shortQty <= 0) return;
    shortNow[code] = {
      name: String(x.name || ''), stock: stock, shortQty: shortQty,
      kubun: x.refOnly ? '参考' : '提案',
    };
  });

  let sh = ss.getSheetByName(SHEET_PROP_STUCK);
  if (!sh) sh = ss.insertSheet(SHEET_PROP_STUCK);

  const prevShortSince = {};
  if (sh.getLastRow() > 1) {
    sh.getDataRange().getValues().slice(1).forEach(r => {
      const code = String(r[0] || '').trim();
      if (code) prevShortSince[code] = String(r[2] || '');
    });
  }

  const entries = Object.keys(shortNow).map(code => {
    const s = shortNow[code];
    const shortSince = prevShortSince[code] || today;   // 既存レコードがあれば初回不足日を据え置き
    return { code: code, name: s.name, shortSince: shortSince, stock: s.stock,
             shortQty: s.shortQty, kubun: s.kubun, stuckDays: daysBetweenIso_(shortSince, today) };
  });

  sh.clearContents();
  sh.getRange(1, 1, 1, PROP_STUCK_HEADERS.length).setValues([PROP_STUCK_HEADERS]);
  if (entries.length > 0) {
    const rows = entries.map(e => [e.code, e.name, e.shortSince, today, e.stock, e.shortQty, e.kubun]);
    sh.getRange(2, 1, rows.length, PROP_STUCK_HEADERS.length).setValues(rows);
  }
  return entries;   // 通知（updateOrderProposals）で滞留日数の集計に使う
}

// POST(APIキー): 発注提案・過剰在庫・死蔵在庫シートを全面書き換え、在庫KPI履歴に1行記録
// リクエスト: { proposals: [{code,name,supplierCode,supplierName,pattern,stock,recommended,
//              proposedQty,lot,meanMonthly,p95Order,maxOrder,note,refOnly}],
//              ※ refOnly=true は「自動条件では提案対象外だが在庫が推奨在庫を下回った」参考行。
//                 シートの区分列に'参考'と書き、KPI・LINE WORKS通知の件数/金額からは除外する
//              excess: [{code,name,supplierCode,supplierName,stock,recommended,excessQty,
//              unitCost,excessAmount,monthsOfStock,abcRank,pattern}],
//              dead: [{code,name,supplierCode,supplierName,stock,unitCost,deadAmount,
//              lastSaleDate,monthsSinceLastSale,tier,reason}],
//              kpi: {date,stockValue,monthlyCogs,turnoverDays,excessAmount,excessCount,
//              proposalAmount,proposalCount,holdingCostAnnual,deadAmount,deadCount}, analyzedAt }
function updateOrderProposals(p) {
  const proposals  = p.proposals || [];
  const excess     = p.excess || [];
  const dead       = p.dead || [];
  const kpi        = p.kpi || null;
  const analyzedAt = String(p.analyzedAt || '');

  const ss = getSS();

  // 提案滞留の更新（Phase M, v1.31.0）。「発注提案」シートを上書きする前に、
  // 今回の proposals から不足状態を判定して前回との差分を取る必要がある
  const stuckEntries = updateProposalStuckTracking_(ss, proposals, analyzedAt);

  let sh = ss.getSheetByName(SHEET_PROPOSALS);
  if (!sh) sh = ss.insertSheet(SHEET_PROPOSALS);

  sh.clearContents();
  sh.getRange(1, 1, 1, PROPOSAL_HEADERS.length).setValues([PROPOSAL_HEADERS]);

  if (proposals.length > 0) {
    const rows = proposals.map(x => [
      String(x.code         || '').trim(),
      String(x.name         || ''),
      String(x.supplierCode || ''),
      String(x.supplierName || ''),
      String(x.pattern      || ''),
      parseFloat(x.stock)       || 0,
      parseFloat(x.onOrder)     || 0,
      parseFloat(x.recommended) || 0,
      parseFloat(x.proposedQty) || 0,
      parseFloat(x.lot)         || 1,
      parseFloat(x.meanMonthly) || 0,
      parseFloat(x.p95Order)    || 0,
      parseFloat(x.maxOrder)    || 0,
      String(x.note || ''),
      '',           // AI説明（updateProposalExplanations で追記）
      analyzedAt,
      parseFloat(x.unitCost) || 0,
      parseFloat(x.amount)   || 0,
      String(x.abcRank || ''),
      x.refOnly ? '参考' : '提案',
      String(x.groupId   || ''),
      String(x.allocTier || ''),
      parseFloat(x.recentDemandMonthly) || 0
    ]);
    sh.getRange(2, 1, rows.length, PROPOSAL_HEADERS.length).setValues(rows);
  }

  // 発注グループ状況シート（全面書き換え。Phase J, v1.27.0〜）
  const groupStatus = p.groupStatus || [];
  let gsSh = ss.getSheetByName(SHEET_GROUP_STATUS);
  if (!gsSh) gsSh = ss.insertSheet(SHEET_GROUP_STATUS);
  gsSh.clearContents();
  gsSh.getRange(1, 1, 1, GROUP_STATUS_HEADERS.length).setValues([GROUP_STATUS_HEADERS]);
  if (groupStatus.length > 0) {
    const gsRows = groupStatus.map(x => [
      String(x.groupId      || ''),
      String(x.groupName    || ''),
      String(x.supplierCode || ''),
      String(x.supplierName || ''),
      parseFloat(x.groupUnit)   || 0,
      parseFloat(x.itemUnit)    || 0,
      parseFloat(x.itemCount)   || 0,
      parseFloat(x.meanMonthly) || 0,
      parseFloat(x.stock)       || 0,
      parseFloat(x.onOrder)     || 0,
      parseFloat(x.recommended) || 0,
      parseFloat(x.shortage)    || 0,
      parseFloat(x.threshold)   || 0,
      x.due ? 'TRUE' : 'FALSE',
      parseFloat(x.proposedQty) || 0,
      parseFloat(x.lots)        || 0,
      parseFloat(x.coverDays)   || 0,
      x.daysUntilDue === null || x.daysUntilDue === undefined ? '' : parseFloat(x.daysUntilDue),
      analyzedAt
    ]);
    gsSh.getRange(2, 1, gsRows.length, GROUP_STATUS_HEADERS.length).setValues(gsRows);
  }

  // 過剰在庫シート（全面書き換え）
  let exSh = ss.getSheetByName(SHEET_EXCESS);
  if (!exSh) exSh = ss.insertSheet(SHEET_EXCESS);
  exSh.clearContents();
  exSh.getRange(1, 1, 1, EXCESS_HEADERS.length).setValues([EXCESS_HEADERS]);
  if (excess.length > 0) {
    const exRows = excess.map(x => [
      String(x.code         || '').trim(),
      String(x.name         || ''),
      String(x.supplierCode || ''),
      String(x.supplierName || ''),
      parseFloat(x.stock)         || 0,
      parseFloat(x.recommended)   || 0,
      parseFloat(x.excessQty)     || 0,
      parseFloat(x.unitCost)      || 0,
      parseFloat(x.excessAmount)  || 0,
      parseFloat(x.monthsOfStock) || 0,
      String(x.abcRank || ''),
      String(x.pattern || ''),
      analyzedAt
    ]);
    exSh.getRange(2, 1, exRows.length, EXCESS_HEADERS.length).setValues(exRows);
  }

  // 死蔵在庫シート（全面書き換え。Phase E, v1.10.0〜）
  let deadSh = ss.getSheetByName(SHEET_DEAD);
  if (!deadSh) deadSh = ss.insertSheet(SHEET_DEAD);
  deadSh.clearContents();
  deadSh.getRange(1, 1, 1, DEAD_HEADERS.length).setValues([DEAD_HEADERS]);
  if (dead.length > 0) {
    const deadRows = dead.map(x => [
      String(x.code         || '').trim(),
      String(x.name         || ''),
      String(x.supplierCode || ''),
      String(x.supplierName || ''),
      parseFloat(x.stock)     || 0,
      parseFloat(x.unitCost)  || 0,
      parseFloat(x.deadAmount) || 0,
      String(x.lastSaleDate || ''),
      x.monthsSinceLastSale === null || x.monthsSinceLastSale === undefined ? '' : parseFloat(x.monthsSinceLastSale),
      String(x.tier   || ''),
      String(x.reason || ''),
      analyzedAt
    ]);
    deadSh.getRange(2, 1, deadRows.length, DEAD_HEADERS.length).setValues(deadRows);
  }

  // 在庫KPI履歴（実行日単位でappend。同日の再実行はその行を上書き）
  upsertKpiHistory(kpi);

  // LINE WORKS通知（LINEWORKS_WEBHOOK 未設定なら何もしない）
  // 参考表示（refOnly）は発注提案そのものではないため、件数・金額とも通知には含めない（v1.26.0）
  const realProposals = proposals.filter(x => !x.refOnly);
  const refCount = proposals.length - realProposals.length;
  if (realProposals.length > 0) {
    const bySupplier = {};
    realProposals.forEach(x => {
      const key = String(x.supplierName || '不明');
      bySupplier[key] = (bySupplier[key] || 0) + 1;
    });
    const lines = Object.keys(bySupplier)
      .sort((a, b) => bySupplier[b] - bySupplier[a])
      .slice(0, 8)
      .map(name => '・' + name + ': ' + bySupplier[name] + '件');
    const more = Object.keys(bySupplier).length > 8 ? '\n…ほか' + (Object.keys(bySupplier).length - 8) + '社' : '';
    const totalAmount = realProposals.reduce((s, x) => s + (parseFloat(x.amount) || 0), 0);
    const excessAmount = excess.reduce((s, x) => s + (parseFloat(x.excessAmount) || 0), 0);
    const excessLine = excess.length > 0
      ? ('\n⚠️ 持ちすぎ在庫: ' + excess.length + '件・過剰' + formatManYen(excessAmount))
      : '';
    const deadAmount = dead.reduce((s, x) => s + (parseFloat(x.deadAmount) || 0), 0);
    const deadLine = dead.length > 0
      ? ('\n🧹 死蔵在庫: ' + dead.length + '件・' + formatManYen(deadAmount))
      : '';
    const refLine = refCount > 0
      ? ('\n📎 参考: 提案対象外だが在庫が推奨を下回った商品 ' + refCount + '件（提案タブにグレー表示）')
      : '';
    // まとめ発注グループ（Phase J）: 発注時期になったグループだけ通知に出す。
    // まだ時期でないグループは参考表示扱いなので通知しない（件数・金額にも入らない）
    const dueGroups = groupStatus.filter(x => x.due);
    const groupLine = dueGroups.length > 0
      ? ('\n📦 まとめ発注: ' + dueGroups.map(x =>
          x.groupName + ' ' + (parseFloat(x.proposedQty) || 0) + '本（' +
          (parseFloat(x.groupUnit) || 0) + '本単位）').join(' / '))
      : '';
    // 提案滞留（Phase M, v1.31.0）: 何度も分析が回っているのに手つかずの提案を知らせる。
    // 同じ提案が出続けていても件数だけ見ると「今日もN件」としか分からないため、
    // 「減っていない」ことに気づけるよう滞留件数・欠品中件数を別枠で出す
    const stuckLong = stuckEntries.filter(e => e.stuckDays >= PROP_STUCK_NOTIFY_DAYS);
    const stuckOut = stuckLong.filter(e => e.stock <= 0).length;
    const stuckLine = stuckLong.length > 0
      ? ('\n⏳ ' + PROP_STUCK_NOTIFY_DAYS + '日以上未処理の提案: ' + stuckLong.length + '件（うち欠品中 ' + stuckOut + '件）')
      : '';
    notifyLineWorks(
      '📋 発注提案の分析が完了しました（' + analyzedAt + '）\n' +
      '提案: ' + realProposals.length + '件・合計' + formatManYen(totalAmount) + '\n' + lines.join('\n') + more +
      excessLine + deadLine + groupLine + refLine + stuckLine + '\n' +
      '発注アプリの「発注提案」タブで確認してください。'
    );
  }

  Logger.log('✅ 発注提案更新完了: ' + realProposals.length + '件（参考 ' + refCount + '件）/ 過剰在庫 '
             + excess.length + '件 / 死蔵在庫 ' + dead.length + '件 / 発注グループ '
             + groupStatus.length + '件（うち発注時期 ' + groupStatus.filter(x => x.due).length + '件）');
  return { success: true, count: realProposals.length, refCount: refCount,
           excessCount: excess.length, deadCount: dead.length,
           groupCount: groupStatus.length };
}

// POST(APIキー): 発注×仕入の突合結果を書き込む（Phase G, v1.25.0〜）
// analyze_demand.py が毎日、仕入データ明細表.CSV と発注明細を突き合わせた結果を送る。
// 「入荷実績（自動）」「計上ラグ（自動）」の2シートを全面書き換えする。
//
// リクエスト: {
//   matches: [{orderNo, code, status:'入荷済み'|'打ち切り', orderDate, qty}],
//   postingLags: {仕入先コード: ラグ日数},
//   analyzedAt
// }
//
// ⚠️ matches は「入荷待ちから除外すべき発注明細行」のみを送ること（未入荷の行は送らない）。
//   このシートは getReceivedKeys(true) の除外リストとしてそのまま使われる
function updateReceiptMatches(p) {
  const matches     = p.matches || [];
  const postingLags = p.postingLags || {};
  const analyzedAt  = String(p.analyzedAt || '');

  const ss = getSS();

  // ---- 入荷実績（自動）: 全面書き換え ----
  let sh = ss.getSheetByName(SHEET_RECEIPT_AUTO);
  if (!sh) sh = ss.insertSheet(SHEET_RECEIPT_AUTO);
  sh.clearContents();
  sh.getRange(1, 1, 1, RECEIPT_AUTO_HEADERS.length).setValues([RECEIPT_AUTO_HEADERS]);
  if (matches.length > 0) {
    const rows = matches.map(x => [
      String(x.orderNo   || '').trim(),
      String(x.code      || '').trim(),
      String(x.status    || ''),
      String(x.orderDate || ''),
      parseFloat(x.qty) || 0,
      analyzedAt
    ]);
    sh.getRange(2, 1, rows.length, RECEIPT_AUTO_HEADERS.length).setValues(rows);
  }

  // ---- 計上ラグ（自動）: 全面書き換え ----
  let lagSh = ss.getSheetByName(SHEET_POSTING_LAG);
  if (!lagSh) lagSh = ss.insertSheet(SHEET_POSTING_LAG);
  lagSh.clearContents();
  lagSh.getRange(1, 1, 1, POSTING_LAG_HEADERS.length).setValues([POSTING_LAG_HEADERS]);
  const lagCodes = Object.keys(postingLags);
  if (lagCodes.length > 0) {
    const lagRows = lagCodes.map(code => [
      String(code),
      parseFloat(postingLags[code]) || 0,
      analyzedAt
    ]);
    lagSh.getRange(2, 1, lagRows.length, POSTING_LAG_HEADERS.length).setValues(lagRows);
  }

  Logger.log('✅ 入荷突合結果を更新: 除外対象' + matches.length + '行 / 計上ラグ' + lagCodes.length + '社');
  return { success: true, count: matches.length, lagCount: lagCodes.length };
}

// 在庫KPI履歴シートへの記録（日付が一致する行があれば上書き、無ければ追記）
// kpi: {date,stockValue,monthlyCogs,turnoverDays,excessAmount,excessCount,proposalAmount,proposalCount,holdingCostAnnual,deadAmount,deadCount}
function upsertKpiHistory(kpi) {
  if (!kpi) return;
  const ss = getSS();
  let sh = ss.getSheetByName(SHEET_KPI_HISTORY);
  if (!sh) sh = ss.insertSheet(SHEET_KPI_HISTORY);
  // ヘッダーは毎回上書き（v1.10.0で列追加した際、既存シートのヘッダーが古いまま残るのを防ぐ）
  sh.getRange(1, 1, 1, KPI_HEADERS.length).setValues([KPI_HEADERS]);
  const dateStr = String(kpi.date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'));
  const row = [
    dateStr,
    parseFloat(kpi.stockValue)        || 0,
    parseFloat(kpi.monthlyCogs)       || 0,
    parseFloat(kpi.turnoverDays)      || 0,
    parseFloat(kpi.excessAmount)      || 0,
    parseFloat(kpi.excessCount)       || 0,
    parseFloat(kpi.proposalAmount)    || 0,
    parseFloat(kpi.proposalCount)     || 0,
    parseFloat(kpi.holdingCostAnnual) || 0,
    parseFloat(kpi.deadAmount)        || 0,
    parseFloat(kpi.deadCount)         || 0
  ];

  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    const dates = sh.getRange(2, 1, lastRow - 1, 1).getValues().map(r => cellToStr(r[0], 'yyyy-MM-dd'));
    const idx = dates.indexOf(dateStr);
    if (idx !== -1) {
      sh.getRange(idx + 2, 1, 1, KPI_HEADERS.length).setValues([row]);
      return;
    }
  }
  sh.appendRow(row);
}

// POST(APIキー): AI説明の追記（Claude Code から実行）
// リクエスト: { explanations: [{code, text}] }
function updateProposalExplanations(p) {
  const explanations = p.explanations || [];
  if (explanations.length === 0) return { success: true, count: 0 };

  const sh = getSS().getSheetByName(SHEET_PROPOSALS);
  if (!sh || sh.getLastRow() < 2) return { success: false, error: '発注提案シートが空です' };

  const data = sh.getDataRange().getValues();
  const colAi = PROPOSAL_HEADERS.indexOf('AI説明');  // 0-based
  const map = {};
  explanations.forEach(x => { map[String(x.code || '').trim()] = String(x.text || ''); });

  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][0]).trim();
    if (map[code] !== undefined) {
      sh.getRange(i + 1, colAi + 1).setValue(map[code]);
      count++;
    }
  }
  return { success: true, count };
}

// GET(セッション): 発注提案の取得（アプリの発注提案タブ用）
function getOrderProposals() {
  const ss = getSS();
  const sh = ss.getSheetByName(SHEET_PROPOSALS);
  let proposals = [];
  let analyzedAt = '';
  if (sh && sh.getLastRow() > 1) {
    const data = sh.getDataRange().getValues();
    proposals = data.slice(1)
      .filter(r => String(r[0]).trim() !== '')
      .map(r => ({
        code:         String(r[0]).trim(),
        name:         String(r[1] || ''),
        supplierCode: String(r[2] || '').trim(),
        supplierName: String(r[3] || ''),
        pattern:      String(r[4] || ''),
        stock:        parseFloat(r[5])  || 0,
        onOrder:      parseFloat(r[6])  || 0,
        recommended:  parseFloat(r[7])  || 0,
        proposedQty:  parseFloat(r[8])  || 0,
        lot:          parseFloat(r[9])  || 1,
        meanMonthly:  parseFloat(r[10]) || 0,
        p95Order:     parseFloat(r[11]) || 0,
        maxOrder:     parseFloat(r[12]) || 0,
        note:         String(r[13] || ''),
        aiNote:       String(r[14] || ''),
        unitCost:     parseFloat(r[16]) || 0,
        amount:       parseFloat(r[17]) || 0,
        abcRank:      String(r[18] || ''),
        // 区分列（v1.26.0〜）。列が無い旧データは空文字→false＝通常の提案として扱う
        refOnly:      String(r[19] || '').trim() === '参考',
        // まとめ発注グループ（v1.27.0〜）。列が無い旧データは空文字＝グループ外の通常商品
        groupId:      String(r[20] || '').trim(),
        allocTier:    String(r[21] || '').trim(),
        // 直近90日実需要（v1.31.0〜）。列が無い旧データは0（アプリの「要対応」判定で使う）
        recentDemandMonthly: parseFloat(r[22]) || 0
      }));
    analyzedAt = cellToStr(data[1][15], 'yyyy-MM-dd HH:mm');
  }

  // 提案滞留（Phase M, v1.31.0）: 「いつから不足しているか」を各提案行にマージする。
  // 発注提案シートには持たせず別シートで追跡しているため、ここで突き合わせる
  const stuckByCode = {};
  const stuckSh = ss.getSheetByName(SHEET_PROP_STUCK);
  if (stuckSh && stuckSh.getLastRow() > 1) {
    const todayIso = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    stuckSh.getDataRange().getValues().slice(1).forEach(r => {
      const code = String(r[0] || '').trim();
      const shortSince = String(r[2] || '');
      if (!code || !shortSince) return;
      stuckByCode[code] = {
        shortSince: shortSince,
        stuckDays: Math.max(0, daysBetweenIso_(shortSince, todayIso)),
      };
    });
  }
  proposals.forEach(x => {
    const s = stuckByCode[x.code];
    x.shortSince = s ? s.shortSince : '';
    x.stuckDays  = s ? s.stuckDays  : 0;
  });

  // 除外リスト（除外管理UIでの表示・解除用）
  let exclusions = [];
  const exSh = ss.getSheetByName(SHEET_PROPOSAL_EXCL);
  if (exSh && exSh.getLastRow() > 1) {
    exclusions = exSh.getDataRange().getValues().slice(1)
      .filter(r => String(r[0]).trim() !== '')
      .map(r => ({
        code:   String(r[0]).trim(),
        name:   String(r[1] || ''),
        reason: String(r[2] || ''),
        addedBy: String(r[3] || ''),
        addedAt: cellToStr(r[4], 'yyyy-MM-dd HH:mm')
      }));
  }

  // 終売商品リスト（終売管理UIでの表示・解除用。除外リストとは別枠）
  let eol = [];
  const eolSh2 = ss.getSheetByName(SHEET_EOL);
  if (eolSh2 && eolSh2.getLastRow() > 1) {
    eol = eolSh2.getDataRange().getValues().slice(1)
      .filter(r => String(r[0]).trim() !== '')
      .map(r => ({
        code:   String(r[0]).trim(),
        name:   String(r[1] || ''),
        reason: String(r[2] || ''),
        addedBy: String(r[3] || ''),
        addedAt: cellToStr(r[4], 'yyyy-MM-dd HH:mm')
      }));
  }

  // 過剰在庫の確認済みリスト（先に読み、確認済みの商品を過剰在庫リストから除外する）
  let excessAcks = [];
  const ackSh = ss.getSheetByName(SHEET_EXCESS_ACK);
  if (ackSh && ackSh.getLastRow() > 1) {
    excessAcks = ackSh.getDataRange().getValues().slice(1)
      .filter(r => String(r[0]).trim() !== '')
      .map(r => ({
        code:    String(r[0]).trim(),
        name:    String(r[1] || ''),
        reason:  String(r[2] || ''),
        ackedBy: String(r[3] || ''),
        ackedAt: cellToStr(r[4], 'yyyy-MM-dd HH:mm')
      }));
  }
  const ackedCodes = new Set(excessAcks.map(x => x.code));

  // 過剰在庫リスト（確認済みは除く）
  let excess = [];
  const excSh = ss.getSheetByName(SHEET_EXCESS);
  if (excSh && excSh.getLastRow() > 1) {
    excess = excSh.getDataRange().getValues().slice(1)
      .filter(r => String(r[0]).trim() !== '')
      .map(r => ({
        code:          String(r[0]).trim(),
        name:          String(r[1] || ''),
        supplierCode:  String(r[2] || '').trim(),
        supplierName:  String(r[3] || ''),
        stock:         parseFloat(r[4]) || 0,
        recommended:   parseFloat(r[5]) || 0,
        excessQty:     parseFloat(r[6]) || 0,
        unitCost:      parseFloat(r[7]) || 0,
        excessAmount:  parseFloat(r[8]) || 0,
        monthsOfStock: parseFloat(r[9]) || 0,
        abcRank:       String(r[10] || ''),
        pattern:       String(r[11] || '')
      }))
      .filter(x => !ackedCodes.has(x.code));
  }

  // 経営KPI（直近2回分＝今週・先週）
  let kpi = null;
  const kpiSh = ss.getSheetByName(SHEET_KPI_HISTORY);
  if (kpiSh && kpiSh.getLastRow() > 1) {
    const kpiLastRow = kpiSh.getLastRow();
    const n = Math.min(2, kpiLastRow - 1);
    const kpiData = kpiSh.getRange(kpiLastRow - n + 1, 1, n, KPI_HEADERS.length).getValues();
    const kpiRows = kpiData.map(r => ({
      date:              cellToStr(r[0], 'yyyy-MM-dd'),
      stockValue:        parseFloat(r[1]) || 0,
      monthlyCogs:       parseFloat(r[2]) || 0,
      turnoverDays:      parseFloat(r[3]) || 0,
      excessAmount:      parseFloat(r[4]) || 0,
      excessCount:       parseFloat(r[5]) || 0,
      proposalAmount:    parseFloat(r[6]) || 0,
      proposalCount:     parseFloat(r[7]) || 0,
      holdingCostAnnual: parseFloat(r[8]) || 0,
      deadAmount:        parseFloat(r[9]) || 0,
      deadCount:         parseFloat(r[10]) || 0
    }));
    kpi = { current: kpiRows[kpiRows.length - 1], previous: kpiRows.length > 1 ? kpiRows[0] : null };
  }

  // 死蔵在庫の確認済みリスト（先に読み、確認済みの商品を死蔵在庫リストから除外する）
  let deadAcks = [];
  const deadAckSh = ss.getSheetByName(SHEET_DEAD_ACK);
  if (deadAckSh && deadAckSh.getLastRow() > 1) {
    deadAcks = deadAckSh.getDataRange().getValues().slice(1)
      .filter(r => String(r[0]).trim() !== '')
      .map(r => ({
        code:    String(r[0]).trim(),
        name:    String(r[1] || ''),
        reason:  String(r[2] || ''),
        ackedBy: String(r[3] || ''),
        ackedAt: cellToStr(r[4], 'yyyy-MM-dd HH:mm')
      }));
  }
  const deadAckedCodes = new Set(deadAcks.map(x => x.code));

  // 死蔵在庫リスト（確認済みは除く。Phase E, v1.10.0〜）
  let dead = [];
  const deadSh2 = ss.getSheetByName(SHEET_DEAD);
  if (deadSh2 && deadSh2.getLastRow() > 1) {
    dead = deadSh2.getDataRange().getValues().slice(1)
      .filter(r => String(r[0]).trim() !== '')
      .map(r => ({
        code:                 String(r[0]).trim(),
        name:                 String(r[1] || ''),
        supplierCode:         String(r[2] || '').trim(),
        supplierName:         String(r[3] || ''),
        stock:                parseFloat(r[4]) || 0,
        unitCost:             parseFloat(r[5]) || 0,
        deadAmount:           parseFloat(r[6]) || 0,
        lastSaleDate:         cellToStr(r[7], 'yyyy-MM-dd'),
        monthsSinceLastSale:  r[8] === '' ? null : (parseFloat(r[8]) || 0),
        tier:                 String(r[9] || ''),
        reason:               String(r[10] || '')
      }))
      .filter(x => !deadAckedCodes.has(x.code));
  }

  // 手動ロット設定リスト（ロット管理UIでの表示・解除用）
  let lotOverrides = [];
  const lotSh2 = ss.getSheetByName(SHEET_LOT_OVERRIDE);
  if (lotSh2 && lotSh2.getLastRow() > 1) {
    lotOverrides = lotSh2.getDataRange().getValues().slice(1)
      .filter(r => String(r[0]).trim() !== '')
      .map(r => ({
        code:    String(r[0]).trim(),
        name:    String(r[1] || ''),
        lot:     parseInt(r[2], 10) || 0,
        addedBy: String(r[3] || ''),
        addedAt: cellToStr(r[4], 'yyyy-MM-dd HH:mm')
      }));
  }

  // まとめ発注グループの状況（Phase J, v1.27.0〜。analyze_demand.py が毎回全面書き換え）
  let groupStatus = [];
  const gsSh2 = ss.getSheetByName(SHEET_GROUP_STATUS);
  if (gsSh2 && gsSh2.getLastRow() > 1) {
    groupStatus = gsSh2.getDataRange().getValues().slice(1)
      .filter(r => String(r[0]).trim() !== '')
      .map(r => ({
        groupId:      String(r[0]).trim(),
        groupName:    String(r[1] || ''),
        supplierCode: String(r[2] || '').trim(),
        supplierName: String(r[3] || ''),
        groupUnit:    parseFloat(r[4])  || 0,
        itemUnit:     parseFloat(r[5])  || 6,
        itemCount:    parseFloat(r[6])  || 0,
        meanMonthly:  parseFloat(r[7])  || 0,
        stock:        parseFloat(r[8])  || 0,
        onOrder:      parseFloat(r[9])  || 0,
        recommended:  parseFloat(r[10]) || 0,
        shortage:     parseFloat(r[11]) || 0,
        threshold:    parseFloat(r[12]) || 0,
        due:          String(r[13] || '').trim().toUpperCase() === 'TRUE',
        proposedQty:  parseFloat(r[14]) || 0,
        lots:         parseFloat(r[15]) || 0,
        coverDays:    parseFloat(r[16]) || 0,
        daysUntilDue: r[17] === '' ? null : (parseFloat(r[17]) || 0)
      }));
  }

  // グループ設定（アプリ側の再配分に必要なパラメータ。シート未作成なら既定値で自動作成される）
  const orderGroups = readOrderGroups_();

  // 入荷待ちリスト（Phase F, v1.24.0〜）。GAS内でリアルタイム集計するため常に最新
  const pendingOrders = buildPendingOrders();

  return { success: true, proposals, analyzedAt, exclusions, eol, excess, excessAcks, dead, deadAcks,
           pendingOrders, kpi, lotOverrides, groupStatus, orderGroups };
}

// POST(セッション): 過剰在庫の確認済み登録/解除
// リクエスト: { action:'saveExcessAck', mode:'add'|'delete', code, name, reason }
function saveExcessAck(p, user_id) {
  const mode   = p.mode || '';
  const code   = String(p.code   || '').trim();
  const name   = String(p.name   || '').trim();
  const reason = String(p.reason || '').trim();
  if (!code) return { success: false, error: '商品コードが未指定です' };

  const ss = getSS();
  let sh = ss.getSheetByName(SHEET_EXCESS_ACK);
  if (!sh) {
    sh = ss.insertSheet(SHEET_EXCESS_ACK);
    sh.appendRow(['商品コード','商品名','理由','確認者','確認日時']);
  }
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  if (mode === 'add') {
    const data = sh.getDataRange().getValues();
    const exists = data.slice(1).some(r => String(r[0]).trim() === code);
    // 通信エラーによる再送で「既に確認済み」に来ることがある。冪等に成功扱いにする
    if (exists) return { success: true, alreadyExists: true };
    sh.appendRow([code, name, reason, user_id, now]);
    return { success: true };
  } else if (mode === 'delete') {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === code) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    // 通信エラーによる再送で「既に解除済み」に来ることがある。冪等に成功扱いにする
    return { success: true, notFound: true };
  }
  return { success: false, error: '不明なmode: ' + mode };
}

// POST(セッション): 死蔵在庫の確認済み登録/解除（Phase E, v1.10.0〜）
// リクエスト: { action:'saveDeadAck', mode:'add'|'delete', code, name, reason }
function saveDeadAck(p, user_id) {
  const mode   = p.mode || '';
  const code   = String(p.code   || '').trim();
  const name   = String(p.name   || '').trim();
  const reason = String(p.reason || '').trim();
  if (!code) return { success: false, error: '商品コードが未指定です' };

  const ss = getSS();
  let sh = ss.getSheetByName(SHEET_DEAD_ACK);
  if (!sh) {
    sh = ss.insertSheet(SHEET_DEAD_ACK);
    sh.appendRow(['商品コード','商品名','理由','確認者','確認日時']);
  }
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  if (mode === 'add') {
    const data = sh.getDataRange().getValues();
    const exists = data.slice(1).some(r => String(r[0]).trim() === code);
    // 通信エラーによる再送で「既に確認済み」に来ることがある。冪等に成功扱いにする
    if (exists) return { success: true, alreadyExists: true };
    sh.appendRow([code, name, reason, user_id, now]);
    return { success: true };
  } else if (mode === 'delete') {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === code) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    // 通信エラーによる再送で「既に解除済み」に来ることがある。冪等に成功扱いにする
    return { success: true, notFound: true };
  }
  return { success: false, error: '不明なmode: ' + mode };
}

// POST(セッション): 入荷済みの手動登録/解除（Phase F, v1.24.0〜）
// キーは 発注No + 商品コード（同じ商品を複数回発注していることがあるため商品コード単体では特定できない）
// 登録すると getReorderConfig の recentOrders / buildPendingOrders の両方から除外される
// リクエスト: { action:'saveReceived', mode:'add'|'delete', orderNo, code, name, qty }
function saveReceived(p, user_id) {
  const mode    = p.mode    || '';
  const orderNo = String(p.orderNo || '').trim();
  const code    = String(p.code    || '').trim();
  const name    = String(p.name    || '').trim();
  const qty     = parseFloat(p.qty) || 0;
  if (!orderNo || !code) return { success: false, error: '発注No・商品コードが未指定です' };

  const ss = getSS();
  let sh = ss.getSheetByName(SHEET_RECEIVED);
  if (!sh) {
    sh = ss.insertSheet(SHEET_RECEIVED);
    sh.appendRow(RECEIVED_HEADERS);
  }
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  if (mode === 'add') {
    const data = sh.getDataRange().getValues();
    const exists = data.slice(1).some(r => String(r[0]).trim() === orderNo && String(r[1]).trim() === code);
    // 通信エラーによる再送で「既に入荷済み」に来ることがある。冪等に成功扱いにする
    if (exists) return { success: true, alreadyExists: true };
    sh.appendRow([orderNo, code, name, qty, user_id, now]);
    return { success: true };
  } else if (mode === 'delete') {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === orderNo && String(data[i][1]).trim() === code) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    // 通信エラーによる再送で「既に解除済み」に来ることがある。冪等に成功扱いにする
    return { success: true, notFound: true };
  }
  return { success: false, error: '不明なmode: ' + mode };
}

// POST(セッション): 提案除外の登録/解除
// リクエスト: { action:'saveProposalExclusion', mode:'add'|'delete', code, name, reason }
function saveProposalExclusion(p, user_id) {
  const mode   = p.mode || '';
  const code   = String(p.code   || '').trim();
  const name   = String(p.name   || '').trim();
  const reason = String(p.reason || '').trim();
  if (!code) return { success: false, error: '商品コードが未指定です' };

  const ss = getSS();
  let sh = ss.getSheetByName(SHEET_PROPOSAL_EXCL);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PROPOSAL_EXCL);
    sh.appendRow(['商品コード','商品名','理由','登録者','登録日時']);
  }
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  if (mode === 'add') {
    const data = sh.getDataRange().getValues();
    const exists = data.slice(1).some(r => String(r[0]).trim() === code);
    // 通信エラーによる再送で「既に除外登録済み」に来ることがある。冪等に成功扱いにする
    if (exists) return { success: true, alreadyExists: true };
    sh.appendRow([code, name, reason, user_id, now]);
    // 現在の提案シートからも即時削除（次回分析を待たずに消す）
    const prSh = ss.getSheetByName(SHEET_PROPOSALS);
    if (prSh && prSh.getLastRow() > 1) {
      const prData = prSh.getDataRange().getValues();
      for (let i = prData.length - 1; i >= 1; i--) {
        if (String(prData[i][0]).trim() === code) prSh.deleteRow(i + 1);
      }
    }
    return { success: true };
  } else if (mode === 'delete') {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === code) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    // 通信エラーによる再送で「既に解除済み」に来ることがある。冪等に成功扱いにする
    return { success: true, notFound: true };
  }
  return { success: false, error: '不明なmode: ' + mode };
}

// POST(セッション): 終売フラグの登録/解除（在庫はあるが再発注できない商品用。除外設定とは別枠）
// リクエスト: { action:'saveEolFlag', mode:'add'|'delete', code, name, reason }
function saveEolFlag(p, user_id) {
  const mode   = p.mode   || '';
  const code   = String(p.code   || '').trim();
  const name   = String(p.name   || '').trim();
  const reason = String(p.reason || '').trim();
  if (!code) return { success: false, error: '商品コードが未指定です' };

  const ss = getSS();
  let sh = ss.getSheetByName(SHEET_EOL);
  if (!sh) {
    sh = ss.insertSheet(SHEET_EOL);
    sh.appendRow(['商品コード','商品名','理由','登録者','登録日時']);
  }
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  if (mode === 'add') {
    const data = sh.getDataRange().getValues();
    const exists = data.slice(1).some(r => String(r[0]).trim() === code);
    // 通信エラーによる再送で「既に終売登録済み」に来ることがある。冪等に成功扱いにする
    if (exists) return { success: true, alreadyExists: true };
    sh.appendRow([code, name, reason, user_id, now]);
    // 現在の提案シートからも即時削除（次回分析を待たずに消す）
    const prSh = ss.getSheetByName(SHEET_PROPOSALS);
    if (prSh && prSh.getLastRow() > 1) {
      const prData = prSh.getDataRange().getValues();
      for (let i = prData.length - 1; i >= 1; i--) {
        if (String(prData[i][0]).trim() === code) prSh.deleteRow(i + 1);
      }
    }
    return { success: true };
  } else if (mode === 'delete') {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === code) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    // 通信エラーによる再送で「既に解除済み」に来ることがある。冪等に成功扱いにする
    return { success: true, notFound: true };
  }
  return { success: false, error: '不明なmode: ' + mode };
}

// POST(セッション): 手動の最低発注数の登録/解除
// 過去発注実績からの自動推定（GCD）より優先して使う。既存設定があれば上書き（upsert）
// 登録時は「発注提案」シートの該当行があれば提案数量も即時再計算する
// （次回のanalyze_demand.py実行を待たずに画面へ反映するため。実行後は分析結果で上書きされる）
// リクエスト: { action:'saveLotOverride', mode:'add'|'delete', code, name, lot }
function saveLotOverride(p, user_id) {
  const mode = p.mode || '';
  const code = String(p.code || '').trim();
  const name = String(p.name || '').trim();
  if (!code) return { success: false, error: '商品コードが未指定です' };

  const ss = getSS();
  let sh = ss.getSheetByName(SHEET_LOT_OVERRIDE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LOT_OVERRIDE);
    sh.appendRow(['商品コード','商品名','最低発注数','登録者','登録日時']);
  }
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  if (mode === 'add') {
    const lot = parseInt(p.lot, 10);
    if (!(lot > 0)) return { success: false, error: '最低発注数は1以上の整数で指定してください' };
    const data = sh.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === code) {
        sh.getRange(i + 1, 1, 1, 5).setValues([[code, name, lot, user_id, now]]);
        found = true;
        break;
      }
    }
    if (!found) sh.appendRow([code, name, lot, user_id, now]);
    applyLotToProposal_(ss, code, lot);
    return { success: true };
  } else if (mode === 'delete') {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === code) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    // 通信エラーによる再送で「既に解除済み」に来ることがある。冪等に成功扱いにする
    return { success: true, notFound: true };
  }
  return { success: false, error: '不明なmode: ' + mode };
}

// ヘルパー: 「発注提案」シートに該当商品の行があれば、最低発注数と提案数量・提案金額を
// その場で再計算して書き込む（次回のanalyze_demand.py実行を待たずに画面へ反映するため）
//
// ⚠️ まとめ発注グループ（Phase J）に属する商品はここで再計算してはいけない。
//   提案数量は「系列合計を発注単位ぴったりにする配分結果」なので、1商品だけを
//   個別の不足数から再計算すると系列合計が発注単位からズレる。最低発注数の列だけ更新する
function applyLotToProposal_(ss, code, lot) {
  const prSh = ss.getSheetByName(SHEET_PROPOSALS);
  if (!prSh || prSh.getLastRow() <= 1) return;
  const data = prSh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== code) continue;
    if (String(data[i][20] || '').trim() !== '') {   // U列=グループID
      prSh.getRange(i + 1, 10, 1, 1).setValue(lot);  // J列=最低発注数のみ更新
      return;
    }
    const stock       = parseFloat(data[i][5])  || 0;
    const onOrder     = parseFloat(data[i][6])  || 0;
    const recommended = parseFloat(data[i][7])  || 0;
    const unitCost     = parseFloat(data[i][16]) || 0;
    const shortage     = recommended - (stock + onOrder);
    const proposedQty  = shortage > 0 ? Math.ceil(shortage / lot) * lot : 0;
    prSh.getRange(i + 1, 9, 1, 2).setValues([[proposedQty, lot]]);   // I列=提案数量 J列=最低発注数
    prSh.getRange(i + 1, 18, 1, 1).setValue(Math.round(unitCost * proposedQty)); // R列=提案金額
    return;
  }
}

// ヘルパー: LINE WORKS Incoming Webhook 通知（失敗しても本処理は継続）
// ペイロードは {body:{text}} 形式が正（LINE WORKS Incoming Webhook仕様）
// 戻り値: { sent, hasWebhook, httpStatus, responseBody } — testNotify のデバッグにも使う
function notifyLineWorks(text) {
  const result = { sent: false, hasWebhook: false, httpStatus: null, responseBody: '' };
  try {
    const webhook = _PROPS.getProperty('LINEWORKS_WEBHOOK');
    if (!webhook) return result;
    result.hasWebhook = true;
    const res = UrlFetchApp.fetch(webhook.trim(), {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ body: { text: text } }),
      muteHttpExceptions: true
    });
    result.httpStatus = res.getResponseCode();
    result.responseBody = String(res.getContentText() || '').slice(0, 300);
    result.sent = result.httpStatus >= 200 && result.httpStatus < 300;
    if (!result.sent) Logger.log('LINE WORKS通知 HTTPエラー: ' + result.httpStatus + ' ' + result.responseBody);
  } catch(e) {
    Logger.log('LINE WORKS通知エラー（無視）: ' + e);
    result.responseBody = String(e).slice(0, 300);
  }
  return result;
}

// POST(APIキー): LINE WORKS通知の疎通テスト（デバッグ用）
// レスポンスに webhook の設定有無・HTTPステータス・応答本文を返す
function testNotify() {
  const r = notifyLineWorks('🔔 発注アプリからのテスト通知です（testNotify）');
  return { success: true, ...r };
}

// ============================================================
// 初期設定関数（GASエディタから【1回だけ】手動実行）
// ============================================================
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function ensureSheet(name, headers) {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); Logger.log('シートを作成しました: ' + name); }
    if (sh.getLastRow() === 0) { sh.appendRow(headers); Logger.log('ヘッダーを設定しました: ' + name); }
    return sh;
  }

  ensureSheet(SHEET_HISTORY,  ['発注No','発注日','発注先コード','発注先名','FAX番号','担当者','品目数','出力方法','登録日時','user_id','requestId','requestHash','revisionBaseOrderNo','saveState']);
  ensureSheet(SHEET_ITEMS,    ['発注No','JANコード','Beaufieldコード','商品名','数量','単位','備考','手書きフラグ','登録日時']);
  ensureSheet(SHEET_REORDER,  ['商品コード','適正在庫','更新日時']);
  ensureSheet(SHEET_RECEIPT_AUTO, RECEIPT_AUTO_HEADERS);
  ensureSheet(SHEET_POSTING_LAG,  POSTING_LAG_HEADERS);

  const suppSh = ensureSheet(SHEET_SUPPLIERS, ['コード','名称','FAX','更新日時','発注方法']);
  if (suppSh.getLastRow() <= 1) {
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    suppSh.getRange(2, 1, 9, 4).setValues([
      ['10', 'デミコスメティクス',     '',             now],
      ['15', 'アペティート化粧品',      '052-883-1222', now],
      ['24', 'シュワルツコフ',          '',             now],
      ['48', '千代田化学',             '',             now],
      ['62', 'プレジール',             '0948-24-9801', now],
      ['67', 'ナプラ',                 '',             now],
      ['77', 'earth walk republic',   '078-200-6869', now],
      ['81', 'GO-ON',                 '078-200-6678', now],
      ['58', 'パシフィックプロダクツ', '03-5299-0435', now],
    ]);
    Logger.log('発注先マスターに初期データを登録しました');
  }

  Logger.log('✅ 初期設定完了 (Version: ' + VERSION + ')');
}

// ============================================================
// メーカー発注書テンプレート配信（Drive経由・認証付き）
// ============================================================
// 用途: ブラウザがGitHub Pages経由で公開PDFをfetchするのを廃止し、
//      Driveに保管したテンプレートをセッション認証経由で配信する。
//
// 前提:
//   - Driveに「Beaufield発注書テンプレート」フォルダを作成し、
//     スクリプトプロパティ ORDER_TEMPLATE_FOLDER_ID にIDを設定
//   - フォルダ内に ORDER_TEMPLATES で定義した名前の PDF をアップロード
//   - GASの実行ユーザー設定が「自分」になっていれば、利用者側のGoogle権限は不要
//
// レスポンス:
//   { success: true, filename, mimeType: 'application/pdf', data: <base64> }
//   または { success: false, error }
// ============================================================
function getOrderTemplate(makerKey) {
  const fileName = ORDER_TEMPLATES[makerKey];
  if (!fileName) {
    return { success: false, error: '不明なメーカーキー: ' + makerKey };
  }

  let folderId = _PROPS.getProperty('ORDER_TEMPLATE_FOLDER_ID');
  if (!folderId) {
    return { success: false, error: 'ORDER_TEMPLATE_FOLDER_ID 未設定（GASスクリプトプロパティ）' };
  }
  // URLが設定されていた場合もIDを正しく抽出
  const m = folderId.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m) folderId = m[1];

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files  = folder.getFilesByName(fileName);

    if (!files.hasNext()) {
      return { success: false, error: 'テンプレートPDFが見つかりません: ' + fileName };
    }

    const file   = files.next();
    const blob   = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());

    return {
      success: true,
      filename: fileName,
      mimeType: 'application/pdf',
      data: base64
    };
  } catch(err) {
    Logger.log('getOrderTemplate error: ' + err);
    return { success: false, error: 'INTERNAL_ERROR' };
  }
}
