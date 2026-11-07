var RC_SH_PRODUCT='ProductMaster',RC_SH_SHIPPING='SerialShipping';
// 検証専用。切戻し先はこの処理が新規作成した非公開Sheetsだけ。元原票へは書かない。
function rcSheetId_(){
 if(_PROPS.getProperty('ENV')!=='test'||DB_URL!=='https://qxzdypqbqduwuwmwnbgk.supabase.co')throw Error('TEST_ENV_REQUIRED');
 const id=_PROPS.getProperty('SERIAL_DR_SHEET');if(!id||id===SX_SOURCE_ID_||id===SX_AUTH_ID_)throw Error('ISOLATED_SHEET_REQUIRED');
 const ss=SpreadsheetApp.openById(id);if(!ss.getName().startsWith('検証専用_serial_cutback_'))throw Error('ISOLATED_SHEET_MARKER');return id;
}
function rcDecode_(v){return v&&v.type==='date'?new Date(v.iso):v;}
function rcRow_(entry,kind){
 const fields=kind==='products'?['code','name','jan','maker','series','size','serialType']:['id','shipDate','productCode','productName','jan','serial','status','returnDate','cancelDate','reason','method','customer','createdAt','registeredBy','maker','series','size'];
 const old=entry.raw_source?entry.raw_source.map(rcDecode_):new Array(fields.length).fill(''),baseline=sxData_(old,kind);
 fields.forEach((key,i)=>{if(String(entry.data[key]||'')!==String(baseline[key]||''))old[i]=String(entry.data[key]||'');});
 if(old.some(v=>typeof v==='string'&&v.startsWith('=')))throw Error('FORMULA_LIKE_VALUE_REQUIRES_REVIEW');return old;
}
function rcExport_(){
 const lock=LockService.getScriptLock();lock.waitLock(15000);
 try{
  let after=0,revision=null,products=[],shipping=[];
  for(;;){const page=sxRpc_('sr_recovery_page',{p_after:after});if(revision!==null&&revision!==page.revision)throw Error('REVISION_CHANGED');revision=page.revision;
   if(!products.length)products=page.products.map(v=>rcRow_(v,'products'));shipping.push(...page.rows.map(v=>rcRow_(v,'shipping')));
   if(page.rows.length<500)break;after=page.rows[page.rows.length-1].row_id;
  }
  let id=_PROPS.getProperty('SERIAL_DR_SHEET');if(id)throw Error('RECOVERY_SHEET_ALREADY_EXISTS');if(!id){const created=SpreadsheetApp.create('検証専用_serial_cutback_'+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMdd_HHmmss'));id=created.getId();_PROPS.setProperty('SERIAL_DR_SHEET',id);}
  const ss=SpreadsheetApp.openById(rcSheetId_());ss.setSpreadsheetTimeZone('Asia/Tokyo');const headers=sxSnapshot_();const expected={products:[headers.products[0],...products],shipping:[headers.shipping[0],...shipping]};
  for(const pair of [['ProductMaster',expected.products],['SerialShipping',expected.shipping]]){let sh=ss.getSheetByName(pair[0]);if(!sh)sh=ss.insertSheet(pair[0]);sh.clearContents();if(sh.getMaxRows()<pair[1].length)sh.insertRowsAfter(sh.getMaxRows(),pair[1].length-sh.getMaxRows());sh.getRange(1,1,pair[1].length,pair[1][0].length).setValues(pair[1]);}
  SpreadsheetApp.flush();const actual={products:ss.getSheetByName('ProductMaster').getDataRange().getValues(),shipping:ss.getSheetByName('SerialShipping').getDataRange().getValues()};
  const normalized=v=>({products:v.products.map(sxRaw_),shipping:v.shipping.map(sxRaw_)});const hash=sxHash_(normalized(expected));if(hash!==sxHash_(normalized(actual)))throw Error('CUTBACK_READBACK_MISMATCH');
  _PROPS.setProperty('SERIAL_DR_OPEN','true');return {success:true,products:products.length,shipping:shipping.length,revision,all_columns_verified:true,sha256:hash};
 }finally{lock.releaseLock();}
}
function rcLegacy_(p,jwt){
 const control=sxRpc_('sr_control',{},jwt);if(!['closed','legacy'].includes(control.writer)||_PROPS.getProperty('SERIAL_DR_OPEN')!=='true')throw Error('LEGACY_WRITER_CLOSED');rcSheetId_();
 const q=Object.assign({},p);if(q.productCode!=='DRTEST-20261107'&&['registerShipping','registerReturn'].includes(q.action))throw Error('SYNTHETIC_PRODUCT_ONLY');
 if(q.action==='getProductMaster')return rc_getProductMaster_();
 if(q.action==='search'){const r=rc_searchRecords_(q);if(r.success)r.data=r.data.slice(0,100);return r;}
 if(q.action==='registerShipping')return rc_registerShipping_(q);
 if(q.action==='registerReturn')return rc_registerReturn_(q);
 throw Error('LEGACY_ACTION_NOT_ALLOWED');
}
function rcHandle_(p){try{
 const jwt=sxAdmin_(p);if(p.mode==='close'){_PROPS.setProperty('SERIAL_DR_OPEN','false');return jsonResponse({success:true,legacy_closed:true});}
 if(p.mode==='legacy')return jsonResponse(rcLegacy_(p.params||{},jwt));
 if(p.mode!=='cutback')throw Error('UNKNOWN_RECOVERY_MODE');
 const exported=rcExport_();const name='DR-CUTBACK-'+Utilities.getUuid().slice(0,8),q={action:'registerShipping',productCode:'DRTEST-20261107',productName:'検証専用_負荷復元',jan:'0000000000000',maker:'検証専用',series:'検証',size:'1',serials:[name],shipDate:Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd'),customer:'検証専用',method:'単独',registeredBy:'検証専用'};
 const first=rcLegacy_(q,jwt),again=rcLegacy_(q,jwt),returned=rcLegacy_({action:'registerReturn',productCode:q.productCode,serials:[name],returnDate:q.shipDate,kind:'返品'},jwt),reship=rcLegacy_(q,jwt);
 const read=rcLegacy_({action:'search',productCode:q.productCode,serialNo:name,statuses:[]},jwt);
 const success=first.registered===1&&again.registered===0&&again.skipped===1&&returned.updated===1&&reship.registered===1&&read.success&&read.total===2;
 return jsonResponse({success,exported,legacy_write:{registered:first.registered,retry_registered:again.registered,retry_skipped:again.skipped,returned:returned.updated,reship:reship.registered,readback_rows:read.total},source_sheets_written:false,isolated_sheet_only:true});
 }catch(e){return jsonResponse({success:false,error:e.message});}}

var RC_COL = {
  ID:          0,  // A: ID（UUID）
  SHIP_DATE:   1,  // B: 出荷日
  PROD_CODE:   2,  // C: 商品コード
  PROD_NAME:   3,  // D: 商品名
  JAN:         4,  // E: JANコード
  SERIAL:      5,  // F: シリアルNo
  STATUS:      6,  // G: 状態（出荷中/返品済/取消）
  RETURN_DATE: 7,  // H: 返品日
  CANCEL_DATE: 8,  // I: 取消日（未使用）
  REASON:      9,  // J: 取消理由（「返品」「取消」を格納）
  METHOD:      10, // K: 登録方法（単独/連番）
  CUSTOMER:    11, // L: 得意先
  CREATED_AT:  12, // M: 登録日時
  REGISTERED_BY: 13, // N: 登録者（v2.3.0で追加。登録時のログイン氏名をそのまま保存）
  MAKER:       14, // O: メーカー（v2.3.0で追加。登録時点の商品マスタ値を複製保存）
  SERIES:      15, // P: 商品シリーズ（v2.3.0で追加。同上）
  SIZE:        16  // Q: サイズ（v2.3.0で追加。同上）
};
var RC_PCOL = {
  CODE:        0, // A: 商品コード
  NAME:        1, // B: 商品名
  JAN:         2, // C: JANコード
  MAKER:       3, // D: メーカー
  SERIES:      4, // E: 商品シリーズ
  SIZE:        5, // F: サイズ
  SERIAL_TYPE: 6  // G: シリアル種別（バーコード/QR/手入力、空欄=未設定）
};
function rc_getProductMaster_() {
  try {
    var ss   = SpreadsheetApp.openById(rcSheetId_());
    var sh   = ss.getSheetByName(RC_SH_PRODUCT);
    var data = sh.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[RC_PCOL.CODE]) continue;
      result.push({
        code:       String(row[RC_PCOL.CODE]),
        name:       String(row[RC_PCOL.NAME]),
        jan:        String(row[RC_PCOL.JAN]),
        maker:      String(row[RC_PCOL.MAKER]),
        series:     String(row[RC_PCOL.SERIES]),
        size:       String(row[RC_PCOL.SIZE]),
        serialType: String(row[RC_PCOL.SERIAL_TYPE] || '')
      });
    }
    return { success: true, data: result };
  } catch (e) {
    Logger.log('rc_getProductMaster_ error: ' + e);
    return { success: false, message: '商品マスタの取得に失敗しました' };
  }
}

// ============================================================
// 出荷登録（まとめ書き込み＝高速版）
// params: productCode, productName, jan, shipDate (YYYY/MM/DD),
//         customer, method (単独|連番), serials (JSON文字列 or 配列)
// 【改善点】
//  A. appendRowの1件ずつループを廃止 → setValuesで一括書き込み（数十倍高速）
//  C. 重複チェックと追記でシートを2回開いていたのを1回の読み込みに統合
//  D. 同時登録による行の衝突を防ぐためLockServiceで排他制御
// ============================================================

function rc_registerShipping_(params) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // 他の登録処理の完了を最大15秒待つ
  } catch (e) {
    return { success: false, message: '他の処理と競合しました。少し待って再度お試しください' };
  }
  try {
    var serials = rc_parseArray_(params.serials);
    var sh      = SpreadsheetApp.openById(rcSheetId_()).getSheetByName(RC_SH_SHIPPING);
    var data    = sh.getDataRange().getValues(); // ← シート読み込みは1回だけ（重複チェックに再利用）

    // 重複（同一商品コード × 出荷中）のシリアルを集合化
    var serialSet = {};
    serials.forEach(function(s) { serialSet[s] = true; });
    var dupSet = {};
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (String(r[RC_COL.PROD_CODE]) === String(params.productCode) &&
          r[RC_COL.STATUS] === '出荷中' &&
          serialSet[String(r[RC_COL.SERIAL])]) {
        dupSet[String(r[RC_COL.SERIAL])] = true;
      }
    }

    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    var newRows = [], skipped = 0, skippedSerials = [];

    serials.forEach(function(serial) {
      if (dupSet[serial]) { skipped++; skippedSerials.push(serial); return; }
      var row = new Array(17).fill('');
      row[RC_COL.ID]           = Utilities.getUuid();
      row[RC_COL.SHIP_DATE]    = params.shipDate;
      row[RC_COL.PROD_CODE]    = params.productCode;
      row[RC_COL.PROD_NAME]    = params.productName;
      row[RC_COL.JAN]          = params.jan || '';
      row[RC_COL.SERIAL]       = serial;
      row[RC_COL.STATUS]       = '出荷中';
      row[RC_COL.METHOD]       = params.method || '単独';
      row[RC_COL.CUSTOMER]     = params.customer || '';
      row[RC_COL.CREATED_AT]   = now;
      row[RC_COL.REGISTERED_BY] = params.registeredBy || '';
      row[RC_COL.MAKER]         = params.maker  || '';
      row[RC_COL.SERIES]        = params.series || '';
      row[RC_COL.SIZE]          = params.size   || '';
      newRows.push(row);
    });

    // 追記は1回のsetValuesで一括（appendRowループの数十倍高速）
    if (newRows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 17).setValues(newRows);
    }

    return { success: true, registered: newRows.length, skipped: skipped, skippedSerials: skippedSerials };
  } catch (e) {
    Logger.log('rc_registerShipping_ error: ' + e);
    return { success: false, message: '出荷登録に失敗しました' };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 返品・取消登録
// params: productCode, returnDate (YYYY/MM/DD),
//         kind (返品|取消), serials (JSON文字列 or 配列)
// 【改善点】
//  B. 1行につきsetValueを3回呼んでいたのを、変更行だけ1回のsetValuesにまとめて高速化
//  D. 読み書きの競合を防ぐためLockServiceで排他制御
// ============================================================

function rc_registerReturn_(params) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { success: false, message: '他の処理と競合しました。少し待って再度お試しください' };
  }
  try {
    var serials   = rc_parseArray_(params.serials);
    var sh        = SpreadsheetApp.openById(rcSheetId_()).getSheetByName(RC_SH_SHIPPING);
    var data      = sh.getDataRange().getValues();
    var newStatus = (params.kind === '取消') ? '取消' : '返品済';
    var serialSet = {};
    serials.forEach(function(s) { serialSet[s] = true; });

    var updated = 0, skipped = 0, skippedSerials = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[RC_COL.PROD_CODE]) !== String(params.productCode)) continue;
      if (!serialSet[String(row[RC_COL.SERIAL])]) continue;
      if (row[RC_COL.STATUS] !== '出荷中') {
        skipped++;
        skippedSerials.push(String(row[RC_COL.SERIAL]));
        continue;
      }
      // メモリ上で更新し、変更行のG〜J列だけを1回のsetValuesで書き戻す（3回→1回）
      row[RC_COL.STATUS]      = newStatus;
      row[RC_COL.RETURN_DATE] = params.returnDate;
      row[RC_COL.REASON]      = params.kind;
      // G(STATUS)〜J(REASON)の4列を一括書き込み（間のI列CANCEL_DATEは元の値のまま維持）
      sh.getRange(i + 1, RC_COL.STATUS + 1, 1, 4).setValues([row.slice(RC_COL.STATUS, RC_COL.STATUS + 4)]);
      updated++;
    }

    return { success: true, updated: updated, skipped: skipped, skippedSerials: skippedSerials };
  } catch (e) {
    Logger.log('rc_registerReturn_ error: ' + e);
    return { success: false, message: '返品・取消登録に失敗しました' };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 検索
// params: dateFrom, dateTo, statuses (JSON文字列 or 配列),
//         productCode, serialNo, productName, maker, series
// ============================================================

function rc_searchRecords_(params) {
  try {
    var ss   = SpreadsheetApp.openById(rcSheetId_());
    var sh   = ss.getSheetByName(RC_SH_SHIPPING);
    var data = sh.getDataRange().getValues();

    // 日付はYYYY-MM-DD文字列で比較（タイムゾーン問題回避）
    var fromStr = params.dateFrom ? String(params.dateFrom).replace(/\//g, '-').substring(0, 10) : '';
    var toStr   = params.dateTo   ? String(params.dateTo).replace(/\//g, '-').substring(0, 10)   : '';

    var statuses = rc_parseArray_(params.statuses);

    // メーカー・シリーズ・サイズ絞り込み用にProductMasterをメモリ展開
    var prodMap = {};
    if (params.maker || params.series || params.size) {
      var pData = ss.getSheetByName(RC_SH_PRODUCT).getDataRange().getValues();
      for (var p = 1; p < pData.length; p++) {
        var pr = pData[p];
        prodMap[String(pr[RC_PCOL.CODE])] = {
          maker:  String(pr[RC_PCOL.MAKER]),
          series: String(pr[RC_PCOL.SERIES]),
          size:   String(pr[RC_PCOL.SIZE])
        };
      }
    }

    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // UUID がない移行データも対象とするため、シリアルNoで空行判定
      if (!row[RC_COL.SERIAL]) continue;

      var shipDateStr = '';
      if (row[RC_COL.SHIP_DATE]) {
        if (row[RC_COL.SHIP_DATE] instanceof Date) {
          // Sheetsがdate型で返した場合（日本時間で YYYY-MM-DD に変換）
          shipDateStr = Utilities.formatDate(row[RC_COL.SHIP_DATE], 'Asia/Tokyo', 'yyyy-MM-dd');
        } else {
          // 文字列で返した場合（"2025/9/19" 形式など 0埋めなしに対応）
          var parts = String(row[RC_COL.SHIP_DATE]).split(/[\/\-]/);
          if (parts.length === 3) {
            shipDateStr = parts[0].padStart(4,'0') + '-' + parts[1].padStart(2,'0') + '-' + parts[2].padStart(2,'0');
          }
        }
      }
      if (fromStr && shipDateStr && shipDateStr < fromStr) continue;
      if (toStr   && shipDateStr && shipDateStr > toStr)   continue;

      if (statuses.length > 0 && statuses.indexOf(String(row[RC_COL.STATUS])) < 0) continue;
      if (params.productCode && rc_normalizeText_(row[RC_COL.PROD_CODE]) !== rc_normalizeText_(params.productCode)) continue;
      if (params.serialNo    && rc_normalizeText_(row[RC_COL.SERIAL])    !== rc_normalizeText_(params.serialNo))    continue;
      if (params.productName && !_matchesQuery(String(row[RC_COL.PROD_NAME]), rc_normalizeText_(params.productName))) continue;

      // メーカー/シリーズ/サイズは行自体の複製値を優先し、
      // 未設定（複製列導入前の旧データ）の場合のみ商品マスタとの突き合わせにフォールバックする。
      // →商品マスタ側が後で変更・削除されても、登録当時の情報で絞り込みできる。
      var pm        = prodMap[String(row[RC_COL.PROD_CODE])];
      var effMaker  = String(row[RC_COL.MAKER]  || '') || (pm ? pm.maker  : '');
      var effSeries = String(row[RC_COL.SERIES] || '') || (pm ? pm.series : '');
      var effSize   = String(row[RC_COL.SIZE]   || '') || (pm ? pm.size   : '');

      if (params.maker  && rc_normalizeText_(effMaker)  !== rc_normalizeText_(params.maker))  continue;
      if (params.series && rc_normalizeText_(effSeries) !== rc_normalizeText_(params.series)) continue;
      if (params.size   && rc_normalizeText_(effSize)   !== rc_normalizeText_(params.size))   continue;

      result.push({
        id:           String(row[RC_COL.ID]),
        shipDate:     rc_formatDate_(row[RC_COL.SHIP_DATE]),
        productCode:  String(row[RC_COL.PROD_CODE]),
        productName:  String(row[RC_COL.PROD_NAME]),
        jan:          String(row[RC_COL.JAN]),
        serial:       String(row[RC_COL.SERIAL]),
        status:       String(row[RC_COL.STATUS]),
        returnDate:   rc_formatDate_(row[RC_COL.RETURN_DATE]),
        reason:       String(row[RC_COL.REASON]        || ''),
        method:       String(row[RC_COL.METHOD]        || ''),
        customer:     String(row[RC_COL.CUSTOMER]      || ''),
        createdAt:    String(row[RC_COL.CREATED_AT]    || ''),
        registeredBy: String(row[RC_COL.REGISTERED_BY] || ''),
        maker:        effMaker,
        series:       effSeries,
        size:         effSize
      });
    }

    result.sort(function(a, b) { return b.createdAt.localeCompare(a.createdAt); });
    return { success: true, data: result, total: result.length };
  } catch (e) {
    Logger.log('rc_searchRecords_ error: ' + e);
    return { success: false, message: '検索に失敗しました' };
  }
}

// ============================================================
// メーカー別・月単位のCSV出力
// params: maker（完全一致・商品マスタ突き合わせ含む）, yearMonth（"YYYY-MM"）
// 対象: 出荷日が yearMonth に該当し、状態が「出荷中」の行のみ
// 出力列: 出荷日/商品コード/商品名/JANコード/シリアルNo/状態/返品日/取消日/取消理由/得意先
// ============================================================

function rc_parseArray_(val) {
  if (Array.isArray(val)) return val;
  if (!val) return [];
  try { return JSON.parse(val); } catch (e) { return [String(val)]; }
}

// 半角全角・大文字小文字・スペース正規化

function rc_formatDate_(val) {
  if (!val) return '';
  try {
    var d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return String(val);
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
  } catch (e) { return String(val); }
}


function rc_shipDateYmd_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  var parts = String(val).split(/[\/\-]/);
  if (parts.length === 3) {
    return parts[0].padStart(4, '0') + '-' + parts[1].padStart(2, '0') + '-' + parts[2].padStart(2, '0');
  }
  return '';
}

// CSVフィールドのエスケープ（カンマ・改行・ダブルクォートを含む場合のみクォート）

function rc_normalizeText_(str) {
  try {
    return String(str || '').normalize('NFKC').toLowerCase().replace(/\s+/g, ' ').trim();
  } catch(e) {
    return String(str || '').toLowerCase().trim();
  }
}

// スペース区切りAND部分一致
