// エディタから手動実行する読み取り計測。Web API・トリガーには登録しない。
// 計測中はgetSS経由の書き込みを拒否し、商品情報や認証情報をログに出さない。
function diagnoseOrderReadPerformance() {
  const originalGetSS = getSS;
  const originalCellToStr = cellToStr;
  let stages = {};
  function measured(label, fn) {
    const started = Date.now();
    try { return fn(); }
    finally {
      const stat = stages[label] || (stages[label] = { calls: 0, ms: 0 });
      stat.calls++;
      stat.ms += Date.now() - started;
    }
  }
  function rangeView(range, name) {
    return {
      getValues: function() { return measured(name + '.getValues', function() { return range.getValues(); }); },
      getDisplayValues: function() { return measured(name + '.getDisplayValues', function() { return range.getDisplayValues(); }); },
      setValues: function() { throw new Error('DIAGNOSTIC_WRITE_BLOCKED'); },
      setValue: function() { throw new Error('DIAGNOSTIC_WRITE_BLOCKED'); }
    };
  }
  function sheetView(sh, name) {
    if (!sh) return null;
    return {
      getLastRow: function() { return measured(name + '.getLastRow', function() { return sh.getLastRow(); }); },
      getLastColumn: function() { return measured(name + '.getLastColumn', function() { return sh.getLastColumn(); }); },
      getDataRange: function() { return rangeView(sh.getDataRange(), name); },
      getRange: function() { return rangeView(sh.getRange.apply(sh, arguments), name); },
      appendRow: function() { throw new Error('DIAGNOSTIC_WRITE_BLOCKED'); },
      clearContents: function() { throw new Error('DIAGNOSTIC_WRITE_BLOCKED'); }
    };
  }
  try {
    getSS = function() {
      const ss = measured('openSpreadsheet', originalGetSS);
      return {
        getSheetByName: function(name) { return sheetView(ss.getSheetByName(name), name); },
        insertSheet: function() { throw new Error('DIAGNOSTIC_WRITE_BLOCKED'); }
      };
    };
    cellToStr = function(val, fmt) {
      return measured('cellToStr', function() { return originalCellToStr(val, fmt); });
    };
    [getMasters, getProductMaster, getOrderProposals].forEach(function(fn) {
      stages = {};
      const started = Date.now();
      let success = false;
      let count = 0;
      try {
        const result = fn();
        success = result.success === true;
        count = (result.suppliers || result.products || result.proposals || []).length;
      } catch(e) {
        // 例外本文は業務値を含む可能性があるため記録しない。
        success = false;
      }
      console.log(JSON.stringify({ action: fn.name, totalMs: Date.now() - started, success: success, count: count, stages: stages }));
    });
  } finally {
    getSS = originalGetSS;
    cellToStr = originalCellToStr;
  }
}

function benchmarkProposalPriceCurrent() { benchmarkProposalPrice_(false); }
function benchmarkProposalPriceBatch() { benchmarkProposalPrice_(true); }

// 別実行で交互に測る。返却データの一致はハッシュだけを記録し、単価は出力しない。
function benchmarkProposalPrice_(batch) {
  const sh = getSS().getSheetByName(SHEET_PRODUCTS);
  const count = sh.getLastRow() - 1;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(v) { return String(v).trim(); });
  const code = findColIdxGAS(headers, 'コード');
  const cost = findColIdxGAS(headers, '仕入単価');
  if (code < 0 || cost < 0 || count < 1) throw new Error('DIAGNOSTIC_COLUMNS_MISSING');
  const started = Date.now();
  let pairs;
  if (batch) {
    const first = Math.min(code, cost);
    const values = sh.getRange(2, first + 1, count, Math.abs(code - cost) + 1).getValues();
    pairs = values.map(function(r) { return [r[code - first], r[cost - first]]; });
  } else {
    const codes = sh.getRange(2, code + 1, count, 1).getValues();
    const costs = sh.getRange(2, cost + 1, count, 1).getValues();
    pairs = codes.map(function(r, i) { return [r[0], costs[i][0]]; });
  }
  const readMs = Date.now() - started;
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(pairs)));
  console.log(JSON.stringify({ mode: batch ? 'batch' : 'current', rows: count, columns: batch ? Math.abs(code - cost) + 1 : 2, readMs: readMs, hash: hash }));
}

function benchmarkProductDateCurrent() { benchmarkProductDate_(false); }
function benchmarkProductDateMemo() { benchmarkProductDate_(true); }

function benchmarkProductDate_(memo) {
  const original = cellToStr;
  const cache = new Map();
  let formatCalls = 0;
  let formatMs = 0;
  cellToStr = function(value, fmt) {
    const cacheable = memo && value instanceof Date && fmt === 'yyyy/MM/dd';
    const key = cacheable ? value.getTime() : null;
    if (cacheable && cache.has(key)) return cache.get(key);
    const start = Date.now();
    const text = original(value, fmt);
    formatCalls++;
    formatMs += Date.now() - start;
    if (cacheable) cache.set(key, text);
    return text;
  };
  try {
    const start = Date.now();
    const data = getProductMaster();
    const totalMs = Date.now() - start;
    const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(data)));
    console.log(JSON.stringify({ mode: memo ? 'memo' : 'current', totalMs: totalMs, formatCalls: formatCalls, formatMs: formatMs, products: data.products.length, hash: hash }));
  } finally { cellToStr = original; }
}
