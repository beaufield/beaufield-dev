/* 棚卸しアプリ core.js — 画面に依存しない計算（仕様書 v1.0 §9.1・§10）
   ブラウザでは window.StCore、Node（単体試験）では module.exports で使う。
   ここに置くのは「同じ入力なら同じ答え」になる関数だけ（通信・保存・画面の操作はしない）。 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.StCore = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const STOCK = '@STOCK';

  // ---------- JAN（§10.2） ----------
  const digits = s => String(s == null ? '' : s).replace(/[^0-9]/g, '');

  // 照合キー: 13桁はそのまま、12桁(UPC-A)は先頭に0、8桁(EAN-8)はそのまま、それ以外は ''
  function janKey(s) {
    const d = digits(s);
    if (d.length === 13 || d.length === 8) return d;
    if (d.length === 12) return '0' + d;
    return '';
  }

  // チェックデジットの検査（誤読対策）。8/12/13桁のみ
  function eanCheckOk(s) {
    const d = digits(s);
    if (![8, 12, 13].includes(d.length)) return false;
    const k = d.length === 12 ? '0' + d : d;
    const body = k.slice(0, -1);
    let sum = 0;
    for (let i = 0; i < body.length; i++) {
      // 右端（チェックデジットの隣）から 3,1,3,1… の重み
      const w = (body.length - i) % 2 === 1 ? 3 : 1;
      sum += Number(body[i]) * w;
    }
    return (10 - (sum % 10)) % 10 === Number(k[k.length - 1]);
  }

  // ---------- 検索の表記ゆれ（§8.8） ----------
  // NFKC（半角カナ→全角、全角英数→半角）→ カタカナをひらがなに → 小文字 → 空白除去
  function normSearch(s) {
    return String(s == null ? '' : s)
      .normalize('NFKC')
      .replace(/[ァ-ヶ]/g, c => String.fromCharCode(c.charCodeAt(0) - 0x60))
      .toLowerCase()
      .replace(/\s+/g, '');
  }

  // ---------- 商品の写し ----------
  // st_get_items の配列 [code,jan,janKey,name,kana,unit,shelf,tier,discontinued,bookQty,inScope,price] を扱いやすい形に
  function buildItems(rows) {
    const list = [];
    const byCode = new Map();
    const byJan = new Map();
    rows.forEach((r, i) => {
      const it = {
        code: r[0], jan: r[1], janKey: r[2], name: r[3], kana: r[4], unit: r[5], shelf: r[6], tier: r[7],
        discontinued: !!r[8], book: r[9] == null ? null : Number(r[9]), inScope: !!r[10],
        price: r[11] == null ? null : Number(r[11]), order: i,
      };
      it.searchKey = normSearch(it.name + ' ' + it.kana + ' ' + it.code + ' ' + it.jan);
      list.push(it);
      byCode.set(it.code, it);
      if (it.janKey) {
        if (!byJan.has(it.janKey)) byJan.set(it.janKey, []);
        byJan.get(it.janKey).push(it);
      }
    });
    return { list, byCode, byJan };
  }

  // 読み取った（入力した）JANの照合（§10.3）。結果: {kind:'one'|'many'|'outside'|'none', items}
  function lookupJan(index, raw) {
    const k = janKey(raw);
    const hits = k ? (index.byJan.get(k) || []) : [];
    const inScope = hits.filter(it => it.inScope);
    if (inScope.length === 1) return { kind: 'one', items: inScope };
    if (inScope.length > 1) return { kind: 'many', items: inScope };
    if (hits.length > 0) return { kind: 'outside', items: hits };
    return { kind: 'none', items: [] };
  }

  // 検索（§8.8）。スペース区切りの全語を含む商品。対象商品を先に、最大 limit 件
  function searchItems(index, query, limit) {
    const terms = String(query || '').split(/[\s　]+/).map(normSearch).filter(Boolean);
    if (!terms.length) return [];
    const hits = index.list.filter(it => terms.every(t => it.searchKey.includes(t)));
    hits.sort((a, b) => (b.inScope - a.inScope) || (a.order - b.order));
    return hits.slice(0, limit || 50);
  }

  // ---------- 記録の重ね合わせ（§9.1） ----------
  // st_get_state の配列を記録オブジェクトに
  function entryFromRow(r) {
    return {
      id: r[0], seq: Number(r[1]), kind: r[2], code: r[3], janRaw: r[4], place: r[5], mode: r[6], qty: Number(r[7]),
      inputMethod: r[8], userId: r[9], clientAt: r[10], voided: !!r[11], voidReason: r[12], voidedBy: r[13],
      stateSeq: Number(r[14]), memo: r[15] || '',
    };
  }

  // server: Map(entryId → 記録。取消済みも含むサーバーの最新状態)
  // ops: 端末の未送信操作（保存順）。{type:'put', entry} / {type:'void', entryId}
  // 戻り値: active(有効な記録), byKey('code|place' → {qty,n,pending}), byCode(code → {total,n,places,pending})
  function computeView(server, ops) {
    const active = new Map();
    for (const e of server.values()) if (!e.voided) active.set(e.id, Object.assign({}, e, { pending: false }));
    const localVictims = new Map();  // 未送信の置き換え → それが消した記録
    for (const op of ops || []) {
      if (op.type === 'put') {
        const e = op.entry;
        if (server.has(e.id)) continue;  // 送信済みでサーバーに反映されている
        if (e.kind === 'item' && e.mode === 'replace') {
          const victims = [];
          for (const [id, x] of active) {
            if (x.kind === 'item' && x.code === e.code && x.place === e.place) { victims.push(x); active.delete(id); }
          }
          localVictims.set(e.id, victims);
        }
        active.set(e.id, Object.assign({}, e, { pending: true }));
      } else if (op.type === 'void') {
        const x = active.get(op.entryId);
        if (!x) continue;
        active.delete(op.entryId);
        if (x.mode === 'replace') {
          const victims = localVictims.has(x.id) ? localVictims.get(x.id)
            : [...server.values()].filter(v => v.voided && v.voidedBy === x.id && v.voidReason === 'replaced');
          for (const v of victims) active.set(v.id, Object.assign({}, v, { voided: false, pending: true }));
        }
      }
    }
    const byKey = new Map();
    const byCode = new Map();
    for (const e of active.values()) {
      if (e.kind !== 'item') continue;
      const k = e.code + '|' + e.place;
      const a = byKey.get(k) || { qty: 0, n: 0, pending: false };
      a.qty += e.qty; a.n += 1; a.pending = a.pending || e.pending;
      byKey.set(k, a);
      const c = byCode.get(e.code) || { total: 0, n: 0, places: {}, pending: false };
      c.total += e.qty; c.n += 1; c.places[e.place] = (c.places[e.place] || 0) + e.qty; c.pending = c.pending || e.pending;
      byCode.set(e.code, c);
    }
    return { active, byKey, byCode };
  }

  // 「済」の判定（§10.1）。棚番ありは自分の棚に記録があること、棚番なしはどこかに記録があること
  function isCounted(item, view) {
    if (item.shelf === '') return view.byCode.has(item.code);
    return view.byKey.has(item.code + '|' + item.shelf);
  }

  // 棚ごとの進捗 {total, counted}
  function shelfProgress(index, view, shelf) {
    let total = 0, counted = 0;
    for (const it of index.list) {
      if (it.shelf !== shelf || !it.inScope) continue;
      total++;
      if (isCounted(it, view)) counted++;
    }
    return { total, counted };
  }

  // 帳簿との差の確認を出すか（§10.4）
  function needsDiffCheck(book, total) {
    if (book == null || !isFinite(book) || book < 0) return false;
    if (book >= 1 && total === 0) return true;
    const d = Math.abs(total - book);
    return d >= 3 && d >= book * 0.3;
  }

  // ---------- 出力 ----------
  // CSV（UTF-8 BOM付き・CRLF・全項目をダブルクォート）（§11）
  function toCsv(rows) {
    const cell = v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
    return '﻿' + rows.map(r => r.map(cell).join(',')).join('\r\n') + '\r\n';
  }

  // ---------- 小道具 ----------
  function uuid() {
    const c = (typeof crypto !== 'undefined') ? crypto : null;
    if (c && typeof c.randomUUID === 'function') return c.randomUUID();
    const b = new Uint8Array(16);
    c.getRandomValues(b);
    b[6] = (b[6] & 0x0f) | 0x40; b[8] = (b[8] & 0x3f) | 0x80;
    const h = [...b].map(x => x.toString(16).padStart(2, '0')).join('');
    return `${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}`;
  }

  const shelfLabel = s => (s === '' ? '棚番なし' : s === STOCK ? 'ストック' : s + '棚');
  const placeLabel = p => (p === STOCK ? 'ストック' : p === '' ? '棚番なし' : p + '棚');

  // 日本時間の表示（ISO文字列 → 'M/D HH:mm'）
  function fmtJst(iso, withDate) {
    if (!iso) return '';
    const d = new Date(new Date(iso).getTime() + 9 * 3600e3);
    const hm = String(d.getUTCHours()).padStart(2, '0') + ':' + String(d.getUTCMinutes()).padStart(2, '0');
    return withDate === false ? hm : (d.getUTCMonth() + 1) + '/' + d.getUTCDate() + ' ' + hm;
  }
  function jstDate(iso) {
    const d = new Date((iso ? new Date(iso).getTime() : Date.now()) + 9 * 3600e3);
    return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
  }

  return {
    STOCK, digits, janKey, eanCheckOk, normSearch, buildItems, lookupJan, searchItems, entryFromRow,
    computeView, isCounted, shelfProgress, needsDiffCheck, toCsv, uuid, shelfLabel, placeLabel, fmtJst, jstDate,
  };
});
