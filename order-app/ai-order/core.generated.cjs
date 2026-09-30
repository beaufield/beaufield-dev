'use strict';
// Generated from index.html; edit the source and rebuild.
/* AI_ORDER_CORE_BEGIN: 入力解釈は補助。数量・配分はこの決定的処理だけで行う。 */
function aiOrderNorm(value) {
  return String(value || '').normalize('NFKC').toUpperCase().replace(/[\s　]/g, '');
}
function aiOrderDemand(row) {
  const value = row.recentDemandMonthly !== undefined && row.recentDemandMonthly !== null
    ? row.recentDemandMonthly : row.meanMonthly;
  return Number(value);
}
function aiOrderParseText(value) {
  const text = String(value || '').normalize('NFKC');
  const result = { mode: '', series: '', sizes: [], eachCases: null, days: null, questions: [] };
  const seriesMatch = text.match(/(?:の|^|[\s、。])([^\s、。の]{2,30})シリーズ/) ||
    text.match(/(?:の|^|[\s、。])([^\s、。の]{2,30})の\d{3,4}(?=サイズ|ML|G)/i);
  if (seriesMatch) result.series = seriesMatch[1];
  const sizes = [...text.matchAll(/(?:^|[^\d])(\d{3,4})(?=\s*(?:ML|G|サイズ|と|、|は|で|$))/gi)]
    .map(m => m[1]);
  result.sizes = [...new Set(sizes)].map(size => {
    const packMatch = text.match(new RegExp(size + '[^、。\\n]{0,30}?(\\d+)\\s*本\\s*(?:で\\s*1\\s*ケース|[/／]\\s*ケース)'));
    const casesMatch = text.match(new RegExp(size + '\\s*(?:サイズ|ml|g)?\\s*(?:を|は|で)?\\s*(\\d+)\\s*ケース', 'i'));
    return { size, pack: packMatch ? Number(packMatch[1]) : null,
      cases: casesMatch ? Number(casesMatch[1]) : null };
  });
  if (/ケース/.test(text)) {
    result.mode = 'cases';
    const each = text.match(/(?:各(?:サイズ)?|それぞれ)\s*(\d+)\s*ケース/);
    if (each) result.eachCases = Number(each[1]);
    else if (!result.sizes.length || result.sizes.some(x => !x.cases))
      result.questions.push('ケース数が各サイズか合計かを確認し、サイズ別に入力してください。');
    if (result.sizes.length === 0) result.questions.push('対象サイズを入力してください。');
  } else if (/日分/.test(text)) {
    result.mode = 'extra';
    const range = /\d+\s*[〜~～、,－-]\s*\d+\s*日分/.test(text);
    const day = text.match(/(\d+)\s*日分/);
    if (!range && day) result.days = Number(day[1]);
    else result.questions.push('追加する日数を1つ選んでください。');
  } else {
    result.questions.push('ケース配分か、通常提案へ追加する日数を選んでください。');
  }
  return result;
}
function aiOrderAllocate(rows, targetCases, pack) {
  if (!Number.isInteger(targetCases) || targetCases < 1 || targetCases > 100 ||
      !Number.isInteger(pack) || pack < 1 || pack > 1000) {
    throw new Error('ケース数と入数は正の整数で入力してください。');
  }
  if ((rows || []).some(r => r.stock == null || r.onOrder == null || r.recommended == null ||
      !Number.isFinite(Number(r.stock)) || !Number.isFinite(Number(r.onOrder)) ||
      !Number.isFinite(aiOrderDemand(r)) || aiOrderDemand(r) < 0 ||
      !Number.isFinite(Number(r.recommended)))) {
    throw new Error('分析値に欠損があります。通常提案で在庫と需要を確認してください。');
  }
  const active = (rows || []).filter(r => !r.excluded && !r.eolFlagged &&
    aiOrderDemand(r) > 0).map(r => ({
      ...r, demand: aiOrderDemand(r),
      pos: Number(r.stock || 0) + Number(r.onOrder || 0), cases: 0
    }));
  if (!active.length) throw new Error('需要のある対象商品がありません。');
  const totalDemand = active.reduce((s, r) => s + r.demand, 0);
  const targetMonths = (active.reduce((s, r) => s + r.pos, 0) + targetCases * pack) / totalDemand;
  // 需要加重の発注後在庫月数の二乗偏差。ケースごとの限界費用が増加するため逐次最小を選ぶ。
  for (let i = 0; i < targetCases; i++) {
    let chosen = null;
    let best = Infinity;
    for (const r of active) {
      const before = r.pos + r.cases * pack - r.demand * targetMonths;
      const delta = (2 * before * pack + pack * pack) / r.demand;
      if (delta < best - 1e-9 || (Math.abs(delta - best) <= 1e-9 && chosen && String(r.code) < String(chosen.code))) {
        best = delta;
        chosen = r;
      }
    }
    chosen.cases++;
  }
  return {
    rows: active.map(r => ({
      ...r, pack, qty: r.cases * pack,
      monthsAfter: (r.pos + r.cases * pack) / r.demand
    })),
    ignored: (rows || []).length - active.length,
    allAboveRecommended: active.every(r => Number.isFinite(Number(r.recommended)) &&
      r.pos >= Number(r.recommended)),
    targetMonths
  };
}
function aiOrderExtraQty(baseQty, monthlyDemand, days, lot) {
  if (!Number.isInteger(days) || days < 1 || days > 30) throw new Error('追加日数は1〜30日の整数にしてください。');
  if (!Number.isInteger(lot) || lot < 1) throw new Error('発注単位を確認できません。');
  const base = Number(baseQty);
  const demand = Number(monthlyDemand);
  if (!Number.isInteger(base) || base < 1 || !Number.isFinite(demand) || demand < 0)
    throw new Error('通常提案または需要の値が不正です。');
  return base + Math.ceil(demand * days / 30.4 / lot) * lot;
}
// Portable snapshot calculator. Cloud/local/browser use the same quantity and blocking rules.
function aiOrderCalculateSnapshot(conditions, snapshot, nowMs = Date.now()) {
  const { mode, supplierCode } = conditions;
  const sameSupplier = value => String(value || '').replace(/^0+/, '') === String(supplierCode || '').replace(/^0+/, '');
  if (!supplierCode || !['cases', 'extra'].includes(mode)) throw new Error('計算方法と発注先を確認してください。');
  if (!snapshot || !snapshot.analysisId || !snapshot.analyzedAt || !Number.isFinite(snapshot.analyzedMs) ||
      snapshot.analyzedMs > nowMs + 5 * 60 * 1000 || !Array.isArray(snapshot.blockedCodes) ||
      !Array.isArray(snapshot.products) || !snapshot.products.length || !Array.isArray(snapshot.proposals))
    throw new Error('分析世代、時刻、商品マスター、未入荷確認のデータが不足しています。');
  const draft = { mode, supplierCode, analysisId: snapshot.analysisId, analyzedAt: snapshot.analyzedAt,
    analysisInputAt: snapshot.analysisInputAt || '', sections: [], rows: [], warnings: [] };
  const blocked = new Set(snapshot.blockedCodes);
  if (nowMs - snapshot.analyzedMs > 36 * 60 * 60 * 1000)
    draft.warnings.push('分析時点から36時間以上経過しています。現在庫と未入荷を確認してください。');
  if (mode === 'cases') {
    if (!conditions.packConfirmed) throw new Error('今回のケース入数を確認してチェックしてください。');
    if (aiOrderNorm(conditions.series).length < 2) throw new Error('系列名を2文字以上で入力してください。');
    const inputs = conditions.sizes;
    if (!Array.isArray(inputs) || !inputs.length || inputs.some(x => !/^\d{3,4}$/.test(x.size) ||
        !Number.isInteger(x.cases) || x.cases < 1 || x.cases > 100 ||
        !Number.isInteger(x.pack) || x.pack < 1 || x.pack > 1000) ||
        new Set(inputs.map(x => x.size)).size !== inputs.length)
      throw new Error('サイズ、ケース数、本/ケースを正の整数で入力してください。サイズの重複も確認してください。');
    if (!Array.isArray(snapshot.metrics)) throw new Error('商品別分析値がありません。');
    for (const input of inputs) {
      const matches = snapshot.metrics.filter(r => sameSupplier(r.supplierCode) &&
        aiOrderNorm(r.name).includes(aiOrderNorm(conditions.series)) && aiOrderNorm(r.name).includes(input.size) &&
        !r.excluded && !r.eolFlagged);
      if (!matches.length) throw new Error(input.size + 'サイズの商品を特定できません。系列名とサイズを確認してください。');
      if (matches.some(r => r.groupId)) throw new Error(input.size + 'サイズには系列単位の既存発注ルールがあります。現在は従来の系列カードで確認してください。');
      if (matches.some(r => !snapshot.products.some(m => m.code === r.code && sameSupplier(m.supplierCD) && m.discontinued !== '廃番')))
        throw new Error(input.size + 'サイズの商品マスターが不足または終売です。更新してから確認してください。');
      if (matches.some(r => blocked.has(r.code))) throw new Error(input.size + 'サイズに分析後の未入荷発注があります。二重発注を避けるため通常の履歴で確認してください。');
      const allocation = aiOrderAllocate(matches, input.cases, input.pack);
      if (allocation.allAboveRecommended) draft.warnings.push(input.size + 'は対象商品の現在庫＋未入荷が分析上の推奨在庫以上です。指定ケース数を見直してください。');
      if (allocation.ignored) draft.warnings.push(input.size + 'で需要ゼロなどの' + allocation.ignored + '品目は自動配分から外しました。');
      const rows = allocation.rows.map(r => ({ ...r, size: input.size }));
      if (rows.some(r => r.qty > 0 && (!Number.isInteger(Number(r.lot)) || Number(r.lot) < 1 || r.qty % Number(r.lot) !== 0)))
        throw new Error(input.size + 'サイズのケース入数と発注単位が一致しません。確認してください。');
      draft.sections.push({ size: input.size, targetCases: input.cases, pack: input.pack, rows });
      draft.rows.push(...rows);
    }
  } else {
    const days = conditions.days;
    if (!Number.isInteger(days) || days < 1 || days > 30) throw new Error('追加日数は1〜30日の整数で入力してください。');
    draft.days = days;
    const target = aiOrderNorm(conditions.target);
    const source = snapshot.proposals.filter(p => sameSupplier(p.supplierCode) && !p.refOnly && Number(p.proposedQty) > 0 &&
      (!target || aiOrderNorm(p.name).includes(target) || aiOrderNorm(p.code).includes(target)));
    const eligible = source.filter(p => !p.groupId && !blocked.has(p.code));
    if (!eligible.length) throw new Error('追加できる通常提案がありません。系列単位の商品は従来の系列カードで確認してください。');
    if (eligible.length < source.length) draft.warnings.push('系列単位または分析後に発注済みの商品は追加対象から外しました。');
    draft.rows = eligible.map(p => {
      if (!snapshot.products.some(m => m.code === p.code && sameSupplier(m.supplierCD) && m.discontinued !== '廃番'))
        throw new Error('対象の商品マスターが不足または終売です。');
      const lot = Number(p.lot);
      if (Number(p.proposedQty) % lot !== 0) throw new Error('通常提案と発注単位が一致しません。');
      return { ...p, baseQty: Number(p.proposedQty), qty: aiOrderExtraQty(Number(p.proposedQty), aiOrderDemand(p), days, lot),
        lot, stock: Number(p.stock || 0), onOrder: Number(p.onOrder || 0), demand: aiOrderDemand(p) };
    });
  }
  return draft;
}

module.exports = { aiOrderNorm, aiOrderAllocate, aiOrderExtraQty, aiOrderCalculateSnapshot };
