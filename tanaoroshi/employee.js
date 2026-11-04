"use strict";
// 社員利用版。既存の認証・IndexedDB・通信再送を使い、場所と操作の意図を明示する。
const EMP = {
  assignment: (place) => S.employee.assignments.find((a) => a.place === place),
  tier: (place, item) =>
    place === STOCK
      ? ""
      : cur().params.tier != null
        ? cur().params.tier
        : item
          ? item.tier
          : "",
  mode: () => cur().params.countMode || "key",
  draft: null,
  busy: false,
  draftKey: () => "tally:" + sid() + ":" + S.me.userId,
  verified: new Set(),
};
const recountStatus = { open: '再カウント待ち', counted: '管理者の確認待ち', adopted: '採用済み', dismissed: '確認済み（初回を維持）' };
const workLocationKey = (place, tier) =>
  sid() + "|" + place + "|" + tier + "|" + (EMP.assignment(place)?.epoch || 0);
async function confirmWorkLocation(place, tier) {
  const key = workLocationKey(place, tier);
  if (EMP.verified.has(key)) return true;
  if (
    !(await confirmYes(
      "棚札と場所を確認",
      C.placeLabel(C.location(place, tier)) +
        " の現物を数えます。棚札と一致していますか？",
      "一致している",
    ))
  )
    return false;
  EMP.verified.add(key);
  return true;
}
const originalSaveEntry = saveEntry;
saveEntry = (o) => {
  const tier = o.tier != null ? o.tier : EMP.tier(o.place);
  const a = EMP.assignment(o.place);
  return originalSaveEntry({
    ...o,
    tier,
    assignmentEpoch: a && a.epoch,
    expectedIds:
      o.expectedIds ||
      C.groupIds(S.view, o.code, o.place, tier, o.janRaw, o.memo),
  });
};

Object.assign(ERR_TEXT, {
  PLACE_BUSY:
    "ほかの担当者または端末が作業中です。担当者の提出を確認してください。",
  PLACE_NOT_OWNED: "この場所の担当が変わっています。記録を確認してください。",
  ASSIGNMENT_CHANGED: "担当交代前の記録です。管理者の確認が必要です。",
  PLACE_SUBMITTED: "提出済みです。再開してから入力してください。",
  COUNT_CHANGED:
    "同じ場所の記録が変更されています。最新の記録と照合してください。",
  RECOUNT_ACTIVE: "再カウントの確認中です。通常入力を止めています。",
  SAME_COUNTER: "初回を数えた人以外を再カウント担当に選んでください。",
  SUBMISSIONS_PENDING: "未提出の担当場所があります。",
  REVIEWS_PENDING: "不明品・対象範囲の確認が残っています。",
  CONFLICTS_PENDING: "重複・遅延記録の確認が残っています。",
  RECOUNTS_PENDING: "再カウントの確認が残っています。",
  PHYSICAL_CHECK_REQUIRED: "現物が残っていないことを確認してください。",
});

async function employeeAction(action, extra = {}) {
  const r = await rpc("st_employee", {
    sessionId: sid(),
    action,
    deviceId: S.deviceId,
    ...extra,
  });
  if (!r.success) {
    toast(errText(r.error), "error");
    return null;
  }
  if (r.employee) S.employee = r.employee;
  return r;
}
async function reloadEmployee() {
  const r = await rpc("st_get_items", { sessionId: sid() });
  if (!r.success) throw new Error(r.error);
  S.index = C.buildItems(r.items);
  S.shelves = parseShelves(r.shelves);
  await kvPut("items", {
    sessionId: sid(),
    role: S.me.role,
    items: r.items,
    shelves: r.shelves,
  });
  await fetchState(true);
}
async function ensureClaim(place) {
  let a = EMP.assignment(place);
  if (
    a &&
    !a.submitted &&
    a.userId === S.me.userId &&
    a.deviceId === S.deviceId
  )
    return true;
  if (S.offline || navigator.onLine === false) {
    toast("担当を開始・交代するときは通信が必要です", "error");
    return false;
  }
  if (
    !(await confirmYes(
      "この場所を担当しますか？",
      C.placeLabel(place) +
        "\n同じ現物を二重に数えないよう、この端末で担当します。",
      "担当して数える",
    ))
  )
    return false;
  const r = await employeeAction("claim", { place });
  if (!r) return false;
  await fetchState(false);
  return true;
}
function groupRows(code) {
  const groups = new Map();
  for (const e of S.view.active.values()) {
    if (e.kind !== "item" || e.code !== code) continue;
    const key = C.location(e.place, e.tier) + "|" + e.userId;
    const g = groups.get(key) || {
      place: e.place,
      tier: e.tier || "",
      userId: e.userId,
      qty: 0,
      at: "",
      pending: false,
    };
    g.qty += e.qty;
    g.at = e.clientAt > g.at ? e.clientAt : g.at;
    g.pending ||= e.pending;
    groups.set(key, g);
  }
  return [...groups.values()];
}
function recountFor(code, place, tier) {
  return S.employee.recounts.find(
    (r) =>
      r.code === code &&
      r.place === place &&
      r.tier === tier &&
      ["open", "counted"].includes(r.state),
  );
}
function sharedHtml(item, place, tier, controls = false) {
  const rows = groupRows(item.code);
  const sum = rows.reduce((n, x) => n + x.qty, 0);
  const hidden = S.employee.recounts.some(
    (r) =>
      r.code === item.code &&
      r.assignee === S.me.userId &&
      ["open", "counted"].includes(r.state),
  );
  if (hidden)
    return '<div class="warnbox">再カウント対象です。初回数量を見ずに「再カウント」画面から数えてください。</div>';
  return `<div class="card shared-records"><b>誰が・どこで登録したか</b>${
    rows.length
      ? `<div class="scrollx"><table class="rep"><thead><tr><th>場所・担当</th><th>数量</th><th>時刻</th></tr></thead><tbody>${rows
          .map(
            (g) =>
              `<tr><td>${esc(C.placeLabel(C.location(g.place, g.tier)))}<br>${esc(S.users[g.userId] || (g.userId === S.me.userId ? S.me.name : g.userId))}${g.pending ? "・端末に保存／未送信" : ""}</td><td>${num(g.qty)} ${esc(item.unit)}</td><td>${esc(C.fmtJst(g.at))}${controls && isAdmin() ? `<button class="btn sec sm" data-act="requestRecount" data-code="${esc(item.code)}" data-place="${esc(g.place)}" data-tier="${esc(g.tier)}">再カウント依頼</button>` : ""}</td></tr>`,
          )
          .join(
            "",
          )}</tbody></table></div><p>商品全体の合計 <b>${num(sum)} ${esc(item.unit)}</b></p>`
      : "<p>取得済みの記録には登録がありません。</p>"
  }
    <p class="muted">最終同期 ${esc(C.fmtJst(S.lastSyncAt) || "未取得")}。${S.offline || navigator.onLine === false ? "通信できません。他の人の最新記録は確認できていません。" : "保存時にも同じ場所の最新記録を確認します。"}</p></div>`;
}

// 作業位置を明示する。連続スキャンは選択したときだけ有効になる。
function workControls(place) {
  const a = EMP.assignment(place);
  const ti = EMP.tier(place);
  return `<div class="card"><h2>${esc(C.placeLabel(C.location(place, ti)))}</h2><p>${a ? `${esc(a.name)}・${a.submitted ? "提出済み" : "作業中"}` : "担当者はまだ決まっていません"}</p>
    ${place !== STOCK ? `<label>今数える段（棚札の末尾3桁）<input type="text" id="workTier" inputmode="numeric" maxlength="3" value="${esc(ti)}" placeholder="例 001"></label><button class="btn sec" data-act="scanLocation" data-place="${esc(place)}">棚札QRで場所を確認</button>` : ""}
    <label>数え方<select id="countMode"><option value="key">数量キー入力（基本）</option><option value="tally">1個ずつ連続スキャン</option></select></label>
    <p class="muted">${EMP.mode() === "tally" ? "1個読むごとに計数済み側へ動かし、「次の1個」を押してください。" : "商品を1回読んで、その場所の数量を入力します。"}</p>
    ${!a || a.submitted || a.userId !== S.me.userId || a.deviceId !== S.deviceId ? `<button class="btn sec" data-act="claimPlace" data-place="${esc(place)}">${a && a.submitted ? "この場所を再開して数える" : "この場所を担当して数える"}</button>` : ""}</div>`;
}
function bindWorkControls(place) {
  const m = $("countMode");
  if (m) {
    m.value = EMP.mode();
    m.onchange = () => {
      cur().params.countMode = m.value;
      render();
    };
  }
  const t = $("workTier");
  if (t)
    t.onchange = () => {
      const value = t.value.trim();
      if (!/^\d{3}$/.test(value))
        return toast("段は棚札と同じ3桁で入力してください", "error");
      cur().params.tier = value;
      render();
    };
}
const originalHome = SCREENS.home;
SCREENS.home = () => {
  originalHome();
  if (S.session)
    $("main").insertAdjacentHTML(
      "afterbegin",
      `<button class="btn sec" data-act="employeeTasks">${isAdmin() ? "発見品の確認・担当状況・再カウント" : "再カウントの依頼を確認"}</button>`,
    );
};
const originalPlaces = SCREENS.places;
SCREENS.places = () => {
  originalPlaces();
  $("main").insertAdjacentHTML(
    "afterbegin",
    `<div class="card"><b>担当・提出状況</b>${S.employee.assignments.map((a) => `<p>${esc(C.placeLabel(a.place))}：${esc(a.name)}・${a.submitted ? "提出済み" : "作業中"}</p>`).join("") || "<p>棚を開き、担当を開始してください。</p>"}</div>`,
  );
};
SCREENS.shelf = ({ shelf }) => {
  if (cur().params.tier == null)
    cur().params.tier =
      S.index.list.find((i) => i.shelf === shelf && i.inScope)?.tier || "";
  const tier = EMP.tier(shelf);
  const found = new Set(
    [...S.view.active.values()]
      .filter((e) => e.place === shelf && e.tier === tier)
      .map((e) => e.code),
  );
  const items = S.index.list.filter(
    (i) =>
      (i.shelf === shelf && i.tier === tier && i.inScope) || found.has(i.code),
  );
  const p = shelfStats(shelf);
  $("main").innerHTML =
    workControls(shelf) +
    `<div class="list">${
      items
        .map(
          (i) =>
            `<div class="row tap" data-act="tapEmployeeItem" data-code="${esc(i.code)}" data-place="${esc(shelf)}"><div class="main"><div class="name">${esc(i.name)}</div><div class="sub">${esc(i.code)}・${esc(i.unit)} ${i.review === "pending" ? "・管理者確認待ち" : ""}${i.shelf !== shelf ? "・別置き" : ""}</div><div class="sub">${groupRows(
              i.code,
            )
              .map(
                (g) =>
                  esc(C.placeLabel(C.location(g.place, g.tier))) +
                  " " +
                  esc(S.users[g.userId] || g.userId) +
                  " " +
                  num(g.qty),
              )
              .join(
                "／",
              )}</div></div><div class="side">${statusHtml(i, C.location(shelf, tier))}</div></div>`,
        )
        .join("") ||
      '<div class="row">この段の一覧には商品がありません。現物があれば検索して追加してください。</div>'
    }</div>`;
  setChrome(
    `${C.shelfLabel(shelf)}　済 ${p.counted}／${p.total}`,
    `<button class="btn dark" data-act="employeeScan" data-place="${esc(shelf)}">読取</button><button class="btn sec" data-act="search" data-place="${esc(shelf)}">一覧にない商品を追加</button><button class="btn pri" data-act="submitPlace" data-place="${esc(shelf)}">棚を提出</button>`,
  );
  bindWorkControls(shelf);
};
const originalStock = SCREENS.stock;
SCREENS.stock = () => {
  originalStock();
  $("main").insertAdjacentHTML("afterbegin", workControls(STOCK));
  bindWorkControls(STOCK);
  setChrome(
    "ストック",
    `<button class="btn dark" data-act="employeeScan" data-place="${STOCK}">読取</button><button class="btn sec" data-act="search" data-place="${STOCK}">検索して追加</button><button class="btn pri" data-act="submitPlace" data-place="${STOCK}">ストックを提出</button>`,
  );
};
SCREENS.search = ({ place }) => {
  setChrome(
    `全商品から検索（${C.placeLabel(C.location(place, EMP.tier(place)))}）`,
  );
  $("main").innerHTML =
    `<input type="text" id="q" placeholder="商品名・かな・商品コード・JAN" autocomplete="off"><p class="muted">棚卸開始時の商品マスター全体から探します。${S.offline ? "通信断中は端末に保持した範囲です。" : ""}</p><div id="qres"></div><button class="btn sec" data-act="unknownNoJan" data-place="${esc(place)}">見つからない商品を仮登録</button>`;
  $("q").oninput = () => {
    const hits = C.searchItems(S.index, $("q").value, 50);
    $("qres").innerHTML =
      hits
        .map(
          (i) =>
            `<div class="card"><button class="btn sec" data-act="tapEmployeeItem" data-code="${esc(i.code)}" data-place="${esc(place)}">${esc(i.name)}（${esc(i.code)}）</button><p>${esc(i.unit)}・定位置 ${esc(C.placeLabel(C.location(i.shelf, i.tier)))}${i.inScope ? "" : "・一覧外／発見登録"}</p>${sharedHtml(i, place, EMP.tier(place))}</div>`,
        )
        .join("") ||
      ($("q").value.trim()
        ? "<p>保持している商品に該当がありません。仮登録で記録を残せます。</p>"
        : "");
  };
};

// キー入力。置換対象を同じ場所・段だけに限定し、保存後の結果も表示する。
openQty = async (item, place, method, janRaw) => {
  const tier = EMP.tier(place, item);
  if (place !== STOCK && !/^\d{3}$/.test(tier))
    return toast("先に棚札と同じ段を選んでください", "error");
  if (!(await ensureClaim(place))) return;
  if (!(await confirmWorkLocation(place, tier))) return;
  if (recountFor(item.code, place, tier))
    return toast(
      "再カウントの確認中です。「再カウント」画面をご利用ください",
      "error",
    );
  await fetchState(false).catch(() => {
    S.offline = true;
  });
  const ids = C.groupIds(S.view, item.code, place, tier);
  const here = S.view.byKey.get(item.code + "|" + C.location(place, tier));
  const st = { value: "", mode: ids.length ? "" : "add", saving: false };
  openSheet(`<div class="q-name">${esc(item.name)}</div><p>${esc(item.code)}・${esc(item.unit)}・今数える場所 <b>${esc(C.placeLabel(C.location(place, tier)))}</b></p>
    ${sharedHtml(item, place, tier)}${!item.inScope ? '<div class="warnbox">一覧外の商品です。現物の数量を保存し、管理者が対象を確認します。</div>' : ""}
    ${ids.length ? '<div class="modes"><button data-emode="replace">この場所を数え直す</button><button data-emode="add">未計数の現物を追加で発見</button></div>' : "<p>この場所では新しく記録します。</p>"}
    <div class="q-val empty" id="employeeValue">数量</div><div class="keys">${[1, 2, 3, 4, 5, 6, 7, 8, 9, 0].map((n) => `<button data-ekey="${n}">${n}</button>`).join("")}<button data-ekey="back">⌫</button><button data-ekey="clear">消去</button></div>
    <button class="btn sec" id="employeeZero">0個・現物なし</button><p id="employeePreview">数量を入力してください。</p><button class="btn pri" id="employeeSave">端末に保存</button><button class="btn sec" id="employeeCancel">やめる</button>`);
  const update = () => {
    $("employeeValue").textContent = st.value === "" ? "数量" : st.value;
    const q = Number(st.value);
    $("employeePreview").textContent = !st.mode
      ? "追加か数え直しかを選んでください。"
      : st.value === ""
        ? "数量を入力してください。"
        : st.mode === "replace"
          ? `この場所の ${here?.qty || 0} ${item.unit} を ${q} ${item.unit} に変更します。他の場所は変更しません。`
          : `${q} ${item.unit} を追加し、この場所は合計 ${(here?.qty || 0) + q} ${item.unit} になります。`;
  };
  $("sheetBody")
    .querySelectorAll("[data-emode]")
    .forEach(
      (b) =>
        (b.onclick = () => {
          st.mode = b.dataset.emode;
          $("sheetBody")
            .querySelectorAll("[data-emode]")
            .forEach((x) => x.classList.toggle("on", x === b));
          update();
        }),
    );
  $("sheetBody")
    .querySelectorAll("[data-ekey]")
    .forEach(
      (b) =>
        (b.onclick = () => {
          const k = b.dataset.ekey;
          st.value =
            k === "clear"
              ? ""
              : k === "back"
                ? st.value.slice(0, -1)
                : (st.value === "0" ? "" : st.value) + k;
          st.value = st.value.slice(0, 5);
          update();
        }),
    );
  $("employeeZero").onclick = () => {
    st.value = "0";
    update();
  };
  $("employeeCancel").onclick = closeSheet;
  $("employeeSave").onclick = async () => {
    if (st.saving) return;
    if (st.value === "" || !st.mode)
      return toast("数量と登録方法を確認してください", "error");
    st.saving = true;
    try {
      const q = Number(st.value);
      if (
        q >= 500 &&
        !(await confirmYes(
          "数量の確認",
          `${q} ${item.unit} です。桁と単位を確認してください。`,
          "この数で保存",
        ))
      )
        return;
      const ok = await saveEntry({
        code: item.code,
        place,
        tier,
        qty: q,
        mode: st.mode,
        inputMethod: method,
        janRaw: janRaw || "",
        expectedIds: ids,
      });
      if (ok) {
        closeSheet();
        toast("端末に保存しました。送信状態を確認してください。");
      }
    } finally {
      st.saving = false;
    }
  };
};
checkPlaceThenQty = (item, place, method, janRaw) =>
  openQty(item, place, method, janRaw);
handleJan = async (raw, method, place) => {
  const hits = S.index.byJan.get(C.janKey(raw)) || [];
  if (!hits.length) return openUnknown(raw, place, method);
  let item = hits[0];
  if (hits.length > 1) {
    const k = await ask(
      "どの商品ですか？",
      "JANが重複しています。容量・色番・商品コードを確認してください。",
      hits
        .map((i) => ({ key: i.code, label: i.name + "（" + i.code + "）" }))
        .concat([{ key: "", label: "やめる" }]),
    );
    item = hits.find((i) => i.code === k);
  }
  if (item) await openQty(item, place, method, raw);
  else scan.resume();
};
const originalUnknown = openUnknown;
openUnknown = async (raw, place, method) => {
  if (!(await ensureClaim(place))) return;
  const tier = EMP.tier(place);
  if (place !== STOCK && !/^\d{3}$/.test(tier))
    return toast("先に段を選んでください", "error");
  if (!(await confirmWorkLocation(place, tier))) return;
  originalUnknown(raw, place, method);
  const h = $("sheetBody").querySelector(".q-name");
  if (h)
    h.textContent = raw ? "商品が見つかりません" : "一覧にない商品の仮登録";
  if ($("uMemo"))
    $("uMemo").placeholder = "商品名・容量・色番（JANなしは必須）";
  $("uSearch").onclick = () => {
    $("sheet").classList.add("hidden");
    scan.close();
    go("search", { place, tier, countMode: "key" });
  };
  const old = $("uSave").onclick;
  $("uSave").onclick = () => {
    if (!raw && !$("uMemo").value.trim())
      return toast("商品が分かるメモを入れてください", "error");
    return old();
  };
};

// 連続スキャンは1個ごとに明示的に再開する。検出イベントの連発では増えない。
const originalScanOpen = scan.open.bind(scan),
  originalScanClose = scan.close.bind(scan),
  originalDecode = scan.onDecode.bind(scan),
  originalResume = scan.resume.bind(scan);
scan.open = async function (place) {
  if (!(await ensureClaim(place))) return;
  const tier = EMP.tier(place);
  if (place !== STOCK && !/^\d{3}$/.test(tier))
    return toast("先に段を選んでください", "error");
  if (!(await confirmWorkLocation(place, tier))) return;
  this.mode = EMP.mode();
  this.tier = tier;
  $("scanManual").classList.remove("hidden");
  let saved = await kvGet(EMP.draftKey());
  if (saved && !saved.ops.length) {
    await kvPut(EMP.draftKey(), null);
    saved = null;
  }
  if (
    saved &&
    (saved.place !== place ||
      saved.tier !== tier ||
      saved.epoch !== EMP.assignment(place).epoch)
  ) {
    await ask(
      "別の作業の読取下書きがあります",
      C.placeLabel(C.location(saved.place, saved.tier)) +
        " の下書きが残っています。元の場所で保存または確認してください。",
      [{ key: "ok", label: "戻る" }],
    );
    return;
  }
  if (saved && this.mode !== "tally") {
    toast(
      "連続スキャンの下書きがあります。連続スキャンを選んで確認してください。",
      "error",
    );
    return;
  }
  EMP.draft = saved || {
    sessionId: sid(),
    userId: S.me.userId,
    place,
    tier,
    epoch: EMP.assignment(place).epoch,
    rows: [],
    ops: [],
  };
  if (!$("tallyControls"))
    $("scanBox").insertAdjacentHTML(
      "afterend",
      '<div id="tallyControls" style="background:#fff;color:#111;padding:12px;max-height:40vh;overflow:auto"></div>',
    );
  await originalScanOpen(place);
  drawTally();
};
function drawTally() {
  const el = $("tallyControls");
  if (!el) return;
  el.classList.toggle("hidden", scan.mode !== "tally");
  if (scan.mode !== "tally") return;
  const d = EMP.draft;
  $("scanLabel").textContent =
    C.placeLabel(C.location(d.place, d.tier)) + "・1個ずつ連続スキャン";
  el.innerHTML = `<b>今回 ${d.ops.length} 個を読取（まだ数量へ反映していません）</b><p>${d.ops.length ? esc(S.index.byCode.get(d.ops.at(-1).code)?.name || "") : "現物を1個ずつ読んで、計数済み側へ動かしてください。"}</p>
    <p>状態：${scan.paused ? "停止中。現物を動かしてから次へ" : "次の1個を待っています"}</p>
    <button class="btn dark" id="tallyNext">次の1個を読む</button><button class="btn sec" id="tallyPause">一時停止</button>
    <button class="btn sec" id="tallyUndo" ${d.ops.length ? "" : "disabled"}>直前の1個を取消</button>
    <details><summary>今回の読取一覧</summary>${d.rows.map((x) => `<p>${esc(S.index.byCode.get(x.code)?.name || x.code)}：${x.qty}</p>`).join("")}</details>
    <button class="btn pri" id="tallySave" ${d.ops.length ? "" : "disabled"}>今回分を保存</button>`;
  const lastItem = d.ops.length && S.index.byCode.get(d.ops.at(-1).code);
  if (lastItem)
    el.insertAdjacentHTML(
      "beforeend",
      `<details><summary>既存の登録者・場所を見る</summary>${sharedHtml(lastItem, d.place, d.tier)}</details>`,
    );
  $("tallyNext").onclick = () => {
    if (EMP.busy) return;
    scan.paused = false;
    drawTally();
  };
  $("tallyPause").onclick = () => {
    scan.paused = true;
    drawTally();
  };
  $("tallyUndo").onclick = async () => {
    if (EMP.busy) return;
    const last = d.ops.at(-1);
    if (!last) return;
    EMP.busy = true;
    try {
      const copy = structuredClone(d);
      copy.ops.pop();
      const row = copy.rows.find((x) => x.code === last.code);
      row.qty--;
      if (!row.qty) copy.rows = copy.rows.filter((x) => x !== row);
      await persistDraft(copy);
      scan.paused = true;
      drawTally();
    } catch (e) {
      toast(
        "取消を端末に保存できませんでした。数量は変更していません。",
        "error",
      );
    } finally {
      EMP.busy = false;
    }
  };
  $("tallySave").onclick = saveTally;
}
async function persistDraft(d) {
  await kvPut(EMP.draftKey(), d);
  EMP.draft = d;
}
scan.onDecode = async function (text) {
  if (this.mode === "location") {
    if (this.paused || !this.active) return;
    this.paused = true;
    const raw = String(text).trim();
    const match = /^STLOC\|([^|]+)\|(\d{3})$/.exec(raw);
    const place = match ? match[1] : raw.slice(0, -3).replace(/-$/, ""),
      tier = match ? match[2] : raw.slice(-3);
    if (place !== this.place || !/^\d{3}$/.test(tier)) {
      beep(false);
      await ask(
        "棚札が違います",
        "選択中は " +
          C.placeLabel(this.place) +
          " です。正しい棚札を読んでください。",
        [{ key: "ok", label: "確認した" }],
      );
      this.paused = false;
      return;
    }
    cur().params.tier = tier;
    EMP.verified.add(workLocationKey(place, tier));
    beep(true);
    this.close();
    toast(C.placeLabel(C.location(place, tier)) + " の棚札を確認しました");
    return;
  }
  if (this.mode !== "tally") return originalDecode(text);
  if (this.paused || !this.active || EMP.busy) return;
  this.paused = true;
  EMP.busy = true;
  drawTally();
  try {
    const raw = C.digits(text);
    if (!C.eanCheckOk(raw)) {
      beep(false);
      return;
    }
    let hits = S.index.byJan.get(C.janKey(raw)) || [];
    if (hits.length !== 1) {
      beep(false);
      await ask(
        "この商品は数量に加えていません",
        hits.length
          ? "JANが重複しています。検索して商品を確認してください。"
          : "登録のないJANです。キー入力または検索で発見記録を残してください。",
        [{ key: "ok", label: "確認した" }],
      );
      return;
    }
    const item = hits[0],
      d = structuredClone(EMP.draft);
    if (recountFor(item.code, d.place, d.tier)) {
      beep(false);
      toast("再カウント確認中の商品です", "error");
      return;
    }
    let row = d.rows.find((x) => x.code === item.code);
    if (!row) {
      const ids = C.groupIds(S.view, item.code, d.place, d.tier);
      const prior = groupRows(item.code);
      if (
        prior.length &&
        !(await confirmYes(
          "登録済みの場所を確認",
          item.name +
            "\n" +
            prior
              .map(
                (g) =>
                  C.placeLabel(C.location(g.place, g.tier)) +
                  "・" +
                  (S.users[g.userId] || g.userId) +
                  "・" +
                  g.qty +
                  " " +
                  item.unit,
              )
              .join("\n") +
            "\n今数える場所は " +
            C.placeLabel(C.location(d.place, d.tier)) +
            " です。まだ数えていない別の現物ですか？",
          "この場所の未計数分を登録",
        ))
      )
        return;
      row = {
        id: C.uuid(),
        code: item.code,
        janRaw: raw,
        qty: 0,
        expectedIds: ids,
      };
      d.rows.push(row);
    }
    if (row.qty >= 99999)
      return toast("数量の上限です。下書きを確認してください", "error");
    row.qty++;
    d.ops.push({ code: item.code });
    await persistDraft(d);
    beep(true);
  } catch (e) {
    beep(false);
    toast("端末に保存できませんでした。この読取は加えていません。", "error");
  } finally {
    EMP.busy = false;
    drawTally();
  }
};
scan.resume = function () {
  if (this.mode === "tally") {
    this.paused = true;
    drawTally();
  } else originalResume();
};
scan.close = function () {
  this.paused = true;
  originalScanClose();
};
async function saveTally() {
  if (EMP.busy) return;
  EMP.busy = true;
  scan.paused = true;
  try {
    const d = EMP.draft;
    if (!d?.rows.length) return;
    const a = EMP.assignment(d.place);
    if (!a || a.epoch !== d.epoch || a.userId !== S.me.userId)
      return toast(
        "担当が変わっています。下書きを管理者へ引き継いでください",
        "error",
      );
    const ops = d.rows.map((x) => ({
      k: opKey(),
      createdAt: Date.now(),
      userId: S.me.userId,
      type: "put",
      sessionId: sid(),
      entry: {
        id: x.id,
        kind: "item",
        code: x.code,
        janRaw: x.janRaw,
        place: d.place,
        tier: d.tier,
        mode: "add",
        qty: x.qty,
        inputMethod: "scan",
        memo: "連続スキャン",
        deviceId: S.deviceId,
        clientAt: new Date().toISOString(),
        userId: S.me.userId,
        expectedIds: x.expectedIds,
        assignmentEpoch: d.epoch,
        voided: false,
      },
    }));
    // 下書きの削除と送信待ちへの移動を同じ取引で行い、途中失敗による欠落を防ぐ。
    await new Promise((resolve, reject) => {
      const t = IDB.db.transaction(["kv", "outbox"], "readwrite");
      for (const o of ops) t.objectStore("outbox").put(o);
      t.objectStore("kv").delete(EMP.draftKey());
      t.oncomplete = resolve;
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error);
    });
    S.ops.push(...ops);
    EMP.draft = null;
    rebuildView();
    scan.close();
    flush();
    toast("今回分を端末に保存しました。送信状態を確認してください。");
  } catch (e) {
    toast("保存できませんでした。読取下書きは残っています。", "error");
  } finally {
    EMP.busy = false;
  }
}

// 管理者の確認と、独立した再カウントの作業画面。
SCREENS.employeeTasks = () => {
  setChrome(isAdmin() ? "発見品・担当・再カウント" : "再カウント");
  let h =
    '<button class="btn sec" data-act="refreshEmployeeTasks">最新情報を更新</button>';
  if (isAdmin()) {
    h += `<div class="card"><h2>対象範囲の確認</h2>${S.employee.reviews.map((i) => `<p>${esc(i.name)}（${esc(i.code)}） <button class="btn sec sm" data-act="reviewItem" data-code="${esc(i.code)}" data-decision="accepted">対象に含める</button><button class="btn sec sm" data-act="reviewItem" data-code="${esc(i.code)}" data-decision="excluded">対象外と確認</button></p>`).join("") || "<p>確認待ちはありません。</p>"}</div>`;
    h += `<div class="card"><h2>不明品</h2>${
      [...S.view.active.values()]
        .filter((e) => e.kind === "unknown")
        .map(
          (e) =>
            `<p>${esc(e.janRaw || "JANなし")} ${esc(e.memo)}・${esc(C.placeLabel(C.location(e.place, e.tier)))} ${num(e.qty)}個<button class="btn sec sm" data-act="resolveUnknown" data-id="${esc(e.id)}">商品へ紐付け／除外</button></p>`,
        )
        .join("") || "<p>ありません。</p>"
    }</div>`;
    h += `<div class="card"><h2>重複・遅延記録の確認</h2>${S.employee.conflicts.map((c) => `<p>${esc(c.userId)}・${esc(c.payload.code || c.payload.memo || c.payload.janRaw)}・${esc(C.placeLabel(C.location(c.payload.place, c.payload.tier)))}・${num(c.payload.qty)} ${esc(errText(c.reason))}<button class="btn sec sm" data-act="resolveConflict" data-id="${esc(c.entryId)}">照合して処理済みにする</button></p>`).join("") || "<p>ありません。</p>"}</div>`;
    const codes = [
      ...new Set(
        [...S.view.active.values()]
          .filter((e) => e.kind === "item")
          .map((e) => e.code),
      ),
    ];
    h += `<div class="card"><h2>別置き・再確認</h2><p class="muted">別置き分は残して合算します。商品を動かすのは棚卸確定後です。</p>${codes
      .map((code) => {
        const i = S.index.byCode.get(code);
        return i
          ? `<details><summary>${esc(i.name)}（${groupRows(code).length}件の担当・場所）</summary>${sharedHtml(i, "", "", true)}</details>`
          : "";
      })
      .join("")}</div>`;
  }
  h += `<div class="card"><h2>再カウント</h2>${S.employee.recounts.map((r) => `<div class="card"><b>${esc(S.index.byCode.get(r.code)?.name || r.code)}</b><p>${esc(C.placeLabel(C.location(r.place, r.tier)))}・${esc(r.reason)}・${esc(recountStatus[r.state] || r.state)}</p>${r.state === "open" && r.assignee === S.me.userId ? `<button class="btn pri" data-act="saveRecount" data-id="${esc(r.id)}">現物を数えて入力</button>` : ""}${isAdmin() && ["open", "counted"].includes(r.state) ? `<p>担当：${esc(S.employee.users.find((u) => u.id === r.assignee)?.name || r.assignee)}／再カウント ${r.qty == null ? "未入力" : num(r.qty)}</p>${r.state === "counted" ? `<button class="btn pri" data-act="adoptRecount" data-id="${esc(r.id)}">この場所の数量として採用</button>` : ""}<button class="btn sec" data-act="dismissRecount" data-id="${esc(r.id)}">初回を維持して確認を終える</button>` : ""}</div>`).join("") || "<p>依頼はありません。</p>"}</div>`;
  $("main").innerHTML = h;
};
const originalCloseScreen = SCREENS.close;
SCREENS.close = () => {
  originalCloseScreen();
  const places = S.shelves
    .filter((s) => s.inScope)
    .map((s) => s.shelf)
    .concat(STOCK);
  const pending = places.filter((p) => !EMP.assignment(p)?.submitted);
  const unresolved =
    S.employee.reviews.length +
    S.employee.conflicts.length +
    S.employee.recounts.filter((r) => ["open", "counted"].includes(r.state))
      .length +
    [...S.view.active.values()].filter((e) => e.kind === "unknown").length;
  $("main").insertAdjacentHTML(
    "afterbegin",
    `<div class="card"><b>全担当の提出確認</b><p>${pending.length ? "未提出：" + esc(pending.map(C.placeLabel).join("、")) : "すべての担当場所が送信済みとして提出されています。"}</p><p>未解決の確認事項：${unresolved}</p><button class="btn sec" data-act="employeeTasks">確認画面へ</button></div>`,
  );
  if (
    pending.length ||
    unresolved ||
    S.rejected.some((o) => o.sessionId === sid())
  ) {
    const b = $("bottomInner").querySelector('[data-act="doClose"]');
    if (b) b.disabled = true;
  }
};
Object.assign(ACT, {
  scanLocation: async (d) => {
    if (!(await ensureClaim(d.place))) return;
    scan.mode = "location";
    await originalScanOpen(d.place);
    $("tallyControls")?.classList.add("hidden");
    $("scanManual").classList.add("hidden");
    $("scanLabel").textContent = C.placeLabel(d.place) + " の棚札QRを確認";
  },
  refreshEmployeeTasks: async () => {
    await reloadEmployee();
    render();
  },
  employeeTasks: async () => {
    await fetchState(false);
    go("employeeTasks");
  },
  claimPlace: async (d) => {
    if (await ensureClaim(d.place)) render();
  },
  tapEmployeeItem: (d) => {
    const i = S.index.byCode.get(d.code);
    if (i) return openQty(i, d.place, "list");
  },
  tapItem: (d) => {
    const i = S.index.byCode.get(d.code);
    if (i) return openQty(i, cur().params.shelf, "list");
  },
  pickSearch: (d) => {
    const i = S.index.byCode.get(d.code);
    if (i) return openQty(i, d.place, "search");
  },
  employeeScan: (d) => scan.open(d.place),
  scan: (d) => scan.open(d.place),
  search: (d) =>
    go("search", {
      place: d.place,
      tier: EMP.tier(d.place),
      countMode: EMP.mode(),
    }),
  unknownNoJan: (d) => openUnknown("", d.place, "search"),
  submitPlace: async (d) => {
    const a = EMP.assignment(d.place);
    if (!a || a.userId !== S.me.userId || a.deviceId !== S.deviceId)
      return toast("担当を開始した端末から提出してください", "error");
    const draft = await kvGet(EMP.draftKey());
    if (draft?.ops.length)
      return toast("連続スキャンの下書きを先に保存してください", "error");
    if (
      !(await ensureFlushed(10000)) ||
      S.rejected.some((o) => o.sessionId === sid())
    )
      return toast("未送信・送れなかった記録を確認してください", "error");
    if (
      !(await confirmYes(
        "現物の確認と提出",
        "棚の奥・上・足元も確認し、記録していない商品は残っていませんか？",
        "確認して提出",
      ))
    )
      return;
    if (
      await employeeAction("submit", {
        place: d.place,
        assignmentEpoch: a.epoch,
        physicalChecked: true,
      })
    ) {
      await fetchState(false);
      toast("送信済みとして提出しました");
      back();
    }
  },
  reviewItem: async (d) => {
    const x = await ask(
      "対象範囲の確認",
      "採用・除外する理由を残してください。",
      [
        { key: "ok", label: "記録する" },
        { key: "cancel", label: "やめる" },
      ],
      { placeholder: "理由" },
    );
    if (
      x.key === "ok" &&
      x.value &&
      (await employeeAction("review", {
        code: d.code,
        decision: d.decision,
        reason: x.value,
      }))
    ) {
      await reloadEmployee();
      render();
    }
  },
  resolveConflict: async (d) => {
    const x = await ask(
      "記録の照合",
      "この記録は数量へ加えていません。現物・登録済み数量を確認し、必要な訂正後に理由を残してください。",
      [
        { key: "ok", label: "照合結果を記録" },
        { key: "cancel", label: "やめる" },
      ],
      { placeholder: "二重記録のため除外／別途訂正済み等" },
    );
    if (
      x.key === "ok" &&
      x.value &&
      (await employeeAction("conflictResolve", {
        entryId: d.id,
        reason: x.value,
      }))
    ) {
      await fetchState(false);
      render();
    }
  },
  resolveUnknown: async (d) => {
    const e = S.view.active.get(d.id);
    if (!e) return;
    const x = await ask(
      "不明品の解決",
      "既存商品の商品コードを入力してください。対象外と確認した場合は「対象外」と入力します。",
      [
        { key: "ok", label: "次へ" },
        { key: "cancel", label: "やめる" },
      ],
      { placeholder: "商品コード／対象外" },
    );
    if (x.key !== "ok" || !x.value) return;
    const code = x.value === "対象外" ? null : x.value;
    const i = code && S.index.byCode.get(code);
    if (code && !i) return toast("商品コードが見つかりません", "error");
    if (
      !(await confirmYes(
        "紐付け内容の確認",
        code
          ? `${i.name} に ${e.qty} ${i.unit} を加えます。同じ現物を既に登録していないことを確認してください。`
          : "この発見記録を対象外とします。",
        "確認した",
      ))
    )
      return;
    const why = await ask(
      "理由",
      "確認した内容を残します。",
      [
        { key: "ok", label: "保存" },
        { key: "cancel", label: "やめる" },
      ],
      { placeholder: "理由" },
    );
    if (
      why.key === "ok" &&
      why.value &&
      (await employeeAction("unknownResolve", {
        entryId: d.id,
        code,
        expectedIds: C.groupIds(S.view, code, e.place, e.tier),
        reason: why.value,
      }))
    ) {
      await reloadEmployee();
      render();
    }
  },
  requestRecount: async (d) => {
    const users = S.employee.users.filter(
      (u) =>
        !groupRows(d.code).some(
          (g) => g.place === d.place && g.tier === d.tier && g.userId === u.id,
        ),
    );
    const who = await ask(
      "再カウント担当",
      "初回と違う人を選んでください。",
      users
        .map((u) => ({ key: u.id, label: u.name }))
        .concat([{ key: "", label: "やめる" }]),
    );
    if (!who) return;
    const why = await ask(
      "再カウントの理由",
      C.placeLabel(C.location(d.place, d.tier)),
      [
        { key: "ok", label: "依頼する" },
        { key: "cancel", label: "やめる" },
      ],
      { placeholder: "理由" },
    );
    if (
      why.key === "ok" &&
      why.value &&
      (await employeeAction("recountRequest", {
        code: d.code,
        place: d.place,
        tier: d.tier,
        assignee: who,
        reason: why.value,
      }))
    ) {
      await fetchState(false);
      render();
    }
  },
  saveRecount: async (d) => {
    const x = await ask(
      "現物を数えて入力",
      "帳簿や初回数量に合わせず、この場所に実在する数量を入力してください。",
      [
        { key: "ok", label: "再カウントを提出" },
        { key: "cancel", label: "やめる" },
      ],
      { placeholder: "数量（0を含む整数）" },
    );
    if (x.key !== "ok") return;
    if (!/^\d{1,5}$/.test(x.value))
      return toast("数量を整数で入力してください", "error");
    if (
      await employeeAction("recountSave", { id: d.id, qty: Number(x.value) })
    ) {
      await fetchState(false);
      render();
    }
  },
  adoptRecount: async (d) => {
    if (
      await confirmYes(
        "再カウントを採用",
        "この場所の初回数量を置き換えます。他の場所の数量は維持します。",
        "採用する",
      )
    )
      if (
        await employeeAction("recountAdopt", {
          id: d.id,
          reason: "再カウント結果を確認して採用",
        })
      ) {
        await fetchState(false);
        render();
      }
  },
  dismissRecount: async (d) => {
    const x = await ask(
      "初回の数量を維持",
      "判断理由を残してください。",
      [
        { key: "ok", label: "確認を終える" },
        { key: "cancel", label: "やめる" },
      ],
      { placeholder: "理由" },
    );
    if (
      x.key === "ok" &&
      x.value &&
      (await employeeAction("recountDismiss", { id: d.id, reason: x.value }))
    ) {
      await fetchState(false);
      render();
    }
  },
});

// 前回の読取下書きが残っていることを、再起動後にも発見できるようにする。
const homeWithEmployee = SCREENS.home;
SCREENS.home = () => {
  homeWithEmployee();
  if (S.session)
    kvGet(EMP.draftKey())
      .then((d) => {
        if (cur().name === "home" && d?.ops.length) {
          $("employeeDraftHint")?.remove();
          $("main").insertAdjacentHTML(
            "afterbegin",
            `<div class="warnbox" id="employeeDraftHint">連続スキャンの下書き ${d.ops.length}個：${esc(C.placeLabel(C.location(d.place, d.tier)))}<button class="btn sec" data-act="resumeTally">下書きのある場所で再開</button></div>`,
          );
        }
      })
      .catch(() => {});
};
ACT.resumeTally = async () => {
  const d = await kvGet(EMP.draftKey());
  if (!d) return;
  go(d.place === STOCK ? "stock" : "shelf", {
    shelf: d.place,
    tier: d.tier,
    countMode: "tally",
  });
  await scan.open(d.place);
};

// 現場の選択欄も指で押せる大きさに揃える。
const employeeStyle = document.createElement("style");
employeeStyle.textContent =
  "select,#workTier,#devScanCode{font:inherit;font-size:16px;min-height:48px;width:100%;padding:10px;border:1px solid var(--border);border-radius:10px;background:white;color:var(--text);margin:6px 0 12px}.shared-records{margin:8px 0}.hd{grid-template-columns:76px minmax(0,1fr) 72px}.hd-user{min-width:0;max-width:72px}.bottom .btn{padding:10px 8px;white-space:normal;line-height:1.4}.scanner #scanBox{max-height:38vh;overflow:auto}.scanner .scan-tools{flex-wrap:wrap}";
document.head.appendChild(employeeStyle);

// 段の入力中は自動更新で入力欄を作り直さない。
const employeeRenderLight = renderLight;
renderLight = () => {
  if (["workTier", "q"].includes(document.activeElement?.id)) return;
  employeeRenderLight();
};
