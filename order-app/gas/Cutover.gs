// ============================================================
// S3 切替・切り戻しの操作用（GASエディタで関数を選んで「実行」する）
// 手順書: order-app/docs/S3_2画面切替手順_20260927.md（設計 §13-4・§13-5 を2画面運用に合わせたもの）
//
// - どの関数も、最初に今の状態（DBの writer・移行段階、旧書き込み LEGACY_WRITER）を確かめ、
//   想定と違えば何も変えずに「STOP」をログに出して終わる。同じ関数を2回押しても壊れない
// - 関数名の末尾に _ を付けない（末尾 _ はGASが非公開扱いにし、エディタの実行メニューに出ない）
// - 画面（Webアプリ）からは呼べない（doPost の action に無い）。エディタの実行は HEAD のコードで動くので、
//   このファイルの追加は clasp push だけでよい（利用者が使うデプロイ @68 は変わらない）
// ============================================================

function cutoverWriterState_() {
  const w = callDbAdmin_('get_writer', {});
  if (!w || w.success !== true) throw new Error('get_writer に失敗しました: ' + JSON.stringify(w));
  return { writer: w.writer, epoch: w.writerEpoch, phase: w.migrationPhase, legacy: isLegacyWriterOpen_() ? 'open' : 'closed' };
}

function cutoverLogState_(label) {
  const s = cutoverWriterState_();
  const triggers = ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); });
  Logger.log(label + ': DB writer=' + s.writer + ' / 移行段階=' + s.phase + ' / epoch=' + s.epoch +
    ' / 旧書き込み(LEGACY_WRITER)=' + s.legacy + ' / 写しの位置(DB_SHEET_COPY_SEQ)=' + (_PROPS.getProperty('DB_SHEET_COPY_SEQ') || '未設定') +
    ' / トリガー=' + JSON.stringify(triggers));
  return s;
}

// DB の変更記録（change_log）の最後の番号。写しの開始位置・写しが追いついたかの確認に使う
function cutoverLatestSeq_() {
  let after = 0;
  for (let i = 0; i < 200; i++) {
    const r = callDbAdmin_('get_changes_since', { afterSeq: after, limit: 500 });
    if (!r || r.success !== true) throw new Error('get_changes_since に失敗しました: ' + JSON.stringify(r));
    if (!r.changes || r.changes.length === 0) return after;
    after = r.lastSeq;
  }
  throw new Error('変更記録が多すぎて最後まで数えられませんでした');
}

function cutoverDeleteTrigger_(handler) {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === handler; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
}

// ---- いつでも実行してよい（読むだけ・何も変えない） ----
function cutoverStatus() {
  assertTestEnv_();
  cutoverLogState_('現在の状態');
}

// ---- 切替 1: 旧の受付を止め、DB も受付停止（closed）にする ----
// 前提: writer=shadow。完了後は「どこも発注を受け付けない」状態になる（この間に画面で保存しないこと）
function cutoverStep1_close() {
  assertTestEnv_();
  const s = cutoverLogState_('切替1 開始前');
  if (s.writer === 'closed' && s.phase === 'cutover' && s.legacy === 'closed') {
    Logger.log('切替1 はすでに済んでいます。次は cutoverStep2_importAndVerify を実行してください');
    return;
  }
  if (s.writer !== 'shadow' || s.phase !== 'shadow') {
    Logger.log('STOP: writer=shadow・移行段階=shadow のときだけ実行できます。何も変えていません');
    return;
  }
  deleteDbMigrateTrigger();   // 毎晩2時の取り込みを止める（手順2）
  closeLegacyWriter();        // 旧GASの書き込みを止める（手順3）
  const sw = callDbAdmin_('set_writer', { expectedFrom: 'shadow', to: 'closed', reason: 'S3 cutover (2-screen) step1' });
  if (!sw || sw.success !== true) {
    Logger.log('STOP: DB を closed にできませんでした: ' + JSON.stringify(sw) +
      '。旧の受付は止まったままです。cutoverAbort を実行して元に戻してください');
    return;
  }
  cutoverLogState_('切替1 完了');
  Logger.log('OK 切替1 完了。次は cutoverStep2_importAndVerify を実行してください');
}

// ---- 切替 2: シートの最終取り込み → 全件の内容照合（差が1件でもあれば進めない） ----
function cutoverStep2_importAndVerify() {
  assertTestEnv_();
  const s = cutoverLogState_('切替2 開始前');
  if (s.writer !== 'closed' || s.phase !== 'cutover' || s.legacy !== 'closed') {
    Logger.log('STOP: 切替1 の後（writer=closed・移行段階=cutover・旧書き込み=closed）でだけ実行できます。何も変えていません');
    return;
  }
  _PROPS.deleteProperty('CUTOVER_VERIFIED_EPOCH'); // やり直したら、成功するまで切替3へ進めない
  const m = migrateToDb();
  if (!m || m.success !== true) {
    Logger.log('STOP: 最終取り込みの件数が合いません: ' + JSON.stringify(m && (m.mismatch || m.step)) +
      '。cutoverAbort で元に戻すか、原因を確認してから切替2をやり直してください');
    return;
  }
  Logger.log('最終取り込み OK: ' + JSON.stringify(m.compare));
  const v = verifyDbAgainstSheets_();
  if (!v || v.success !== true) {
    Logger.log('STOP: 内容照合で差があります（' + (v && v.diffTotal) + '件）: ' + JSON.stringify(v && (v.diffs || v.step)).slice(0, 3000) +
      '。cutoverAbort で元に戻してください');
    return;
  }
  _PROPS.setProperty('CUTOVER_VERIFIED_EPOCH', String(s.epoch)); // 切替3 はこの印（同じ epoch）が無ければ進まない
  Logger.log('OK 切替2 完了（取り込み・内容照合とも差0）。次は分析PCで「切替用の分析」を実行してください（手順書 手順5）');
}

// ---- 切替 3: 写しの開始位置を決め、DB を発注の受付先（db）にし、写しトリガーを作る ----
// 前提: 切替2 の後に分析PCで --cutover-analysis が成功している（無ければ DB が WRITER_ANALYSIS_NOT_READY で断る）
function cutoverStep3_openDb() {
  assertTestEnv_();
  const s = cutoverLogState_('切替3 開始前');
  if (s.writer === 'db') {
    Logger.log('切替3 はすでに済んでいます（writer=db）');
    return;
  }
  if (s.writer !== 'closed' || s.phase !== 'cutover' || s.legacy !== 'closed') {
    Logger.log('STOP: 切替2 の後（writer=closed・移行段階=cutover・旧書き込み=closed）でだけ実行できます。何も変えていません');
    return;
  }
  if (_PROPS.getProperty('CUTOVER_VERIFIED_EPOCH') !== String(s.epoch)) {
    Logger.log('STOP: 切替2（最終取り込み・内容照合）が成功していません。cutoverStep2_importAndVerify を先に実行してください。何も変えていません');
    return;
  }
  const seq = cutoverLatestSeq_();
  _PROPS.setProperty('DB_SHEET_COPY_SEQ', String(seq));
  const sw = callDbAdmin_('set_writer', { expectedFrom: 'closed', to: 'db', reason: 'S3 cutover (2-screen) step3' });
  if (!sw || sw.success !== true) {
    Logger.log('STOP: DB を db にできませんでした: ' + JSON.stringify(sw) +
      '（WRITER_ANALYSIS_NOT_READY なら、切替2 の後の切替用の分析が公開されていません）。' +
      'まだどこも受け付けない状態です。分析をやり直すか、cutoverAbort で元に戻してください');
    return;
  }
  setDbSheetCopyTrigger();
  _PROPS.deleteProperty('CUTOVER_VERIFIED_EPOCH');
  cutoverLogState_('切替3 完了');
  Logger.log('OK 切替完了。新画面で発注できます。旧画面は見るだけ（保存は MOVED_TO_DB で止まる）');
}

// ---- 切替の中止（切替1〜3の途中で止めて、影運用に戻す） ----
function cutoverAbort() {
  assertTestEnv_();
  const s = cutoverLogState_('中止 開始前');
  if (s.writer === 'closed' && s.phase === 'cutover') {
    const sw = callDbAdmin_('set_writer', { expectedFrom: 'closed', to: 'shadow', reason: 'S3 cutover aborted' });
    if (!sw || sw.success !== true) {
      Logger.log('STOP: DB を shadow に戻せませんでした: ' + JSON.stringify(sw) + '。旧の受付は止めたままにしています');
      return;
    }
  } else if (!(s.writer === 'shadow' && s.phase === 'shadow')) {
    Logger.log('STOP: 切替の途中ではありません（切替完了後に戻すのは rollbackStep1_closeDb から）。何も変えていません');
    return;
  }
  _PROPS.deleteProperty('CUTOVER_VERIFIED_EPOCH');
  openLegacyWriter();
  setDbMigrateTrigger();
  cutoverLogState_('中止 完了');
  Logger.log('OK 中止しました。旧画面で今までどおり発注できます（影運用に戻りました）');
}

// ---- 切り戻し 1: DB の受付を止め、写しを最後まで進めてから照合する（設計 §13-5 手順2〜4） ----
// 分析の自動実行（6:45〜7:45）と重ならない時間に行う
function rollbackStep1_closeDb() {
  assertTestEnv_();
  const s = cutoverLogState_('切り戻し1 開始前');
  if (s.writer === 'db') {
    const sw = callDbAdmin_('set_writer', { expectedFrom: 'db', to: 'closed', reason: 'S3 rollback step1' });
    if (!sw || sw.success !== true) { Logger.log('STOP: DB を closed にできませんでした: ' + JSON.stringify(sw)); return; }
  } else if (!(s.writer === 'closed' && s.phase === 'done')) {
    Logger.log('STOP: 切替完了後（writer=db）でだけ実行できます。何も変えていません');
    return;
  }
  _PROPS.deleteProperty('ROLLBACK_VERIFIED_SEQ');
  const target = cutoverLatestSeq_();
  for (let i = 0; i < 30; i++) {
    const at = parseInt(_PROPS.getProperty('DB_SHEET_COPY_SEQ'), 10) || 0;
    if (at >= target) break;
    copyChangesToSheets_();
    if ((parseInt(_PROPS.getProperty('DB_SHEET_COPY_SEQ'), 10) || 0) === at) {
      Logger.log('STOP: 写しが進みません（写しの位置 ' + at + ' / 目標 ' + target + '）。実行ログの copyChangesToSheets_ の失敗を確認してください');
      return;
    }
  }
  const copied = parseInt(_PROPS.getProperty('DB_SHEET_COPY_SEQ'), 10) || 0;
  if (copied < target) { Logger.log('STOP: 写しが最後まで進みませんでした（' + copied + ' / ' + target + '）。もう一度実行してください'); return; }
  cutoverDeleteTrigger_('copyChangesToSheets_');
  const v = verifyDbAgainstSheets_();
  if (!v || v.success !== true) {
    Logger.log('STOP: 写しの後の照合で差があります（' + (v && v.diffTotal) + '件）: ' + JSON.stringify(v && (v.diffs || v.step)).slice(0, 3000) +
      '。旧の受付はまだ開けていません。差を直してからやり直してください');
    return;
  }
  _PROPS.setProperty('ROLLBACK_VERIFIED_SEQ', String(target)); // 切り戻し2 はこの印が今の最後の番号と同じときだけ進む
  const snap = callDbAdmin_('get_compare_snapshot', {});
  Logger.log('OK 切り戻し1 完了（写し・照合とも差0）。公開中の分析ID=' + (snap && snap.analysisId) +
    '。次は分析PCで --republish-to-gas <この分析ID> を実行し、その後 rollbackStep2_openLegacy を実行してください');
}

// ---- 切り戻し 2: 提案シートへ今効いている設定を当て直し、旧の受付を開ける（設計 §13-5 手順5〜7） ----
// 前提: 分析PCで --republish-to-gas が成功し、分析PCの config.json を backend:'gas' に戻してある
function rollbackStep2_openLegacy() {
  assertTestEnv_();
  const s = cutoverLogState_('切り戻し2 開始前');
  if (s.writer !== 'closed' || s.phase !== 'done') {
    Logger.log('STOP: 切り戻し1 の後（writer=closed・移行段階=done）でだけ実行できます。何も変えていません');
    return;
  }
  if (_PROPS.getProperty('ROLLBACK_VERIFIED_SEQ') !== String(cutoverLatestSeq_())) {
    Logger.log('STOP: 切り戻し1（写しを最後まで進めて照合）が成功していません。rollbackStep1_closeDb を先に実行してください。何も変えていません');
    return;
  }
  const snap = callDbAdmin_('get_compare_snapshot', {});
  const analysisId = snap && snap.analysisId;
  if (!analysisId) { Logger.log('STOP: 公開中の分析IDが取れません: ' + JSON.stringify(snap).slice(0, 300)); return; }
  reapplyEffectiveToProposals_(analysisId);
  _PROPS.deleteProperty('ROLLBACK_VERIFIED_SEQ');
  openLegacyWriter();
  cutoverLogState_('切り戻し2 完了');
  Logger.log('OK 切り戻し完了。旧画面で発注できます。新画面は保存が WRITES_CLOSED で止まります');
}
