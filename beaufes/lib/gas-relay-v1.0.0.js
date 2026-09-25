/* GAS中継（透明iframe＋google.script.run）。GitHub Pages の画面から、fetch→/exec を通さずにGASを呼ぶ。
 *
 * なぜ: fetch→/exec（ContentService）は結果の配送で失敗する（2026-09-24 A/B実測：fetch 10.6%失敗／google.script.run 0%）。
 * 使えない端末: Googleアカウントに2つ以上同時ログインしているブラウザでは中継ページが開けない（Googleの仕様）。
 *   → 呼び出し側は必ず「中継が使えなければ従来の fetch」にする。この部品は失敗を隠さず、呼び出し側へ返す。
 * 🔴 中継先のGASは、外から呼べる関数を最小にした専用プロジェクトにすること（google.script.run は
 *    名前が _ で終わらない全関数を呼べるため、業務GAS本体に中継ページを置いてはいけない）。
 *
 * 使い方:
 *   const relay = createGasRelay({ url: '<中継GASの /exec>', fn: 'staffRead' });
 *   relay.start();                          // ページを開いたら早めに
 *   if (await relay.waitReady(3000)) { const text = await relay.call([payload], 30000); }
 */
(function (global) {
  'use strict';

  function createGasRelay(opts) {
    const readyTimeoutMs = opts.readyTimeoutMs || 15000;
    let state = 'idle';         // idle → loading → ready | failed
    let port = null, seq = 0, consecutiveFailures = 0;
    const pending = new Map();
    const waiters = [];

    function settle(ok) {
      if (state !== 'loading') return;
      state = ok ? 'ready' : 'failed';
      waiters.splice(0).forEach(function (w) { w(ok); });
    }

    function isRelayOrigin(origin) {
      try {
        const u = new URL(origin);
        return u.protocol === 'https:' && /\.googleusercontent\.com$/.test(u.hostname);
      } catch (e) { return false; }
    }

    function start() {
      if (state !== 'idle') return;
      state = 'loading';
      const onMessage = function (ev) {
        const d = ev.data;
        if (!isRelayOrigin(ev.origin) || !d || d.type !== 'gas-relay-ready' || !ev.ports || !ev.ports[0]) return;
        window.removeEventListener('message', onMessage);
        port = ev.ports[0];
        port.onmessage = function (e) {
          const m = e.data || {};
          const done = pending.get(m.id);
          if (done) { pending.delete(m.id); done(m); }
        };
        settle(true);
      };
      window.addEventListener('message', onMessage);
      const f = document.createElement('iframe');
      f.src = opts.url + '?mode=relay&origin=' + encodeURIComponent(location.origin);
      f.title = '通信用（表示なし）';
      f.setAttribute('aria-hidden', 'true');
      f.tabIndex = -1;
      f.style.cssText = 'position:absolute;width:1px;height:1px;left:-10px;top:-10px;border:0;opacity:0;pointer-events:none;';
      (document.body || document.documentElement).appendChild(f);
      setTimeout(function () {
        if (state === 'loading') { window.removeEventListener('message', onMessage); settle(false); }
      }, readyTimeoutMs);
    }

    // 準備完了を最大 ms 待つ。準備できなければ false（呼び出し側は fetch に切り替える）
    function waitReady(ms) {
      if (state === 'ready') return Promise.resolve(true);
      if (state !== 'loading') return Promise.resolve(false);
      return new Promise(function (resolve) {
        waiters.push(resolve);
        setTimeout(function () { resolve(state === 'ready'); }, ms);
      });
    }

    // 中継経由で呼ぶ。失敗（時間切れ・GAS側の例外）は reject。2回続けて失敗したら以後は使わない
    function call(args, timeoutMs) {
      return new Promise(function (resolve, reject) {
        if (state !== 'ready' || !port) { reject(new Error('RELAY_NOT_READY')); return; }
        const id = ++seq;
        const fail = function (code) {
          consecutiveFailures++;
          if (consecutiveFailures >= 2) state = 'failed';
          reject(new Error(code));
        };
        const timer = setTimeout(function () { pending.delete(id); fail('RELAY_TIMEOUT'); }, timeoutMs);
        pending.set(id, function (m) {
          clearTimeout(timer);
          if (m.ok) { consecutiveFailures = 0; resolve(m.result); }
          else fail('RELAY_FAILED');
        });
        port.postMessage({ type: 'call', id: id, fn: opts.fn, args: args });
      });
    }

    return {
      start: start, waitReady: waitReady, call: call,
      get state() { return state; }
    };
  }

  global.createGasRelay = createGasRelay;
})(window);
