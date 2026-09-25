# -*- coding: utf-8 -*-
"""社員用ページの「読み込み専用」中継GAS（beaufes-staff-read）を、本番GASのコードから生成する。

使い方（beaufes フォルダで）:
    py -3.12 tools/build_staff_read_gas.py <本番で稼働中の Code.gs> <版の説明（例: v0.39.0/版36）>

生成物: gas-staff-read/Code.gs ・ gas-staff-read/appsscript.json

なぜ別プロジェクトか（2026-09-25）:
  GASの fetch→/exec（ContentService）は結果の配送で失敗する（2026-09-24 A/B実測で10.6%）。
  google.script.run はこの失敗を受けないが、HtmlServiceページを置いたGASは、名前が _ で「終わらない」
  全関数を google.script.run から呼べてしまう（本番beaufesは127関数・メール送信やシート書き込みを含む）。
  そこで本番のコードを関数スコープに閉じ込めた別プロジェクトを作り、外から呼べる関数を
  doGet（中継ページ）と staffRead（ログイン必須の読み込み）の2つに限定する。

🔴 生成物を手で直さないこと。本番GASを更新したら、稼働中の版の Code.gs からこのスクリプトで作り直す。
🔴 書き込みはしない: validateSession の期限切れ行削除を外す。権限はスプレッドシートだけ（メール・外部通信なし）。
   ※ spreadsheets.readonly は SpreadsheetApp.openById が受け付けないため使えない（2026-09-25 実測で AUTH_TRANSIENT）。
   ロックは本番と共有されないため、予約の保存と重なった読み込みは staffRead が1回だけ読み直す。
"""
import hashlib, io, json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
OUT = os.path.join(BASE, "gas-staff-read")

# 中継ページを埋め込んでよい画面（通信路を渡す相手）
RELAY_ORIGINS = ["https://beaufield.github.io", "http://localhost:8791"]

# テスト（tests/staff_read_gas_check.js）が、この目印の間を元に戻して本番コードと同一か指紋で確かめる
BEGIN_MARK = "// ===== BEGIN 本番 Code.gs（期限切れ行削除の1行だけ差し替え） ====="
END_MARK = "// ===== END 本番 Code.gs ====="
DELETE_LINE = "sh.deleteRow(i + 1);"
READONLY_NOTE = "/* 中継GASは読み取り専用: 期限切れ行の削除は本番GAS・ポータルに任せる */"


def patch_validate_session(src):
    """validateSession の中の期限切れ行削除だけを外す（他の関数の deleteRow には触れない）。"""
    start = src.find("\nfunction validateSession(")
    if start < 0:
        raise SystemExit("validateSession が見つからない")
    end = src.find("\nfunction ", start + 10)
    body = src[start:end]
    if body.count(DELETE_LINE) != 1:
        raise SystemExit("validateSession の deleteRow が1か所ではない（元コードが変わった。確認すること）")
    return src[:start] + body.replace(DELETE_LINE, READONLY_NOTE) + src[end:]


RELAY_HTML = r"""<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body><script>
(function () {
  var PARENT = __ORIGIN__;
  if (!PARENT) return;
  var ch = new MessageChannel();
  ch.port1.onmessage = function (ev) {
    var m = ev.data || {};
    if (m.type !== 'call' || m.fn !== 'staffRead') { ch.port1.postMessage({ id: m.id, ok: false, error: 'FN_NOT_ALLOWED' }); return; }
    google.script.run
      .withSuccessHandler(function (r) { ch.port1.postMessage({ id: m.id, ok: true, result: r }); })
      .withFailureHandler(function (e) { ch.port1.postMessage({ id: m.id, ok: false, error: String(e && e.message || e) }); })
      .staffRead((m.args || [])[0]);
  };
  window.top.postMessage({ type: 'gas-relay-ready' }, PARENT, [ch.port2]);
})();
</script></body></html>"""


def build(source_path, label):
    src = io.open(source_path, encoding="utf-8").read().replace("\r\n", "\n")
    m = re.search(r"^const VERSION\s*=\s*'([^']+)'", src, re.M)
    if not m:
        raise SystemExit("VERSION が見つからない")
    if "\nfunction adminBootstrap(" not in src:
        raise SystemExit("adminBootstrap が見つからない（v0.39.0 以降の Code.gs を渡すこと）")
    sha = hashlib.sha256(src.encode("utf-8")).hexdigest()
    # 字下げはしない（複数行の文字列リテラルの中身が変わるため、元のコードをそのまま埋め込む）
    body = patch_validate_session(src)

    out = []
    out.append("// 🔴 生成物（tools/build_staff_read_gas.py）。手で編集しないこと。")
    out.append("// beaufes 社員用ページの「読み込み専用」中継GAS（beaufes-staff-read）。")
    out.append("//   元: 本番GAS " + label + "（VERSION " + m.group(1) + " / SHA256 " + sha[:16] + "）の Code.gs を関数スコープに閉じ込めたもの")
    out.append("//   外から呼べる関数は doGet（中継ページ）と staffRead（ログイン必須の読み込み）の2つだけ。")
    out.append("//   書き込みはしない（期限切れセッション行の削除を外す。読み込み処理に書き込みがないことはテストで確認。権限はスプレッドシートだけ）。")
    out.append("var STAFF_READ_BUILD_ = " + json.dumps({"source": label, "version": m.group(1), "sha256": sha[:16]}, ensure_ascii=False) + ";")
    out.append("var STAFF_READ_ORIGINS_ = " + json.dumps(RELAY_ORIGINS) + ";")
    out.append("")
    out.append("var BF_ = (function () {")
    out.append(BEGIN_MARK)
    out.append(body)
    out.append(END_MARK)
    out.append("  return { adminBootstrap: adminBootstrap };")
    out.append("})();")
    out.append("")
    out.append(r"""// 中継ページ。許可した画面にだけ通信路（MessagePort）を渡す。
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.mode !== 'relay') return ContentService.createTextOutput('');
  var origin = STAFF_READ_ORIGINS_.indexOf(p.origin) >= 0 ? p.origin : '';
  return HtmlService.createHtmlOutput(relayHtml_(origin))
    .setTitle('beaufes relay')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function relayHtml_(origin) {
  return __RELAY_HTML__.replace('__ORIGIN__', JSON.stringify(origin));
}

// 読み込み専用の入口。本番の doPost と同じ adminBootstrap を呼び、同じ形の応答（JSON文字列）を返す。
function staffRead(requestJson) {
  var req = null;
  try { req = JSON.parse(String(requestJson || '')); } catch (e) { req = null; }
  if (!req || req.action !== 'adminBootstrap' || !req.data || typeof req.data !== 'object') {
    return JSON.stringify({ success: false, error: 'INVALID_REQUEST' });
  }
  var data = {
    session_token: String(req.data.session_token || ''),
    include_applications: req.data.include_applications !== false,
    diagnostics: req.data.diagnostics === true
  };
  try {
    var res = BF_.adminBootstrap(data);
    // 本番とロックを共有していないため、予約の保存と重なって予約部分が読めなかったときだけ1回読み直す
    if (res && res.success && res.data && res.data.booking && res.data.booking.state === 'unavailable' &&
        /^(LOCK_BUSY|BOOKING_RECOVERY_REQUIRED)$/.test(String(res.data.booking.error || ''))) {
      Utilities.sleep(1500);
      res = BF_.adminBootstrap(data);
    }
    if (res && res.success && res.data) res.data.read_source = 'staff-read ' + STAFF_READ_BUILD_.version;
    return JSON.stringify(res);
  } catch (err) {
    return JSON.stringify({ success: false, error: 'INTERNAL_ERROR' });
  }
}
""".replace("__RELAY_HTML__", json.dumps(RELAY_HTML, ensure_ascii=False)))

    manifest = {
        "timeZone": "Etc/GMT-9",
        "runtimeVersion": "V8",
        "exceptionLogging": "STACKDRIVER",
        "oauthScopes": ["https://www.googleapis.com/auth/spreadsheets"],  # readonly は openById が受け付けない
        "webapp": {"executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS"},
    }
    if not os.path.isdir(OUT):
        os.makedirs(OUT)
    io.open(os.path.join(OUT, "Code.gs"), "w", encoding="utf-8", newline="\n").write("\n".join(out))
    io.open(os.path.join(OUT, "appsscript.json"), "w", encoding="utf-8", newline="\n").write(json.dumps(manifest, indent=2) + "\n")
    print("generated: gas-staff-read/Code.gs (" + label + ", VERSION " + m.group(1) + ", sha " + sha[:16] + ")")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    build(sys.argv[1], sys.argv[2])
