# -*- coding: utf-8 -*-
"""スマホから確認できるテスト環境（beaufes/test/）を、本番のHTMLから生成する。

使い方（beaufes フォルダで）:
    py -3.12 tools/build_test_pages.py                 # モックのみ（GASを使わない・完全に安全）
    py -3.12 tools/build_test_pages.py <テスト用GASのURL>   # テスト用GASに繋いだ本物の動作確認

生成物: beaufes/test/index.html / pass.html / liff.html
公開URL: https://beaufield.github.io/beaufield-dev/beaufes/test/

🔴 test/ の中身は**手で編集しない**。本番のHTMLを直してから、このスクリプトで作り直すこと。
🔴 このリポジトリは Public。テストページは誰でも開ける。個人情報を入れない・入れさせない作りにしてある:
   - 引数なしで生成した場合、送信は一切行わない（IS_MOCK を true に固定する）
   - 検索避け（noindex）と、画面上部の赤い「テスト環境」帯を必ず入れる
🔴 テスト用GASのURLを渡して生成したものを**そのままコミットしない**（URLが公開される）。
   その場合は生成 → 手元で確認 → `git checkout -- test/` で戻す、の順で使う。
"""
import io, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
OUT = os.path.join(BASE, "test")

TEST_GAS_URL = sys.argv[1].strip() if len(sys.argv) > 1 else ""
LIVE = bool(TEST_GAS_URL)

BANNER_CSS = """
<style>
  /* === テスト環境の帯（build_test_pages.py が挿入・本番には存在しない） === */
  #testEnvBar {
    position: sticky; top: 0; z-index: 999;
    background: #b91c1c; color: #fff;
    font-size: 12px; line-height: 1.6; text-align: center;
    padding: 8px 10px;
    font-family: "Hiragino Kaku Gothic ProN", "Yu Gothic", sans-serif;
  }
  #testEnvBar b { font-size: 13px; }
  #testEnvBar .navs { margin-top: 6px; }
  #testEnvBar a {
    color: #fff; text-decoration: underline; margin: 0 6px; white-space: nowrap;
  }
</style>
"""

BANNER_HTML = """
<div id="testEnvBar">
  <b>🔴 テスト環境（本番ではありません）</b><br>
  __MODE__
  <div class="navs">
    <a href="index.html">申込フォーム</a>
    <a href="pass.html?t=TESTTOKEN&amp;new=1">入場パス</a>
    <a href="liff.html">LINE版</a>
  </div>
</div>
"""

MODE_MOCK = "画面の確認用です。入力しても<b>どこにも送信されません</b>。"
MODE_LIVE = "テスト用のスプレッドシートに<b>実際に書き込みます</b>。本番の申込には影響しません。"

FILES = ["index.html", "pass.html", "liff.html"]


# 🔴 本番HTMLは 227a30c でセミナー予約UIが意図的に外されている（Pages と GAS の版が揃うまでの
#    暫定対応）。その状態のまま生成すると、テスト環境からUIが**黙って消える**。
#    UIを直すときは先に tools/restore_from_test_pages.py で戻すこと。
#    🔴 マーカーは3画面に共通のものだけにすること（pass.html の一覧は semList、index/liff は seminarList）。
REQUIRED_MARKERS = ["renderSeminars", "sessionChipHtml", "selectedSessionIds"]


def preflight():
    """🔴 1本でも書き出す前に3本すべてを検査する。
    途中で落ちるとテスト環境が「UIのある画面」と「無い画面」の混在になる。"""
    ng = []
    for name in FILES:
        src = io.open(os.path.join(BASE, name), encoding="utf-8").read()
        missing = [m for m in REQUIRED_MARKERS if m not in src]
        if missing:
            ng.append(name + "（" + ", ".join(missing) + " が無い）")
    if ng:
        raise SystemExit(
            "🔴 本番HTMLに予約UIがありません: " + " / ".join(ng) + "\n"
            "   本番HTMLは 227a30c でUIを外した状態がgitの正です。このまま生成すると\n"
            "   テスト環境からUIが消えます。先に次を実行してください:\n"
            "     py -3.12 tools/restore_from_test_pages.py\n"
            "   （何も書き出していません）")


def build(name):
    src = io.open(os.path.join(BASE, name), encoding="utf-8").read()

    # 1) 検索避け＋帯のCSS（<meta charset> の直後に入れる）
    anchor = '<meta charset="UTF-8">'
    if src.count(anchor) != 1:
        raise SystemExit("meta charset が見つからない: " + name)
    src = src.replace(
        anchor,
        anchor + '\n<meta name="robots" content="noindex,nofollow">' + BANNER_CSS.rstrip())

    # 2) 帯そのもの（<body> の直後）
    if src.count("<body>") != 1:
        raise SystemExit("<body> が見つからない: " + name)
    banner = BANNER_HTML.replace("__MODE__", MODE_LIVE if LIVE else MODE_MOCK)
    src = src.replace("<body>", "<body>" + banner.rstrip(), 1)

    # 3) 送信先
    src, n = re.subn(r"const GAS_URL\s* = '[^']*';",
                     "const GAS_URL = '" + (TEST_GAS_URL if LIVE else "TEST_ENV_NO_GAS") + "';",
                     src, count=1)
    if n != 1:
        raise SystemExit("GAS_URL が見つからない: " + name)

    # 4) モック固定（GASのURLを渡していないとき）
    if not LIVE:
        before = src
        src = src.replace(
            "const IS_MOCK = new URLSearchParams(location.search).get('mock') === '1';",
            "const IS_MOCK = true;   // 🔴 テスト環境では常にモック（送信しない）")
        src = src.replace(
            "const IS_MOCK = qs('mock') === '1';",
            "const IS_MOCK = true;   // 🔴 テスト環境では常にモック（送信しない）")
        if src == before:
            raise SystemExit("IS_MOCK が見つからない: " + name)

    # 5) 1つ上のフォルダにあるライブラリを参照する
    src = src.replace('src="lib/', 'src="../lib/')

    # 6) 生成物であることを明記
    src = src.replace(
        "<!DOCTYPE html>",
        "<!DOCTYPE html>\n<!-- 🔴 このファイルは tools/build_test_pages.py の生成物です。手で編集しないこと。\n"
        "     直すときは beaufes/" + name + " を編集してから再生成してください。 -->", 1)

    if not os.path.isdir(OUT):
        os.makedirs(OUT)
    io.open(os.path.join(OUT, name), "w", encoding="utf-8", newline="").write(src)
    print("generated: test/" + name + ("  [LIVE]" if LIVE else "  [mock only]"))


preflight()
for f in FILES:
    build(f)

print("\n公開URL: https://beaufield.github.io/beaufield-dev/beaufes/test/index.html")
if LIVE:
    print("🔴 テスト用GASのURLが埋め込まれています。この状態でコミットしないこと。")
