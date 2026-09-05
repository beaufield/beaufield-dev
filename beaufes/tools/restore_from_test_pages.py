# -*- coding: utf-8 -*-
"""コミット済みの beaufes/test/*.html から、本番HTML（index/pass/liff）を復元する。

`build_test_pages.py` の逆変換。往復で1文字も変わらないことを検証済み（2026-09-05）。

使い方（beaufes フォルダで）:
    py -3.12 tools/restore_from_test_pages.py

## なぜこれが必要か（🔴 読まずに使わないこと）

`227a30c` で、公開中の本番HTML3本から**セミナー予約のUIが意図的に外された**。
GitHub Pages は push した瞬間に公開されるのに、本番GASは `@19`(v0.18.0) のままで
`listSessions` を持たないため、お客様の画面に「空き状況を確認できませんでした」が
出続けていたことへの暫定対応。

その結果、**本番HTMLはUIを持たない状態がgitの正**になっている。
一方 `build_test_pages.py` は本番HTMLを入力にするので、そのままではテスト環境を
作り直せない。そこで:

  1. このスクリプトで本番HTMLにUIを戻す（作業ツリーのみ）
  2. UIを直す
  3. `build_test_pages.py` でテスト環境を作り直す
  4. 🔴 **本番HTML3本は `git checkout -- beaufes/index.html beaufes/pass.html beaufes/liff.html`
     で必ずHEADへ戻す**（コミットすると本番の申込フォームにUIが出てしまう）

本番へ出すときは、この復元をしたうえで3本ごとコミットし、同時にGASも `v0.24.0` 以降へ
切り替える（Pages と GAS の版が揃うまでUIを出さない、が `227a30c` の判断）。
"""
import io, os, sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

IS_MOCK_ORIG = {
    "index.html": "const IS_MOCK = new URLSearchParams(location.search).get('mock') === '1';",
    "liff.html":  "const IS_MOCK = new URLSearchParams(location.search).get('mock') === '1';",
    "pass.html":  "const IS_MOCK = qs('mock') === '1';",
}
IS_MOCK_TEST = "const IS_MOCK = true;   // 🔴 テスト環境では常にモック（送信しない）"


def gas_url_line(name):
    """いま作業ツリーにある本番ファイルから GAS_URL の行をそのまま持ってくる。
    🔴 テスト用GASのURLで生成した test/ から復元すると本番URLが失われるので、
       送信先だけは必ず本番ファイル側を正とする。"""
    cur = io.open(os.path.join(BASE, name), encoding="utf-8").read()
    for ln in cur.split("\n"):
        if ln.startswith("const GAS_URL"):
            if "TEST_ENV_NO_GAS" in ln or "script.google.com" not in ln:
                raise SystemExit(
                    "本番ファイルの GAS_URL が本番のものではありません: " + name +
                    "\n（テスト用URLで上書きされたまま復元すると送信先が壊れます）")
            return ln
    raise SystemExit("GAS_URL 行が見つからない: " + name)


def restore(name):
    src = io.open(os.path.join(BASE, "test", name), encoding="utf-8").read()

    mark = ("\n<!-- 🔴 このファイルは tools/build_test_pages.py の生成物です。手で編集しないこと。\n"
            "     直すときは beaufes/" + name + " を編集してから再生成してください。 -->")
    if src.count(mark) != 1:
        raise SystemExit("生成物コメントが見つからない: " + name)
    src = src.replace(mark, "")

    a = src.find('\n<meta name="robots" content="noindex,nofollow">')
    if a < 0:
        raise SystemExit("noindex が見つからない: " + name)
    b = src.find("</style>", a)
    if b < 0:
        raise SystemExit("帯CSSの終端が見つからない: " + name)
    src = src[:a] + src[b + len("</style>"):]

    a = src.find('\n<div id="testEnvBar">')
    if a < 0:
        raise SystemExit("帯が見つからない: " + name)
    b = src.find("\n</div>", a)
    if b < 0:
        raise SystemExit("帯の終端が見つからない: " + name)
    src = src[:a] + src[b + len("\n</div>"):]

    if src.count("const GAS_URL = 'TEST_ENV_NO_GAS';") != 1:
        raise SystemExit("TEST_ENV_NO_GAS が見つからない（テスト用GASで生成した test/ からは復元できません）: " + name)
    src = src.replace("const GAS_URL = 'TEST_ENV_NO_GAS';", gas_url_line(name))

    if src.count(IS_MOCK_TEST) != 1:
        raise SystemExit("IS_MOCK(テスト用) が見つからない: " + name)
    src = src.replace(IS_MOCK_TEST, IS_MOCK_ORIG[name])

    src = src.replace('src="../lib/', 'src="lib/')

    for bad in ("testEnvBar", "TEST_ENV_NO_GAS", "noindex", "../lib/", "build_test_pages"):
        if bad in src:
            raise SystemExit("テスト環境の痕跡が残っている（%s）: %s" % (bad, name))

    io.open(os.path.join(BASE, name), "w", encoding="utf-8", newline="").write(src)
    print("restored: " + name)


for f in ("index.html", "pass.html", "liff.html"):
    restore(f)

# 🔴 コンソールが cp932 だと絵文字で UnicodeEncodeError になる。
#    この案内は必ず出したいので、出力には絵文字を使わない。
print("")
print("[!] 作業が終わったら、本番HTML3本は必ずHEADへ戻すこと:")
print("   git checkout -- beaufes/index.html beaufes/pass.html beaufes/liff.html")
