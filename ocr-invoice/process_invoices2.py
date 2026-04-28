import csv
import os

master_path = r"D:\Dropbox\ClaudeWork\開発・自動化\ocr-invoice\商品_utf8.csv"
output_dir = r"D:\Dropbox\ClaudeWork\開発・自動化\ocr-invoice\results"

jan_to_item = {}
maker_code_to_item = {}
n_code_to_item = {}

with open(master_path, encoding='utf-8') as f:
    reader = csv.reader(f)
    next(reader)
    for row in reader:
        if len(row) < 3:
            continue
        code = row[0].strip()
        n_code = row[1].strip() if len(row) > 1 else ''
        name = row[2].strip() if len(row) > 2 else ''
        jan = row[26].strip() if len(row) > 26 else ''
        maker_code = row[27].strip() if len(row) > 27 else ''

        if jan and jan not in jan_to_item:
            jan_to_item[jan] = (code, name)
        if maker_code and maker_code not in maker_code_to_item:
            maker_code_to_item[maker_code] = (code, name)
        if n_code.startswith('N') and len(n_code) > 1:
            num = n_code[1:]
            if num not in n_code_to_item:
                n_code_to_item[num] = (code, name)

def lookup_jan(jan, ocr_name):
    if jan in jan_to_item:
        return jan_to_item[jan]
    return ('1', ocr_name)

def lookup_napla(napla_code, ocr_name):
    code_str = str(napla_code)
    if code_str in maker_code_to_item:
        return maker_code_to_item[code_str]
    if code_str in n_code_to_item:
        return n_code_to_item[code_str]
    return ('1', ocr_name)

def write_csv(filename, rows):
    path = os.path.join(output_dir, filename + '.csv')
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['コード', '商品名', '数量', '単価', '合計'])
        for bcode, bname, qty, price, total in rows:
            writer.writerow([bcode, bname, qty, price, total])
    matched = sum(1 for r in rows if r[0] != '1')
    unmatched = sum(1 for r in rows if r[0] == '1')
    print(f"OK: {filename}.csv  照合済{matched}件 / 未照合{unmatched}件")
    for r in rows:
        if r[0] == '1':
            print(f"  [未照合] {r[1]}")

# ============================================================
# ナプラ 伝票 0003995254 (2026-03-19)
# 14544 SHEAミルク, 14543 SHEAオイル
# 98458 パンフレット → 除外
# ============================================================
rows = []
for nc, oname, qty, price, total in [
    ('14544', 'エヌドット SHEAミルク 150g', 4, 1185, 4740),
    ('14543', 'エヌドット SHEAオイル 150mL', 3, 1185, 3555),
]:
    c, n = lookup_napla(nc, oname)
    rows.append((c, n, qty, price, total))
write_csv('ナプラ_20260319_0003995254', rows)

# ============================================================
# ナプラ 伝票 0003995397 (2026-03-19)
# ============================================================
rows = []
for nc, oname, qty, price, total in [
    ('12067', 'N.スタイリングフォーム ルーズカール 200g', 2, 910, 1820),
    ('12068', 'N.スタイリングフォーム バウンスウェーブ 200g', 3, 910, 2730),
    ('12066', 'N.ナリッシングオイル 150mL', 2, 1185, 2370),
]:
    c, n = lookup_napla(nc, oname)
    rows.append((c, n, qty, price, total))
write_csv('ナプラ_20260319_0003995397', rows)

# ============================================================
# ナプラ 伝票 0003995405 (2026-03-19)
# 98593 リーフレット・99996 運賃 → 除外
# ============================================================
rows = []
for nc, oname, qty, price, total in [
    ('12068', 'N.スタイリングフォーム バウンスウェーブ 200g', 5, 910, 4550),
]:
    c, n = lookup_napla(nc, oname)
    rows.append((c, n, qty, price, total))
write_csv('ナプラ_20260319_0003995405', rows)

# ============================================================
# ナプラ 伝票 0003996764 (2026-03-19)
# ============================================================
rows = []
for nc, oname, qty, price, total in [
    ('14527', 'N.カラートリートメント Si(シルバー) 300g', 2, 1092, 2184),
]:
    c, n = lookup_napla(nc, oname)
    rows.append((c, n, qty, price, total))
write_csv('ナプラ_20260319_0003996764', rows)

# ============================================================
# ナプラ 伝票 0003997391 (2026-03-19)
# ============================================================
rows = []
for nc, oname, qty, price, total in [
    ('16980', 'エヌドットカラー C-6SB 80g', 6, 423, 2538),
]:
    c, n = lookup_napla(nc, oname)
    rows.append((c, n, qty, price, total))
write_csv('ナプラ_20260319_0003997391', rows)

# ============================================================
# ナプラ 伝票 0003997392 (2026-03-19)
# ============================================================
rows = []
for nc, oname, qty, price, total in [
    ('16813', 'ナシードカラー QN-5AB 80g', 3, 390, 1170),
    ('16827', 'ナシードカラー QN-8RPB 80g', 3, 390, 1170),
    ('16849', 'エヌドットカラー G-5NB 80g', 6, 423, 2538),
    ('16851', 'エヌドットカラー G-7NB 80g', 6, 423, 2538),
    ('16855', 'エヌドットカラー G-6AB 80g', 6, 423, 2538),
    ('16865', 'エヌドットカラー G-8BB 80g', 6, 423, 2538),
    ('16982', 'エヌドットカラー G-8SB 80g', 6, 423, 2538),
    ('16985', 'エヌドットカラー G-60LB 80g', 6, 423, 2538),
]:
    c, n = lookup_napla(nc, oname)
    rows.append((c, n, qty, price, total))
write_csv('ナプラ_20260319_0003997392', rows)

# ============================================================
# 滝川 伝票 075193 (2026-03-20)
# MFR07 マニフィーウィッグ F2B
# ============================================================
rows = []
c, n = lookup_jan('4991560809768', 'MFR07 マニフィーウィッグ F2B')
rows.append((c, n, 1, 44500, 44500))
write_csv('滝川_20260320_075193', rows)

# ============================================================
# 滝川 伝票 075194 (2026-03-20)
# F04 ビィラフラ R174
# ============================================================
rows = []
c, n = lookup_jan('4991560834074', 'F04 ビィラフラ R174')
rows.append((c, n, 1, 20000, 20000))
write_csv('滝川_20260320_075194', rows)

print("\n全処理完了")
