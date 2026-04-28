"""
results/ フォルダの簡易CSVを、販売管理システム取込フォーマット（21列）に変換する。
出力先: results/取込用/（Shift-JIS形式）

列定義（仕様書準拠）:
 1: 伝票No        ← ファイル名の伝票番号部分
 2: 仕入先CD      ← 商品マスターの「仕入先CD」列（照合済み商品から取得）
 3: 仕入先子CD    ← 固定: 0
 4: 手書伝票No    ← 空欄（システム自動セット）
 5: 仕入日        ← ファイル名の日付部分（YYYY/MM/DD）
 6: 入荷日        ← 仕入日と同じ
 7: 仕入区分      ← 固定: 0（掛仕入）
 8: 担当者        ← 空欄
 9: 運送会社      ← 空欄
10: 明細区分      ← 固定: 1（仕入）
11: 商品CD        ← コード（未照合は1）
12: 商品名        ← 商品名
13: 税コード      ← 固定: 1（標準税率）※将来的に商品マスター列追加で対応
14: 大分類        ← 商品コードがある場合は空欄、コード1は900
15: 中分類        ← 商品コードがある場合は空欄、コード1は90
16: 小分類        ← 商品コードがある場合は空欄、コード1は950
17: 数量          ← OCR読み取り値
18: 単位          ← 商品マスターの「単位名」列
19: 特値          ← 固定: 1（通常）
20: 仕入単価      ← OCR読み取り値
21: 仕入金額      ← OCR読み取り値
"""

import csv
import os
import re
import sys

# パス設定
script_dir  = os.path.dirname(os.path.abspath(__file__))
results_dir = os.path.join(script_dir, 'results')
output_dir  = os.path.join(results_dir, '取込用')
master_path = os.path.join(script_dir, '商品_utf8.csv')

os.makedirs(output_dir, exist_ok=True)

# ── 商品マスター読み込み ──────────────────────────────────
master = {}  # コード → dict
with open(master_path, encoding='utf-8-sig') as f:
    reader = csv.DictReader(f)
    for row in reader:
        code = row['コード'].strip()
        master[code] = {
            '仕入先CD': row['仕入先CD'].strip(),
            '単位名':   row['単位名'].strip(),
            '大分類':   row['大分類'].strip(),
            '中分類':   row['中分類'].strip(),
            '小分類':   row['小分類'].strip(),
        }

# ── 出力ヘッダー ──────────────────────────────────────────
OUTPUT_HEADER = [
    '伝票No', '仕入先CD', '仕入先子CD', '手書伝票No',
    '仕入日', '入荷日', '仕入区分', '担当者', '運送会社',
    '明細区分', '商品CD', '商品名', '税コード',
    '大分類', '中分類', '小分類',
    '数量', '単位', '特値', '仕入単価', '仕入金額'
]

# ── 各CSVを変換 ───────────────────────────────────────────
ok_count  = 0
skip_count = 0

for fname in sorted(os.listdir(results_dir)):
    if not fname.endswith('.csv'):
        continue
    if fname.startswith('.') or fname.startswith('_'):
        continue

    # ── ファイル名解析: メーカー名_YYYYMMDD_伝票番号.csv ──
    base  = fname[:-4]
    parts = base.split('_')

    date_part = None
    date_idx  = None
    for i, p in enumerate(parts):
        if re.match(r'^\d{8}$', p):
            date_part = p
            date_idx  = i
            break

    if not date_part:
        print(f'  スキップ（日付不明）: {fname}')
        skip_count += 1
        continue

    # 仕入日 YYYYMMDD → YYYY/MM/DD
    shiire_date = f'{date_part[:4]}/{date_part[4:6]}/{date_part[6:8]}'

    # 伝票No（日付より後の部分を結合）
    denpin_no = '_'.join(parts[date_idx + 1:])
    if not denpin_no:
        denpin_no = base  # フォールバック

    # ── 元CSV読み込み ──────────────────────────────────────
    input_path = os.path.join(results_dir, fname)
    rows = []
    with open(input_path, encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)

    if not rows:
        skip_count += 1
        continue

    # 仕入先CD: 照合済み商品のマスターから取得
    shiire_cd = ''
    for row in rows:
        code = row['コード'].strip()
        if code not in ('', '1') and code in master:
            shiire_cd = master[code]['仕入先CD']
            if shiire_cd:
                break

    # ── 取込用CSV出力（Shift-JIS） ─────────────────────────
    output_path = os.path.join(output_dir, fname)
    with open(output_path, 'w', encoding='cp932', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(OUTPUT_HEADER)

        for row in rows:
            code  = (row.get('コード')  or '').strip()
            name  = (row.get('商品名')  or '').strip()
            qty   = (row.get('数量')    or '').strip()
            price = (row.get('単価')    or '').strip()
            total = (row.get('合計')    or '').strip()

            # 空行スキップ
            if not code and not name:
                continue

            # 商品マスターから単位・分類を取得
            if code not in ('', '1') and code in master:
                m     = master[code]
                unit  = m['単位名']
                daibu = ''     # コードあり→システムが自動セット
                chubu = ''
                shobu = ''
            else:
                # コード1（未照合）
                unit  = ''
                daibu = '900'
                chubu = '90'
                shobu = '950'

            writer.writerow([
                denpin_no,    #  1: 伝票No
                shiire_cd,    #  2: 仕入先CD
                '0',          #  3: 仕入先子CD
                '',           #  4: 手書伝票No
                shiire_date,  #  5: 仕入日
                shiire_date,  #  6: 入荷日
                '0',          #  7: 仕入区分（0:掛仕入）
                '',           #  8: 担当者
                '',           #  9: 運送会社
                '1',          # 10: 明細区分（1:仕入）
                code,         # 11: 商品CD
                name,         # 12: 商品名
                '1',          # 13: 税コード（A案:標準税率固定）
                daibu,        # 14: 大分類
                chubu,        # 15: 中分類
                shobu,        # 16: 小分類
                qty,          # 17: 数量
                unit,         # 18: 単位
                '1',          # 19: 特値（1:通常）
                price,        # 20: 仕入単価
                total,        # 21: 仕入金額
            ])

    ok_count += 1
    print(f'  OK: {fname}')

print()
print(f'変換完了: {ok_count}件 / スキップ: {skip_count}件')
print(f'出力先: {output_dir}')
