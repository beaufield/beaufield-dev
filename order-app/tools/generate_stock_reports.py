# -*- coding: utf-8 -*-
"""
死蔵在庫・過剰在庫 HTMLレポート生成
====================================
2026-08-02 に作成した死蔵在庫レポート／過剰在庫レポート（Claude Artifactデザイン）を
テンプレートとして、最新の dead_stock_*.csv / excess_stock_*.csv からレポートを作り直す。
CSS・JSのロジックは一切変更せず、タイトルの基準日と埋め込みJSONデータだけを差し替える。

■ 使い方
    py -3.12 tools/generate_stock_reports.py [--date YYYYMMDD]
    （order-app のプロジェクトルートで実行。省略時は tools/output 内の最新CSVの日付を使う）

出力:
    tools/output/死蔵在庫レポート_{YYYYMMDD}.html
    tools/output/過剰在庫レポート_{YYYYMMDD}.html

※ どちらも在庫金額・商品名等の業務数値を含むため Git コミット禁止（.gitignore登録済み）。
"""
import argparse
import glob
import json
import re

import pandas as pd

OUTPUT_DIR = 'tools/output'
DEAD_TEMPLATE = f'{OUTPUT_DIR}/死蔵在庫レポート_20260802.html'
EXCESS_TEMPLATE = f'{OUTPUT_DIR}/過剰在庫レポート_20260802.html'
MAKER_TOPITEMS_CAP = 12
TOPITEMS_OVERALL_CAP = 50


def latest_date_stamp(prefix):
    files = glob.glob(f'{OUTPUT_DIR}/{prefix}_*.csv')
    if not files:
        raise SystemExit(f'{prefix}_*.csv が {OUTPUT_DIR} に見つかりません')
    stamps = [re.search(r'(\d{8})', f).group(1) for f in files]
    return max(stamps)


def build_dead_items(stamp):
    path = f'{OUTPUT_DIR}/dead_stock_{stamp}.csv'
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    items = []
    for _, r in df.iterrows():
        months = r['monthsSinceLastSale']
        items.append({
            'code': r['code'],
            'name': r['name'],
            'maker': r['supplierName'],
            'stock': float(r['stock']),
            'unitCost': float(r['unitCost']),
            'amount': float(r['deadAmount']),
            'lastSaleDate': r['lastSaleDate'],
            'monthsSinceLastSale': float(months) if months not in ('', 'nan') else None,
            'tier': r['tier'],
            'reason': r['reason'],
        })
    items.sort(key=lambda i: -i['amount'])
    return items


def build_excess(stamp):
    path = f'{OUTPUT_DIR}/excess_stock_{stamp}.csv'
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    items = []
    for _, r in df.iterrows():
        items.append({
            'code': r['code'],
            'name': r['name'],
            'supplierName': r['supplierName'],
            'stock': float(r['stock']),
            'recommended': float(r['recommended']),
            'excessQty': float(r['excessQty']),
            'unitCost': float(r['unitCost']),
            'excessAmount': float(r['excessAmount']),
            'monthsOfStock': float(r['monthsOfStock']),
            'abcRank': r['abcRank'],
            'pattern': r['pattern'],
        })
    items.sort(key=lambda i: -i['excessAmount'])

    total = {
        'count': len(items),
        'qty': sum(i['stock'] for i in items),
        'amount': sum(i['excessAmount'] for i in items),
    }

    makers = {}
    for i in items:
        m = makers.setdefault(i['supplierName'], {
            'maker': i['supplierName'], 'count': 0, 'qty': 0.0, 'amount': 0.0, 'items': [],
        })
        m['count'] += 1
        m['qty'] += i['stock']
        m['amount'] += i['excessAmount']
        m['items'].append(i)

    maker_list = []
    for m in makers.values():
        m['items'].sort(key=lambda i: -i['excessAmount'])
        top_items = []
        for it in m['items'][:MAKER_TOPITEMS_CAP]:
            it2 = dict(it)
            it2.pop('supplierName', None)
            top_items.append(it2)
        maker_list.append({
            'maker': m['maker'], 'count': m['count'], 'qty': m['qty'], 'amount': m['amount'],
            'topItems': top_items,
        })
    maker_list.sort(key=lambda m: -m['amount'])

    maker_data = {'total': total, 'makerCount': len(maker_list), 'makers': maker_list}
    top_items_overall = items[:TOPITEMS_OVERALL_CAP]
    return maker_data, top_items_overall


def render(template_path, out_path, title_date, replacements):
    text = open(template_path, encoding='utf-8').read()
    text = re.sub(r'\d{4}-\d{2}-\d{2}(?=時点）</title>)', title_date, text, count=1)
    text = re.sub(r'(基準日: )\d{4}-\d{2}-\d{2}', r'\g<1>' + title_date, text, count=1)
    for elem_id, payload in replacements.items():
        pattern = re.compile(
            r'(<script id="%s" type="application/json">).*?(</script>)' % re.escape(elem_id),
            re.DOTALL,
        )
        json_text = json.dumps(payload, ensure_ascii=False)
        text, n = pattern.subn(lambda m, j=json_text: m.group(1) + j + m.group(2), text, count=1)
        if n != 1:
            raise SystemExit(f'テンプレート内に id="{elem_id}" の script タグが見つかりません: {template_path}')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(text)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--date', help='YYYYMMDD。省略時は最新のCSVを自動検出（死蔵・過剰それぞれ）')
    args = ap.parse_args()

    dead_stamp = args.date or latest_date_stamp('dead_stock')
    excess_stamp = args.date or latest_date_stamp('excess_stock')

    dead_title_date = f'{dead_stamp[0:4]}-{dead_stamp[4:6]}-{dead_stamp[6:8]}'
    dead_items = build_dead_items(dead_stamp)
    dead_out = f'{OUTPUT_DIR}/死蔵在庫レポート_{dead_stamp}.html'
    render(DEAD_TEMPLATE, dead_out, dead_title_date, {'allItemsData': dead_items})
    n_t1 = sum(1 for i in dead_items if i['tier'] == '完全死蔵')
    n_t2 = len(dead_items) - n_t1
    amt = sum(i['amount'] for i in dead_items)
    print(f'死蔵在庫レポート: {len(dead_items)}件 / {amt:,.0f}円（完全死蔵{n_t1}件・休眠{n_t2}件） -> {dead_out}')

    excess_title_date = f'{excess_stamp[0:4]}-{excess_stamp[4:6]}-{excess_stamp[6:8]}'
    maker_data, top_items = build_excess(excess_stamp)
    excess_out = f'{OUTPUT_DIR}/過剰在庫レポート_{excess_stamp}.html'
    render(EXCESS_TEMPLATE, excess_out, excess_title_date, {
        'makerData': maker_data,
        'topItemsData': top_items,
    })
    print(f"過剰在庫レポート: {maker_data['total']['count']}件 / {maker_data['total']['amount']:,.0f}円 -> {excess_out}")


if __name__ == '__main__':
    main()
