# ============================================================
# Beaufield 需要パターン分析・発注提案スクリプト
# Version: v1.4.0
#
# 概要:
#   売上データ明細表.CSV（過去24ヶ月）を分析し、商品ごとに
#   需要パターンを分類して推奨在庫・発注提案を算出する。
#   結果は GAS の「発注提案」シートへ書き込み（→アプリの発注提案タブに表示、
#   LINE WORKS 通知はGAS側が担当）、ローカルにも CSV / JSON を出力する。
#
# 需要パターン分類（Syntetos-Boylan 分類を月次に適用）:
#   安定型      : ほぼ毎月出て、量のばらつきが小さい
#   変動型      : ほぼ毎月出るが、量のばらつきが大きい
#   まとめ買い型 : 出る月が飛び飛びで、出るときはまとまって出る
#   散発型      : 出る月が飛び飛びだが、出る量は一定
#
# ABCランク（v1.3.0で追加。全商品一律だったサービス水準に濃淡をつける）:
#   月次原価貢献（月平均需要×仕入単価）の降順・累積構成比で分類
#   A: 累積80%まで（主力・欠品許容5%＝現行水準を維持）
#   B: 80〜95%（欠品許容10%）
#   C: 残り5%（欠品許容20%・ロングテール）
#   Cランクの散発型・まとめ買い型は提案そのものを出さず「受注発注推奨」扱いにする
#
# 推奨在庫の考え方:
#   共通: 「リードタイム＋発注サイクル」期間の平均需要 ＋ 安全在庫（ランク別Z値）
#   まとめ買い型・散発型: 上記と「1回のまとまり注文サイズ(P95)」の大きい方
#     （Aランクはフル適用・Bランクは月数キャップ付き・Cランクは適用なし）
#   全ランク共通で「月平均needs×CAP_MONTHS_GLOBALヶ月分」を上限キャップ
#   リードタイム・発注サイクルは発注先マスターのF/G列（未設定は各7日）
#
# 発注提案の対象:
#   月平均1個以上・販売歴6ヶ月以上・提案除外設定に無い・Cランクの間欠需要でない商品のうち、
#   「現在庫＋発注済み未入荷分」が推奨在庫を下回ったもの。
#   発注済み未入荷分 = リードタイム内に発注アプリから発注した数量（二重発注防止）
#   提案数量は推定ロット単位に切り上げ。
#   推定ロット = 過去の発注明細の数量の最大公約数（発注3回以上かつ2以上のときのみ採用）
#   提案金額 = 仕入単価（商品.CSV）× 提案数量。仕入単価未設定（0円）の商品は金額0円扱い
#   （アプリ側で「—」表示。合計からは除外しない）
#
# 過剰在庫（v1.4.0で追加）:
#   現在庫が「推奨在庫のEXCESS_RATIO倍」を超え、かつ超過数量がEXCESS_MIN_QTY以上の商品を
#   「過剰在庫」シートへ書き込む（提案とは別枠）。アプリの提案タブで金額降順に表示し、
#   意図的なまとめ仕入等で問題ない場合は「確認済み」登録すると次回分析後もリストから消える
#
# 実行方法:
#   py -3.12 analyze_demand.py             # 分析＋GASへ書き込み＋通知
#   py -3.12 analyze_demand.py --dry-run   # 分析のみ（ローカル出力・新旧比較レポートのみ・GAS送信なし）
#
# 設定ファイル: 同フォルダの config.json（gitignore済み・Dropbox管理）
#   gas_url / api_key は calc_reorder_point.py と共通
#
# Windowsタスクスケジューラ（推奨: 毎週月曜 07:00）
# ============================================================

import argparse
import json
import logging
import math
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import requests

SCRIPT_DIR  = Path(__file__).parent
CONFIG_PATH = SCRIPT_DIR / 'config.json'
OUTPUT_DIR  = SCRIPT_DIR / 'output'
LOG_DIR     = SCRIPT_DIR / 'logs'

# ---- 分析パラメータ ----
WINDOW_MONTHS    = 24     # 分析対象期間（完全月）
ADI_THRESHOLD    = 1.32   # 需要間隔のしきい値（これ以上=間欠的）
CV2_THRESHOLD    = 0.49   # 需要量ばらつきのしきい値（これ以上=ばらつき大）
MIN_MONTHS_DATA  = 6      # これ未満の販売歴しかない商品は「データ不足」扱い
MIN_MEAN_MONTHLY = 1.0    # 提案対象の最低月平均（これ未満は受注発注品とみなし提案しない）
DEFAULT_LEAD_TIME_DAYS   = 7
DEFAULT_ORDER_CYCLE_DAYS = 7
LOT_MIN_ORDERS   = 3      # ロット推定に必要な最低発注回数
DAYS_PER_MONTH   = 30.4

# ---- ABCランク・在庫水準差別化パラメータ（v1.3.0） ----
SERVICE_Z_LEGACY   = 1.65                                   # 旧ロジック（ランク差別化なし）。新旧比較レポート専用
SERVICE_Z_BY_CLASS = {'A': 1.65, 'B': 1.28, 'C': 0.84}       # 安全在庫係数（欠品許容 5/10/20%）
SERVICE_ALLOWANCE_PCT = {'A': 5, 'B': 10, 'C': 20}           # 根拠メモ表示用（Z値と対応）
ABC_A_CUM        = 0.80   # 月次原価貢献の累積構成比しきい値（ここまでA）
ABC_B_CUM        = 0.95   # 同上（ここまでB。超えたらC）
CAP_MONTHS_B     = 3.0    # BランクのまとめP95フロア上限（月数）
CAP_MONTHS_GLOBAL = 6.0   # 全ランク共通の推奨在庫上限（月平均の何ヶ月分まで許すか）

# ---- 過剰在庫検出パラメータ（v1.4.0） ----
EXCESS_RATIO   = 2.0   # 過剰判定: 現在庫が推奨在庫の何倍を超えたら対象にするか
EXCESS_MIN_QTY = 3     # 過剰判定: 超過数量がこれ未満なら対象外（少額商品のノイズ除外）

# ---- CSVパスの候補（デバイスによりドライブ構成が異なる） ----
_ONEDRIVE_DATA_DIRS = [
    Path(r'D:\OneDrive - Beaufield\PowerBI\Data'),
    Path.home() / 'OneDrive - Beaufield' / 'PowerBI' / 'Data',
]
SALES_CSV_CANDIDATES   = [str(d / '売上データ明細表.CSV') for d in _ONEDRIVE_DATA_DIRS]
PRODUCT_CSV_CANDIDATES = [str(d / '商品.CSV') for d in _ONEDRIVE_DATA_DIRS]


def setup_logger():
    LOG_DIR.mkdir(exist_ok=True)
    log_file = LOG_DIR / f"analyze_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    # cp932コンソールでも絵文字入りログで落ちないようにする
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(errors='replace')
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[logging.FileHandler(log_file, encoding='utf-8'),
                  logging.StreamHandler(sys.stdout)])
    return log_file


def load_config():
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(f'設定ファイルが見つかりません: {CONFIG_PATH}')
    with open(CONFIG_PATH, encoding='utf-8') as f:
        return json.load(f)


def find_csv(explicit, candidates, label):
    paths = ([explicit] if explicit else []) + candidates
    for p in paths:
        if p and Path(p).exists():
            return Path(p)
    raise FileNotFoundError(f'{label} が見つかりません。--csv / --products で指定してください。')


def calc_period():
    """当月を除く完全 WINDOW_MONTHS ヶ月の期間を返す"""
    today = date.today()
    end_date = date(today.year, today.month, 1) - timedelta(days=1)  # 先月末
    m = today.month - WINDOW_MONTHS
    y = today.year
    while m <= 0:
        m += 12
        y -= 1
    return date(y, m, 1), end_date


def month_range(start_date, end_date):
    months = []
    y, m = start_date.year, start_date.month
    while (y, m) <= (end_date.year, end_date.month):
        months.append(f'{y:04d}{m:02d}')
        m += 1
        if m > 12:
            m = 1
            y += 1
    return months


def normalize_code(c, width=6):
    try:
        return str(int(str(c).strip())).zfill(width)
    except (ValueError, TypeError):
        return None


def normalize_supplier_code(c):
    """仕入先コードの表記ゆれ吸収（'  63' / '063' / '63' → '63'）"""
    try:
        return str(int(str(c).strip()))
    except (ValueError, TypeError):
        return None


def load_sales(csv_path, start_str, end_str):
    """売上明細CSVを読み込み、期間フィルタ・型変換して返す
    列: 0=売上日, 1=売上№(取引単位のキー), 24=商品コード, 26=数量
    ※33列目の「伝票No」はほぼ空欄のため使わない
    """
    logging.info(f'売上CSV読み込み開始: {csv_path}')
    t0 = datetime.now()
    df = pd.read_csv(
        csv_path,
        encoding='cp932',
        header=0,
        usecols=[0, 1, 24, 26],
        names=['date', 'slip', 'code', 'qty'],
        dtype=str,
        on_bad_lines='skip',
    )
    logging.info(f'読み込み完了: {len(df):,}行 ({(datetime.now()-t0).total_seconds():.1f}秒)')

    df['date'] = df['date'].fillna('').str.strip()
    df = df[(df['date'] >= start_str) & (df['date'] <= end_str)]

    df['code'] = df['code'].map(normalize_code)
    df = df[df['code'].notna()]

    df['qty'] = pd.to_numeric(
        df['qty'].fillna('0').str.replace(',', '', regex=False), errors='coerce')
    df = df[df['qty'].notna()]

    df['slip'] = df['slip'].fillna('').str.strip()
    df['ym'] = df['date'].str[:6]
    logging.info(f'期間内の有効行数: {len(df):,}行  商品数: {df["code"].nunique():,}件')
    return df


def load_products(csv_path):
    """商品マスターCSV読み込み（コード・商品名・仕入先・在庫・廃番・在庫管理対象・仕入単価・売上単価）"""
    logging.info(f'商品CSV読み込み: {csv_path}')
    df = pd.read_csv(
        csv_path,
        encoding='cp932',
        header=0,
        usecols=[0, 2, 11, 12, 13, 18, 19, 23, 28],
        names=['code', 'name', 'stock_mgmt', 'supplier_cd', 'supplier', 'sale_price', 'cost', 'discontinued', 'stock'],
        dtype=str,
        on_bad_lines='skip',
    )
    df['code'] = df['code'].map(normalize_code)
    df = df[df['code'].notna()]
    for col in ['name', 'stock_mgmt', 'supplier_cd', 'supplier', 'discontinued']:
        df[col] = df[col].fillna('').str.strip()
    df['stock'] = pd.to_numeric(
        df['stock'].fillna('0').str.replace(',', '', regex=False), errors='coerce').fillna(0)
    df['cost'] = pd.to_numeric(
        df['cost'].fillna('0').str.replace(',', '', regex=False), errors='coerce').fillna(0)
    df['sale_price'] = pd.to_numeric(
        df['sale_price'].fillna('0').str.replace(',', '', regex=False), errors='coerce').fillna(0)
    return df.set_index('code')


def fetch_reorder_config(gas_url, api_key):
    """GASから分析設定を取得（発注先リードタイム・除外商品・ロット推定材料）"""
    logging.info('GASから設定を取得中 (getReorderConfig)...')
    resp = requests.post(gas_url, json={'action': 'getReorderConfig', 'api_key': api_key},
                         timeout=120)
    resp.raise_for_status()
    result = resp.json()
    if not result.get('success'):
        raise RuntimeError(f'getReorderConfig エラー応答: {result.get("error")}')
    logging.info(f"設定取得完了: 発注先{len(result.get('suppliers', []))}件 / "
                 f"除外{len(result.get('exclusions', []))}件 / "
                 f"ロット材料{len(result.get('lotStats', {}))}商品")
    return result


def classify(adi, cv2):
    if adi < ADI_THRESHOLD:
        return '安定型' if cv2 < CV2_THRESHOLD else '変動型'
    return '散発型' if cv2 < CV2_THRESHOLD else 'まとめ買い型'


def compute_stats(g_month, order_sizes, months, window_days):
    """1商品の需要パターン統計を計算して dict を返す（推奨在庫の計算は含まない。
    ABCランクが決まってからでないと使うZ値・P95フロアの扱いが決められないため分離している）"""
    monthly = [float(g_month.get(ym, 0.0)) for ym in months]
    total = sum(monthly)
    n = len(months)

    demand_months = [q for q in monthly if q > 0]
    n_demand = len(demand_months)

    mean_monthly = total / n
    var = sum((q - mean_monthly) ** 2 for q in monthly) / (n - 1) if n > 1 else 0.0
    std_monthly = math.sqrt(max(var, 0.0))

    # ADI: 需要のあった月の平均間隔（月数）
    adi = n / n_demand if n_demand else float('inf')

    # CV²: 需要のあった月の量のばらつき
    if n_demand >= 2:
        mean_nz = sum(demand_months) / n_demand
        var_nz = sum((q - mean_nz) ** 2 for q in demand_months) / (n_demand - 1)
        cv2 = var_nz / (mean_nz ** 2) if mean_nz > 0 else 0.0
    else:
        cv2 = 0.0

    pattern = classify(adi, cv2)

    # 取引（売上№）単位の統計 — まとめ買い検出用
    sizes = sorted(order_sizes)
    if sizes:
        p95 = sizes[min(len(sizes) - 1, math.ceil(0.95 * len(sizes)) - 1)]
        max_order = sizes[-1]
        mean_order = sum(sizes) / len(sizes)
    else:
        p95 = max_order = mean_order = 0.0

    return {
        'monthly': monthly,
        'mean_monthly': round(mean_monthly, 1),
        'std_monthly': round(std_monthly, 1),
        'demand_month_count': n_demand,
        'adi': round(adi, 2) if adi != float('inf') else None,
        'cv2': round(cv2, 2),
        'pattern': pattern,
        'order_count': len(sizes),
        'mean_order_size': round(mean_order, 1),
        'p95_order_size': round(float(p95), 1),
        'max_order_size': round(float(max_order), 1),
        'daily_mean': total / window_days,
    }


def compute_recommended(stat, protect_days, service_z, p95_mode='full', p95_cap_months=None,
                         global_cap_months=None):
    """ABCランク別のサービス水準・P95フロアの扱いを反映して推奨在庫を計算する
    p95_mode: 'full'=P95フロアをそのまま適用(Aランク) / 'capped'=月数上限つき(Bランク) / 'none'=適用しない(Cランク)
    """
    base_level = (stat['daily_mean'] * protect_days
                  + service_z * stat['std_monthly'] * math.sqrt(protect_days / DAYS_PER_MONTH))
    if p95_mode != 'none' and stat['pattern'] in ('まとめ買い型', '散発型'):
        p95 = stat['p95_order_size']
        if p95_mode == 'capped' and p95_cap_months is not None:
            p95 = min(p95, stat['mean_monthly'] * p95_cap_months)
        recommended = max(base_level, p95)
    else:
        recommended = base_level
    recommended = math.ceil(max(recommended, 0.0))

    if global_cap_months is not None:
        cap = math.ceil(stat['mean_monthly'] * global_cap_months)
        if cap > 0:
            recommended = min(recommended, cap)
    return recommended


def classify_abc(items, value_key='monthly_value'):
    """月次原価貢献の降順・累積構成比でABCランクを付与する（items内の dict に 'abc' キーを追加）"""
    total_value = sum(x[value_key] for x in items) or 1.0
    ranked = sorted(items, key=lambda x: x[value_key], reverse=True)
    cum = 0.0
    for x in ranked:
        cum += x[value_key]
        share = cum / total_value
        if share <= ABC_A_CUM:
            x['abc'] = 'A'
        elif share <= ABC_B_CUM:
            x['abc'] = 'B'
        else:
            x['abc'] = 'C'


def estimate_lot(lot_stats_entry):
    """過去の発注数量からロット（発注単位）を推定。確信が持てない場合は1"""
    if not lot_stats_entry:
        return 1
    if lot_stats_entry.get('orderCount', 0) < LOT_MIN_ORDERS:
        return 1
    gcd_qty = int(lot_stats_entry.get('gcdQty', 0))
    return gcd_qty if gcd_qty >= 2 else 1


def build_note(stat, protect_days, lot, on_order=0, abc='A'):
    """提案根拠の短い説明文（ルールベース）"""
    parts = []
    pattern = stat['pattern']
    if on_order > 0:
        parts.append(f"発注済み未入荷{on_order:.0f}個を在庫に加算済み")
    if pattern == 'まとめ買い型':
        if stat['adi']:
            parts.append(f"約{stat['adi']:.1f}ヶ月間隔で、1回あたり最大{stat['max_order_size']:.0f}個のまとまり注文あり")
        parts.append(f"まとまり注文対応で{stat['p95_order_size']:.0f}個を確保基準に設定")
    elif pattern == '散発型':
        parts.append(f"注文は間欠的（24ヶ月中{stat['demand_month_count']}ヶ月）だが量は安定")
    elif pattern == '変動型':
        parts.append(f"毎月出るが量のブレ大（月{stat['mean_monthly']:.0f}個±{stat['std_monthly']:.0f}）")
    else:
        parts.append(f"月平均{stat['mean_monthly']:.0f}個の安定需要")
    parts.append(f"{protect_days:.0f}日分＋安全在庫で算出")
    if lot > 1:
        parts.append(f"推定ロット{lot}個単位に切り上げ")
    allowance = SERVICE_ALLOWANCE_PCT.get(abc)
    if allowance is not None:
        parts.append(f"{abc}ランク: 欠品許容{allowance}%基準で算出")
    return '。'.join(parts)


def post_proposals(gas_url, api_key, proposals, excess, analyzed_at):
    payload = {
        'action': 'updateOrderProposals',
        'api_key': api_key,
        'analyzedAt': analyzed_at,
        'proposals': proposals,
        'excess': excess,
    }
    logging.info(f'GASへ発注提案・過剰在庫を送信中... (提案{len(proposals)}件 / 過剰在庫{len(excess)}件)')
    for attempt in range(1, 4):
        try:
            resp = requests.post(gas_url, json=payload, timeout=300)
            resp.raise_for_status()
            result = resp.json()
            if result.get('success'):
                logging.info(f'✅ 書き込み成功: 提案{result.get("count")}件 / '
                             f'過剰在庫{result.get("excessCount")}件 (試行{attempt}回目)')
                return True
            logging.warning(f'GASエラー応答 (試行{attempt}): {result.get("error")}')
        except Exception as e:
            logging.warning(f'通信エラー (試行{attempt}): {e}')
    return False


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--csv', help='売上データ明細表.CSV のパス')
    parser.add_argument('--products', help='商品.CSV のパス')
    parser.add_argument('--dry-run', action='store_true', help='GASへ送信せずローカル出力のみ')
    args = parser.parse_args()

    log_file = setup_logger()
    logging.info('=' * 60)
    logging.info('Beaufield 需要分析・発注提案スクリプト v1.4.0 開始'
                 + ('（dry-run）' if args.dry_run else ''))

    config = load_config()
    sales_path = find_csv(args.csv or config.get('csv_path'), SALES_CSV_CANDIDATES, '売上データ明細表.CSV')
    products_path = find_csv(args.products, PRODUCT_CSV_CANDIDATES, '商品.CSV')

    # ---- GASから設定取得（dry-runで失敗したら既定値で続行） ----
    suppliers_cfg = {}
    exclusions = set()
    lot_stats = {}
    recent_orders = {}
    try:
        cfg = fetch_reorder_config(config['gas_url'], config['api_key'])
        for s in cfg.get('suppliers', []):
            key = normalize_supplier_code(s.get('code'))
            if key:
                suppliers_cfg[key] = s
        exclusions = {normalize_code(c) for c in cfg.get('exclusions', [])} - {None}
        lot_stats = {normalize_code(k): v for k, v in cfg.get('lotStats', {}).items()
                     if normalize_code(k)}
        recent_orders = {normalize_code(k): v for k, v in cfg.get('recentOrders', {}).items()
                         if normalize_code(k)}
    except Exception as e:
        if args.dry_run:
            logging.warning(f'GAS設定取得失敗（dry-runのため既定値で続行）: {e}')
        else:
            logging.error(f'GAS設定取得失敗: {e}')
            sys.exit(1)

    # ---- 期間・データ読み込み ----
    start_date, end_date = calc_period()
    start_str, end_str = start_date.strftime('%Y%m%d'), end_date.strftime('%Y%m%d')
    window_days = (end_date - start_date).days + 1
    months = month_range(start_date, end_date)
    logging.info(f'分析期間: {start_date} 〜 {end_date}（{len(months)}ヶ月 / {window_days}日）')

    sales = load_sales(sales_path, start_str, end_str)
    products = load_products(products_path)

    monthly_sum = sales.groupby(['code', 'ym'])['qty'].sum()
    slip_sum = sales[sales['slip'] != ''].groupby(['code', 'slip'])['qty'].sum()
    slip_sum = slip_sum[slip_sum > 0]
    recent6 = sales[sales['ym'] >= months[-6]].groupby('code')['qty'].sum() / 6
    first_ym = sales.groupby('code')['ym'].min()

    # 商品コード→データの辞書化（ループ内でのMultiIndex参照は遅いため事前展開）
    monthly_by_code = {}
    for (code, ym), qty in monthly_sum.items():
        monthly_by_code.setdefault(code, {})[ym] = qty
    sizes_by_code = {}
    for (code, _slip), qty in slip_sum.items():
        sizes_by_code.setdefault(code, []).append(qty)

    min_mean = float(config.get('min_mean_monthly', MIN_MEAN_MONTHLY))

    # ---- 1周目: 需要統計・仕入単価などを集める（ABCランクはまだ決められない） ----
    analyzed = []
    for code in sorted(sales['code'].unique()):
        if code not in products.index:
            continue
        prod = products.loc[code]
        if isinstance(prod, pd.DataFrame):  # 重複コードは先頭を採用
            prod = prod.iloc[0]
        if prod['stock_mgmt'] != 'する' or prod['discontinued'] == '廃番':
            continue

        # 発注先別のリードタイム・発注サイクル
        supp_key = normalize_supplier_code(prod['supplier_cd'])
        supp = suppliers_cfg.get(supp_key, {}) if supp_key else {}
        lead_time = supp.get('leadTimeDays') or DEFAULT_LEAD_TIME_DAYS
        cycle = supp.get('orderCycleDays') or DEFAULT_ORDER_CYCLE_DAYS
        protect_days = float(lead_time) + float(cycle)

        g_month = monthly_by_code.get(code, {})
        sizes = sizes_by_code.get(code, [])
        stat = compute_stats(g_month, sizes, months, window_days)

        if stat['mean_monthly'] <= 0:
            continue

        months_of_history = len([ym for ym in months if ym >= first_ym.get(code, months[0])])
        insufficient = months_of_history < MIN_MONTHS_DATA

        stock = float(prod['stock'])
        unit_cost = float(prod['cost'])
        # ABCランク分類だけに使う経済価値の代理指標。仕入単価が未設定(0円)の商品は
        # 売上単価で代用する（提案金額の表示にはこのフォールバックを使わず unit_cost のまま）
        value_basis = unit_cost or float(prod['sale_price'])
        lot = estimate_lot(lot_stats.get(code))

        # 発注済み・未入荷分: リードタイム内に発注したものはまだ在庫に反映されて
        # いない可能性が高いため、在庫に加算して二重発注を防ぐ
        on_order = 0.0
        lt_cutoff = (date.today() - timedelta(days=int(lead_time))).strftime('%Y%m%d')
        for o in recent_orders.get(code, []):
            if str(o.get('date', '')) >= lt_cutoff:
                on_order += float(o.get('qty', 0))

        excluded = code in exclusions

        analyzed.append({
            'code': code, 'prod': prod, 'supp_key': supp_key, 'protect_days': protect_days,
            'stat': stat, 'insufficient': insufficient, 'stock': stock, 'unit_cost': unit_cost,
            'lot': lot, 'on_order': on_order, 'excluded': excluded,
            'monthly_value': stat['mean_monthly'] * value_basis,
        })

    # ---- ABCランク付与（月次原価貢献の累積構成比。分析対象商品全体が母集団） ----
    classify_abc(analyzed)

    # ---- 2周目: ランク別のサービス水準・P95フロアの扱いで推奨在庫・提案を確定 ----
    results = []
    stats_json = {}
    proposals = []
    excess_rows = []      # 過剰在庫（提案対象かどうかに関わらず全商品が対象）
    comparison_rows = []  # --dry-run時の新旧比較レポート用
    for item in analyzed:
        code, prod, stat, abc = item['code'], item['prod'], item['stat'], item['abc']
        supp_key, protect_days = item['supp_key'], item['protect_days']
        insufficient, stock, unit_cost = item['insufficient'], item['stock'], item['unit_cost']
        lot, on_order, excluded = item['lot'], item['on_order'], item['excluded']

        if abc == 'A':
            recommended = compute_recommended(stat, protect_days, SERVICE_Z_BY_CLASS['A'],
                                               p95_mode='full', global_cap_months=CAP_MONTHS_GLOBAL)
        elif abc == 'B':
            recommended = compute_recommended(stat, protect_days, SERVICE_Z_BY_CLASS['B'],
                                               p95_mode='capped', p95_cap_months=CAP_MONTHS_B,
                                               global_cap_months=CAP_MONTHS_GLOBAL)
        else:
            recommended = compute_recommended(stat, protect_days, SERVICE_Z_BY_CLASS['C'],
                                               p95_mode='none', global_cap_months=CAP_MONTHS_GLOBAL)

        # Cランクの間欠需要（散発型・まとめ買い型）は提案を出さず受注発注推奨とする
        mto_recommended = (abc == 'C') and (stat['pattern'] in ('まとめ買い型', '散発型'))

        shortage = recommended - (stock + on_order)
        excluded_or_mto = excluded or mto_recommended
        eligible = (not insufficient) and (not excluded_or_mto) and stat['mean_monthly'] >= min_mean
        proposed_qty = 0
        if eligible and shortage > 0:
            proposed_qty = int(math.ceil(shortage / lot) * lot)

        pattern_label = 'データ不足' if insufficient else stat['pattern']
        row = {
            'code': code,
            'name': prod['name'],
            'supplier_cd': supp_key or prod['supplier_cd'],
            'supplier': prod['supplier'],
            'pattern': pattern_label,
            'abc': abc,
            'mto_recommended': mto_recommended,
            'excluded': excluded,
            'stock': stock,
            'on_order': on_order,
            'recommended': recommended,
            'proposed_qty': proposed_qty,
            'unit_cost': unit_cost,
            'amount': round(unit_cost * proposed_qty),
            'lot': lot,
            'protect_days': protect_days,
            'current_rp_6mo': round(float(recent6.get(code, 0.0)), 1),
            'mean_monthly': stat['mean_monthly'],
            'std_monthly': stat['std_monthly'],
            'demand_month_count': stat['demand_month_count'],
            'adi': stat['adi'],
            'cv2': stat['cv2'],
            'order_count': stat['order_count'],
            'mean_order_size': stat['mean_order_size'],
            'p95_order_size': stat['p95_order_size'],
            'max_order_size': stat['max_order_size'],
        }
        results.append(row)
        stats_json[code] = {**row, 'monthly': dict(zip(months, stat['monthly']))}

        # 過剰在庫: 現在庫が推奨在庫のEXCESS_RATIO倍を超え、超過数量がEXCESS_MIN_QTY以上
        excess_qty = stock - recommended
        if stock > recommended * EXCESS_RATIO and excess_qty >= EXCESS_MIN_QTY:
            excess_rows.append({
                'code': code,
                'name': prod['name'],
                'supplierCode': supp_key or prod['supplier_cd'],
                'supplierName': prod['supplier'],
                'stock': stock,
                'recommended': recommended,
                'excessQty': excess_qty,
                'unitCost': unit_cost,
                'excessAmount': round(unit_cost * excess_qty),
                'monthsOfStock': round(stock / stat['mean_monthly'], 1),
                'abcRank': abc,
                'pattern': pattern_label,
            })

        if proposed_qty > 0:
            proposals.append({
                'code': code,
                'name': prod['name'],
                'supplierCode': supp_key or prod['supplier_cd'],
                'supplierName': prod['supplier'],
                'pattern': pattern_label,
                'abcRank': abc,
                'stock': stock,
                'onOrder': on_order,
                'recommended': recommended,
                'proposedQty': proposed_qty,
                'unitCost': unit_cost,
                'amount': round(unit_cost * proposed_qty),
                'lot': lot,
                'meanMonthly': stat['mean_monthly'],
                'p95Order': stat['p95_order_size'],
                'maxOrder': stat['max_order_size'],
                'note': build_note(stat, protect_days, lot, on_order, abc),
            })

        if args.dry_run:
            # 新旧比較レポート用: 旧ロジック（ランク差別化なし・Cランク除外なし）での結果も算出
            recommended_old = compute_recommended(stat, protect_days, SERVICE_Z_LEGACY, p95_mode='full')
            shortage_old = recommended_old - (stock + on_order)
            eligible_old = (not insufficient) and (not excluded) and stat['mean_monthly'] >= min_mean
            proposed_qty_old = 0
            if eligible_old and shortage_old > 0:
                proposed_qty_old = int(math.ceil(shortage_old / lot) * lot)
            comparison_rows.append({
                'name': prod['name'], 'abc': abc, 'pattern': pattern_label,
                'recommended_old': recommended_old, 'recommended_new': recommended,
                'proposed_old': proposed_qty_old, 'proposed_new': proposed_qty,
                'amount_old': round(unit_cost * proposed_qty_old), 'amount_new': round(unit_cost * proposed_qty),
            })

    # 仕入先→提案数量の多い順で並べる（アプリでの見やすさ優先）
    proposals.sort(key=lambda x: (x['supplierName'], -x['proposedQty']))
    # 過剰在庫は金額の多い順（資金インパクトが大きい商品から見せる）
    excess_rows.sort(key=lambda x: -x['excessAmount'])

    # ---- ローカル出力 ----
    OUTPUT_DIR.mkdir(exist_ok=True)
    stamp = date.today().strftime('%Y%m%d')
    analyzed_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    df_out = pd.DataFrame(results)
    csv_out = OUTPUT_DIR / f'demand_analysis_{stamp}.csv'
    df_out.to_csv(csv_out, index=False, encoding='utf-8-sig')

    excess_csv_out = OUTPUT_DIR / f'excess_stock_{stamp}.csv'
    pd.DataFrame(excess_rows).to_csv(excess_csv_out, index=False, encoding='utf-8-sig')

    json_out = OUTPUT_DIR / 'demand_stats.json'
    with open(json_out, 'w', encoding='utf-8') as f:
        json.dump({
            'generated_at': analyzed_at,
            'period': {'start': str(start_date), 'end': str(end_date), 'months': months},
            'params': {'service_z_by_class': SERVICE_Z_BY_CLASS, 'min_mean_monthly': min_mean,
                       'default_lead_time_days': DEFAULT_LEAD_TIME_DAYS,
                       'default_order_cycle_days': DEFAULT_ORDER_CYCLE_DAYS,
                       'cap_months_b': CAP_MONTHS_B, 'cap_months_global': CAP_MONTHS_GLOBAL,
                       'abc_a_cum': ABC_A_CUM, 'abc_b_cum': ABC_B_CUM},
            'items': stats_json,
        }, f, ensure_ascii=False, indent=1)

    # ---- サマリー ----
    logging.info('=' * 60)
    logging.info(f'分析対象商品数: {len(df_out):,}件（在庫管理対象・廃番除く・期間内販売あり）')
    for pat, cnt in df_out['pattern'].value_counts().items():
        logging.info(f'  {pat}: {cnt:,}件')
    for rank, cnt in df_out['abc'].value_counts().reindex(['A', 'B', 'C']).fillna(0).items():
        logging.info(f'  {rank}ランク: {cnt:,.0f}件')
    mto_count = int(df_out['mto_recommended'].sum())
    if mto_count:
        logging.info(f'  うちCランク×間欠需要につき受注発注推奨（提案対象外）: {mto_count:,}件')
    total_amount = sum(p['amount'] for p in proposals)
    logging.info(f'発注提案: {len(proposals):,}件 / 合計 {total_amount:,.0f}円'
                 f'（月平均{min_mean}個以上・除外設定{len(exclusions)}件を反映）')
    excess_total = sum(r['excessAmount'] for r in excess_rows)
    logging.info(f'過剰在庫: {len(excess_rows):,}件 / 過剰額合計 {excess_total:,.0f}円'
                 f'（推奨の{EXCESS_RATIO}倍超・超過{EXCESS_MIN_QTY}個以上が対象）')
    logging.info(f'出力: {csv_out}')
    logging.info(f'出力: {excess_csv_out}')
    logging.info(f'出力: {json_out}')

    # ---- 新旧比較レポート（--dry-run時のみ・本適用前のレビュー用） ----
    if args.dry_run and comparison_rows:
        old_count = sum(1 for r in comparison_rows if r['proposed_old'] > 0)
        old_amount = sum(r['amount_old'] for r in comparison_rows)
        diff_pct = f'（{(total_amount - old_amount) / old_amount * 100:+.1f}%）' if old_amount else ''
        logging.info('=' * 60)
        logging.info('【新旧比較レポート】ABCランク差別化ロジック（Phase B）適用時の変化')
        logging.info(f'  提案件数: 旧{old_count:,}件 → 新{len(proposals):,}件')
        logging.info(f'  提案金額: 旧{old_amount:,.0f}円 → 新{total_amount:,.0f}円{diff_pct}')
        logging.info('  新ロジックでのランク別内訳:')
        by_rank = {}
        for p in proposals:
            d = by_rank.setdefault(p['abcRank'], {'count': 0, 'amount': 0})
            d['count'] += 1
            d['amount'] += p['amount']
        for rank in ('A', 'B', 'C'):
            d = by_rank.get(rank, {'count': 0, 'amount': 0})
            logging.info(f'    {rank}ランク: {d["count"]:,}件 / {d["amount"]:,.0f}円')
        drops = sorted(comparison_rows, key=lambda r: r['recommended_old'] - r['recommended_new'], reverse=True)
        logging.info('  推奨在庫の下げ幅トップ20（旧→新）:')
        shown = 0
        for r in drops:
            diff = r['recommended_old'] - r['recommended_new']
            if diff <= 0:
                break
            logging.info(f"    {r['name'][:30]}: {r['abc']}/{r['pattern']} "
                         f"推奨{r['recommended_old']:.0f}→{r['recommended_new']:.0f}個（-{diff:.0f}）")
            shown += 1
            if shown >= 20:
                break
        logging.info('=' * 60)
        logging.info('↑ この内容をレビューし、たかしさんの承認を得てから本適用（--dry-runなしで再実行）してください')

    # ---- GASへ送信 ----
    if args.dry_run:
        logging.info('dry-run のため GAS への送信をスキップしました')
    else:
        if not post_proposals(config['gas_url'], config['api_key'], proposals, excess_rows, analyzed_at):
            logging.error('GASへの発注提案送信が3回すべて失敗しました。ログ: %s', log_file)
            sys.exit(1)

    logging.info('処理完了')


if __name__ == '__main__':
    main()
