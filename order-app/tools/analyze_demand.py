# ============================================================
# Beaufield 需要パターン分析・発注提案スクリプト
# Version: v1.13.0
#
# 概要:
#   売上データ明細表.CSV（過去24ヶ月）を分析し、商品ごとに
#   需要パターンを分類して推奨在庫・発注提案を算出する。
#   結果は GAS の「発注提案」シートへ書き込み（→アプリの発注提案タブに表示、
#   LINE WORKS 通知はGAS側が担当）、ローカルにも CSV / JSON を出力する。
#
#   v1.9.0〜: 分析対象の全商品（提案対象かどうかに関わらず）の推奨在庫（recommended）を
#   「発注点マスター」にも書き込む（updateReorderPoints流用）。これにより通常の発注画面・
#   在庫一覧の「適正在庫」バッジ／「推奨発注数」バッジ／「要発注」フィルターが、提案タブと
#   同じ需要分析ベースの基準に統一される。旧・calc_reorder_point.py（6ヶ月月平均のみの
#   簡易ロジック）はこの統一により役目を終えたため使用停止（タスクスケジューラに登録している
#   場合は無効化すること。登録したまま残すと月次実行のたびに古い基準で上書きしてしまう）。
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
#     ただし減少トレンド商品（下記）はランクに関わらずP95フロアを適用しない
#   全ランク共通で「月平均needs×CAP_MONTHS_GLOBALヶ月分」を上限キャップ
#   （減少トレンド商品はCAP_MONTHS_GLOBALの代わりにCAP_MONTHS_DECLINEを使う）
#   リードタイム・発注サイクルは発注先マスターのF/G列（未設定は各7日）
#
# 発注提案の対象:
#   月平均1個以上・販売歴6ヶ月以上・提案除外設定に無い・Cランクの間欠需要でない商品のうち、
#   「現在庫＋発注済み未入荷分」が推奨在庫を下回ったもの。
#   ただし現在庫がマイナスの商品は上記の条件（月平均・販売歴・Cランク間欠需要）を無視して
#   強制的に対象にする（v1.11.0。手動の除外・終売設定のみ引き続き優先される）。
#   発注済み未入荷分 = 発注したが仕入計上されていない数量（v1.12.0〜。下記「発注×仕入突合」参照）
#   提案数量はロット単位に切り上げ。
#   ロット = 手動設定（アプリの提案タブから登録。v1.8.0で追加）があれば最優先、
#           なければ過去の発注明細の数量の最大公約数から自動推定（発注3回以上かつ2以上のときのみ採用）
#   提案金額 = 仕入単価（商品.CSV）× 提案数量。仕入単価未設定（0円）の商品は金額0円扱い
#   （アプリ側で「—」表示。合計からは除外しない）
#
# 参考表示（v1.13.0で追加）:
#   上記の自動条件（Cランクの間欠需要・販売歴6ヶ月未満・月平均1個未満）で提案対象外になった商品でも、
#   「現在庫＋発注済み未入荷分」が推奨在庫を下回っていれば refOnly=True を立てて提案シートに載せる。
#   アプリ側ではチェックOFF・グレー表示になり、発注金額の合計にも件数にも入らない（拾いたい時だけ
#   チェックを入れれば通常どおり発注へ送れる）。
#   狙いは「発注画面では推奨発注数バッジが出ているのに提案タブには影も形も無い」というズレを無くし、
#   在庫が切れかけていること自体には必ず気づけるようにすること。提案するかどうかの判断は据え置き。
#   ⚠️ KPI（提案額・提案件数）・LINE WORKS通知・提案件数の急変ガードレールは参考表示を数えない
#   （経営指標の連続性を壊さないため）。手動の「🚫除外」「🔚終売」は従来どおり最優先で常に非表示。
#
# 過剰在庫（v1.4.0で追加）:
#   現在庫が「推奨在庫のEXCESS_RATIO倍」を超え、かつ超過数量がEXCESS_MIN_QTY以上の商品を
#   「過剰在庫」シートへ書き込む（提案とは別枠）。アプリの提案タブで金額降順に表示し、
#   意図的なまとめ仕入等で問題ない場合は「確認済み」登録すると次回分析後もリストから消える
#
# 経営KPI（v1.5.0で追加）:
#   在庫金額・月次売上原価・回転日数・過剰在庫額・提案額・年間保有コスト概算を
#   「在庫KPI履歴」シートに実行日単位で記録（同日の再実行は上書き）。
#   アプリの提案タブ上部に直近2回分（今週・先週）の差分付きで表示する
#
# 死蔵在庫（v1.10.0で追加）:
#   現行の需要分析ループは売上明細に出てくる商品コードを起点に回るため、分析期間(24ヶ月)に
#   1件も売上がない商品は一度もループに入らず、提案にも過剰在庫にも出てこない不可視の在庫になる。
#   これを商品マスター起点の別パスで検出し「死蔵在庫」シートへ書き込む（提案・過剰在庫とは別枠）。
#   tier1=完全死蔵: 在庫管理する・非廃番・在庫あり・24ヶ月販売ゼロ
#   tier2=休眠    : 分析対象商品のうち直近12ヶ月販売ゼロ・在庫あり（tier1と排他）
#   商品マスターに一度も売上記録がない商品は「未販売」として区別する（新規導入直後の可能性）
#
# 直近トレンド判定（v1.6.0で追加、v1.7.0で精緻化）:
#   直近12ヶ月平均が1年前の12ヶ月平均のDECLINE_RATIO_THRESHOLD倍以下の商品は
#   「減少トレンド」とみなし、需要統計（月平均・パターン・P95等）を24ヶ月全体ではなく
#   直近12ヶ月ベースに切り替える（古い高需要期の実績に推奨在庫が引っ張られるのを防ぐ）
#   ただし直近12ヶ月ウィンドウの先頭付近に旧体制最後の大口注文が1件だけ残っている過渡期は、
#   その1件がP95注文サイズ・std（ばらつき）の両方を歪め、推奨在庫を過大にしてしまう。
#   そのため減少トレンド商品は「まとまり注文フロア(P95)」を適用せず、代わりに
#   月平均×CAP_MONTHS_DECLINEヶ月分を上限キャップとして推奨在庫を頭打ちにする。
#
# 終売フラグ（v1.6.0で追加）:
#   在庫はあるが再発注できない商品（キャンペーン終了等）は「終売商品設定」で
#   個別に指定でき、以後は提案対象から外れる（提案除外設定とは別枠で管理）
#
# 発注×仕入の自動突き合わせ（Phase G・v1.12.0で追加。設計原本: 発注仕入突合_設計プラン.md）:
#   発注済み未入荷（on_order）の判定を、時間窓による推測から仕入実績による事実へ変更した。
#   旧: 「リードタイム以内に発注したもの」＝ on_order とみなす
#   新: 「発注したが仕入データ明細表.CSVに計上されていないもの」＝ on_order
#
#   旧方式の問題: リードタイムが短い仕入先（例 千代田化学=1日）では発注の翌日に
#   on_orderから外れる。一方、実物が届いていても納品書の到着待ちで仕入入力が
#   済んでいないためシステム在庫にも計上されない。結果として在庫にもon_orderにも
#   計上されない空白期間が生まれ、その間ずっと満額で再提案されていた
#   （実測: 千代田化学 16件/12.7万円 → 突合後は 5件/7.3万円）。
#
#   突合は商品コード単位のFIFO（古い発注から順に仕入数量を食わせる）。発注明細と
#   仕入明細の1:1対応は保証されない（分納・欠品・アプリ外発注が混在する）ため。
#   ORDER_OPEN_CUTOFF_DAYS(30日)を超えて計上されない発注は欠品・キャンセルとみなして
#   on_orderから外す（永久に提案が止まって欠品するのを防ぐ安全弁）。
#   あわせて仕入先ごとの「計上ラグ」（仕入日→仕入入力日）を実績中央値から自動推定し、
#   入荷待ちリストの遅延判定に使う（実測でデミ2日/千代田7日/ナプラ10日と差が大きく、
#   一律の日数では遅延判定が機能しないため）。
#
# 在庫マイナス商品の強制表示（v1.11.0で追加）:
#   現在庫がマイナス（欠品・バックオーダー中）の商品は、Cランク間欠需要（受注発注推奨）や
#   データ不足・月平均閾値未満によって通常は提案対象外でも、提案タブに強制的に表示する。
#   「発注済み未入荷分」は考慮した上で算出するため、実際に発注済みで解消見込みの分は
#   従来通り on_order で相殺される（この項目はあくまで純粋な在庫マイナスのみを救済する）。
#   ただし手動の「🚫除外」「🔚終売」設定は従来通り最優先で常に非表示のまま
#   （在庫マイナスでも意図的に提案不要と判断された商品を強制的に出さないため）。
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

# ---- 経営KPIパラメータ（v1.5.0） ----
HOLDING_COST_RATE = 0.20   # 年間保有コスト率（資金・場所・廃番リスク、通説15-25%の中央）

# ---- 直近トレンド判定パラメータ（v1.6.0） ----
# 24ヶ月一律平均だと「過去は出ていたが最近止まった」商品の推奨在庫が過大になるため、
# 直近12ヶ月 と 1年前の12ヶ月 を比較して減少トレンドを検出する。
# 12ヶ月単位で比較するのは季節商品（特定の月しか出ない）を誤検出しないため
DECLINE_TREND_MONTHS    = 12    # 比較に使う月数（1年）
DECLINE_RATIO_THRESHOLD = 0.4   # 直近12ヶ月平均が1年前の12ヶ月平均のこの比率以下なら「減少トレンド」
CAP_MONTHS_DECLINING    = 2.0   # 減少トレンド商品の推奨在庫上限（月平均の何ヶ月分まで許すか。P95フロアの代わり）

# ---- 発注×仕入 突き合わせパラメータ（Phase G, v1.12.0） ----
# 発注済み未入荷（on_order）の判定を「リードタイム以内の発注」という時間窓の推測から
# 「仕入計上されていない発注」という事実ベースに切り替えるためのパラメータ。
# 詳細は 発注仕入突合_設計プラン.md（§3.2〜§3.5）参照
RECEIPT_LOOKBACK_DAYS   = 120   # 仕入CSVの読み込み範囲（発注実績90日 + 余裕30日）
ORDER_OPEN_CUTOFF_DAYS  = 30    # これを超えて仕入計上されない発注は欠品・キャンセルとみなし
                                # on_orderから外す（＝再び提案対象に戻す）安全弁。§5-Q4で一律30日に確定
POSTING_LAG_MIN_SAMPLES = 20    # 仕入先別の計上ラグを実績から推定するのに必要な最低件数
POSTING_LAG_MONTHS      = 6     # 計上ラグ推定に使う実績の期間（ヶ月）
POSTING_LAG_MAX_DAYS    = 90    # 外れ値除外（これを超えるラグは推定に使わない）
DEFAULT_POSTING_LAG_DAYS = 6    # 実績が足りない仕入先に使う既定値（全体中央値の実測値）

# ---- CSVパスの候補（デバイスによりドライブ構成が異なる） ----
_ONEDRIVE_DATA_DIRS = [
    Path(r'D:\OneDrive - Beaufield\PowerBI\Data'),
    Path.home() / 'OneDrive - Beaufield' / 'PowerBI' / 'Data',
]
SALES_CSV_CANDIDATES   = [str(d / '売上データ明細表.CSV') for d in _ONEDRIVE_DATA_DIRS]
PRODUCT_CSV_CANDIDATES = [str(d / '商品.CSV') for d in _ONEDRIVE_DATA_DIRS]
RECEIPT_CSV_CANDIDATES = [str(d / '仕入データ明細表.CSV') for d in _ONEDRIVE_DATA_DIRS]


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

    戻り値: (期間フィルタ後のDataFrame, 商品コード別の全期間・最終売上日dict)
    最終売上日は死蔵在庫検出（v1.10.0）用に、分析期間(24ヶ月)に絞る前の全履歴から取る。
    CSVは1回しか読まない（同じ954,122行を2回読むと+5秒かかるため、期間フィルタ前に集計する）
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
    df['code'] = df['code'].map(normalize_code)
    df = df[df['code'].notna()]

    df['qty'] = pd.to_numeric(
        df['qty'].fillna('0').str.replace(',', '', regex=False), errors='coerce')
    df = df[df['qty'].notna()]

    # 死蔵在庫検出用: 全履歴（分析期間に限らない）でのプラス数量の最終売上日を商品コード別に取得
    # （返品のみの行を最終売上日にしないよう qty>0 に絞る）
    last_sale_by_code = df[df['qty'] > 0].groupby('code')['date'].max().to_dict()

    df = df[(df['date'] >= start_str) & (df['date'] <= end_str)]

    df['slip'] = df['slip'].fillna('').str.strip()
    df['ym'] = df['date'].str[:6]
    logging.info(f'期間内の有効行数: {len(df):,}行  商品数: {df["code"].nunique():,}件')
    return df, last_sale_by_code


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


def load_receipts(csv_path, cutoff_str):
    """仕入明細CSVを読み込み、発注との突き合わせに使う形へ整形する（Phase G, v1.12.0）

    列: 0=仕入日, 5=仕入先コード, 20=商品コード, 22=数量, 28=更新日（仕入入力が行われた日）

    ⚠️ 重要: このCSVに行が存在する ＝ 既に仕入入力済み ＝ 商品.CSVの在庫数にも反映済み。
    したがって突合の判定に更新日は使わず「行があるか」だけを見る。更新日は仕入先別の
    計上ラグ推定（estimate_posting_lags）にのみ使う。

    戻り値: (商品コード別の仕入リスト {code: [(仕入日, 数量)]}, ラグ推定用のDataFrame)
    """
    logging.info(f'仕入CSV読み込み: {csv_path}')
    t0 = datetime.now()
    df = pd.read_csv(
        csv_path,
        encoding='cp932',
        header=0,
        usecols=[0, 5, 20, 22, 28],
        names=['date', 'supplier_cd', 'code', 'qty', 'posted'],
        dtype=str,
        on_bad_lines='skip',
    )
    logging.info(f'読み込み完了: {len(df):,}行 ({(datetime.now()-t0).total_seconds():.1f}秒)')

    df['date'] = df['date'].fillna('').str.strip()
    df = df[df['date'] >= cutoff_str]
    df['code'] = df['code'].map(normalize_code)
    df = df[df['code'].notna()]
    df['qty'] = pd.to_numeric(
        df['qty'].fillna('0').str.replace(',', '', regex=False), errors='coerce')
    # 返品・訂正のマイナス行は突合対象にしない（入荷の事実ではないため）
    df = df[df['qty'].notna() & (df['qty'] > 0)]
    df['supplier_cd'] = df['supplier_cd'].map(normalize_supplier_code)
    df['posted'] = df['posted'].fillna('').str.strip()

    receipts = {}
    for code, d, qty in zip(df['code'], df['date'], df['qty']):
        receipts.setdefault(code, []).append((d, float(qty)))
    for lst in receipts.values():
        lst.sort()

    logging.info(f'仕入実績（直近{RECEIPT_LOOKBACK_DAYS}日）: {len(df):,}行 / {len(receipts):,}商品')
    return receipts, df


def estimate_posting_lags(receipt_df):
    """仕入先ごとの「計上ラグ」（仕入日→仕入入力日の日数）を実績の中央値から推定する

    納品書の到着待ちで仕入入力が遅れる日数は仕入先ごとに大きく異なる
    （実測: デミ2日 / 千代田化学7日 / ナプラ10日 / 水谷理美容鋏48日）。
    一律の日数では遅延判定が機能しないため実績から自動推定する（手動設定を増やさない）。

    戻り値: {仕入先コード: ラグ日数(float)}
    """
    def to_ord(s):
        try:
            return date(int(s[:4]), int(s[4:6]), int(s[6:8])).toordinal()
        except (ValueError, TypeError):
            return None

    cutoff = (date.today() - timedelta(days=int(POSTING_LAG_MONTHS * DAYS_PER_MONTH))).strftime('%Y%m%d')
    df = receipt_df[(receipt_df['date'] >= cutoff) & (receipt_df['posted'].str.len() == 8)]

    lags_by_supplier = {}
    for supp, d, posted in zip(df['supplier_cd'], df['date'], df['posted']):
        if not supp:
            continue
        o1, o2 = to_ord(d), to_ord(posted)
        if o1 is None or o2 is None:
            continue
        lag = o2 - o1
        if 0 <= lag <= POSTING_LAG_MAX_DAYS:
            lags_by_supplier.setdefault(supp, []).append(lag)

    lags = {}
    for supp, vals in lags_by_supplier.items():
        if len(vals) >= POSTING_LAG_MIN_SAMPLES:
            vals.sort()
            n = len(vals)
            lags[supp] = float(vals[n // 2] if n % 2 else (vals[n // 2 - 1] + vals[n // 2]) / 2)

    if lags:
        top = sorted(lags.items(), key=lambda x: -len(lags_by_supplier[x[0]]))[:5]
        logging.info('仕入先別の計上ラグ（実績中央値）: '
                     + ' / '.join(f'{s}={v:.0f}日' for s, v in top)
                     + f' … 計{len(lags)}社（他は既定{DEFAULT_POSTING_LAG_DAYS}日）')
    else:
        logging.warning('計上ラグを推定できる仕入先がありません。全社で既定値を使います')
    return lags


def match_orders_to_receipts(recent_orders, receipts, today=None):
    """発注実績と仕入実績を商品コード単位のFIFOで消し込み、未入荷分を求める（Phase G）

    発注明細と仕入明細の1:1対応は保証されない（分納・欠品・アプリ外発注が混在するため）。
    そこで古い発注から順に仕入数量を食わせ、食い切れずに残った分を「未入荷」とする。

    戻り値: {商品コード: {'open_qty': 未入荷数量, 'lines': [発注明細行ごとの判定]}}
      lines の各要素: {'orderNo', 'date', 'qty', 'open_qty', 'received', 'stale'}
        received=True  … 仕入計上を確認できた（在庫に反映済み）
        stale=True     … ORDER_OPEN_CUTOFF_DAYS を超えて未計上（欠品・キャンセル扱い）
    """
    today = today or date.today()
    stale_cutoff = (today - timedelta(days=ORDER_OPEN_CUTOFF_DAYS)).strftime('%Y%m%d')

    result = {}
    for code, orders in recent_orders.items():
        order_list = sorted(orders, key=lambda o: str(o.get('date', '')))
        # 仕入側は消費しながら使うので可変リストにコピーする
        remaining = list(receipts.get(code, []))

        lines = []
        open_qty = 0.0
        for o in order_list:
            odate = str(o.get('date', ''))
            oqty = float(o.get('qty', 0) or 0)
            if oqty <= 0:
                continue
            need = oqty
            for i, (rdate, rqty) in enumerate(remaining):
                if rqty <= 0 or rdate < odate:
                    continue  # 発注より前の仕入では消し込まない
                take = min(need, rqty)
                remaining[i] = (rdate, rqty - take)
                need -= take
                if need <= 0:
                    break
            stale = need > 0 and odate < stale_cutoff
            lines.append({
                'orderNo': str(o.get('orderNo', '') or ''),
                'date': odate,
                'qty': oqty,
                'open_qty': need,
                'received': need <= 0,
                'stale': stale,
            })
            # 打ち切りを過ぎた未計上分はon_orderに含めない（欠品・キャンセルで永久に
            # 提案が止まるのを防ぐ安全弁。§3.5・§5-Q4で一律30日に確定）
            if need > 0 and not stale:
                open_qty += need

        if lines:
            result[code] = {'open_qty': open_qty, 'lines': lines}
    return result


def warn_if_proposal_count_swings(new_count, threshold=0.5):
    """提案件数が前回実行から大きく振れていたら警告する（Phase G のガードレール・§6）

    仕入CSVが欠損・古いまま突合すると「全部入荷済み」と誤判定して提案が激減しうる。
    逆にアプリ外発注が大量に混入すると激増しうる。どちらも静かに壊れるため気づけるようにする。
    """
    prev_files = sorted(OUTPUT_DIR.glob('demand_analysis_*.csv'))
    # 末尾は今回の出力なので、その1つ前を前回分とする
    if len(prev_files) < 2:
        return
    try:
        prev = pd.read_csv(prev_files[-2], encoding='utf-8-sig')
        mask = pd.to_numeric(prev['proposed_qty'], errors='coerce').fillna(0) > 0
        if 'ref_only' in prev.columns:
            # v1.13.0〜: 参考表示の行は提案件数に数えない（列が無い旧CSVはそのまま比較する）
            mask &= ~prev['ref_only'].astype(str).str.strip().str.lower().isin(['true', '1'])
        prev_count = int(mask.sum())
    except Exception as e:
        logging.debug(f'前回実行との比較をスキップ: {e}')
        return
    if prev_count <= 0:
        return
    ratio = (new_count - prev_count) / prev_count
    if abs(ratio) >= threshold:
        logging.warning(
            f'⚠️ 提案件数が前回実行から大きく変動しています: {prev_count:,}件 → {new_count:,}件'
            f'（{ratio:+.0%}・前回={prev_files[-2].name}）。'
            f'仕入CSVの欠損や更新漏れ、アプリ外発注の大量混入が無いか確認してください')


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
                 f"終売{len(result.get('eolCodes', []))}件 / "
                 f"ロット材料{len(result.get('lotStats', {}))}商品 / "
                 f"手動ロット設定{len(result.get('lotOverrides', {}))}商品")
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


def estimate_lot(lot_stats_entry, lot_override=None):
    """商品のロット（発注単位）を決める。手動設定（アプリの提案タブから登録）があれば
    それを最優先し、なければ過去の発注数量からの自動推定（確信が持てない場合は1）を使う"""
    if lot_override and int(lot_override) >= 1:
        return int(lot_override)
    if not lot_stats_entry:
        return 1
    if lot_stats_entry.get('orderCount', 0) < LOT_MIN_ORDERS:
        return 1
    gcd_qty = int(lot_stats_entry.get('gcdQty', 0))
    return gcd_qty if gcd_qty >= 2 else 1


def build_note(stat, protect_days, lot, on_order=0, abc='A', is_declining=False, older_mean=None,
                forced_by_negative_stock=False, ref_reason=''):
    """提案根拠の短い説明文（ルールベース）"""
    parts = []
    pattern = stat['pattern']
    window_months = DECLINE_TREND_MONTHS if is_declining else WINDOW_MONTHS
    if ref_reason:
        parts.append(f"📎参考表示（{ref_reason}のため自動提案の対象外。在庫が推奨を下回ったのでお知らせのみ）")
    if forced_by_negative_stock:
        parts.append("⚠️在庫マイナスのため通常の対象外条件を無視して表示")
    if on_order > 0:
        parts.append(f"発注済み（仕入未計上）{on_order:.0f}個を在庫に加算済み")
    if is_declining and older_mean is not None:
        parts.append(f"直近{DECLINE_TREND_MONTHS}ヶ月の実績を優先（1年前は月平均{older_mean:.0f}個→直近は月平均{stat['mean_monthly']:.0f}個に減少）")
    if pattern == 'まとめ買い型':
        if stat['adi']:
            parts.append(f"約{stat['adi']:.1f}ヶ月間隔で、1回あたり最大{stat['max_order_size']:.0f}個のまとまり注文あり")
        if is_declining:
            parts.append(f"減少トレンドのためまとまり注文基準は適用せず、月平均×{CAP_MONTHS_DECLINING:.1f}ヶ月分を上限に算出")
        else:
            parts.append(f"まとまり注文対応で{stat['p95_order_size']:.0f}個を確保基準に設定")
    elif pattern == '散発型':
        if is_declining:
            parts.append(f"注文は間欠的（{window_months}ヶ月中{stat['demand_month_count']}ヶ月）。"
                         f"減少トレンドのため月平均×{CAP_MONTHS_DECLINING:.1f}ヶ月分を上限に算出")
        else:
            parts.append(f"注文は間欠的（{window_months}ヶ月中{stat['demand_month_count']}ヶ月）だが量は安定")
    elif pattern == '変動型':
        parts.append(f"毎月出るが量のブレ大（月{stat['mean_monthly']:.0f}個±{stat['std_monthly']:.0f}）")
    else:
        parts.append(f"月平均{stat['mean_monthly']:.0f}個の安定需要")
    parts.append(f"{protect_days:.0f}日分＋安全在庫で算出")
    if lot > 1:
        parts.append(f"最低発注数{lot}個単位に切り上げ")
    allowance = SERVICE_ALLOWANCE_PCT.get(abc)
    if allowance is not None:
        parts.append(f"{abc}ランク: 欠品許容{allowance}%基準で算出")
    return '。'.join(parts)


def post_reorder_points(gas_url, api_key, results, analyzed_at):
    # 発注画面・在庫一覧の「適正在庫」バッジ用に、分析対象の全商品分の推奨在庫を送信する
    # （提案タブに出る・出ないに関わらず。0件は除外＝データ不足や需要ゼロの商品は対象外のまま）
    products = [
        {'code': r['code'], 'reorderPoint': r['recommended'], 'updatedAt': analyzed_at}
        for r in results if r['recommended'] > 0
    ]
    payload = {
        'action': 'updateReorderPoints',
        'api_key': api_key,
        'products': products,
    }
    logging.info(f'GASへ適正在庫（発注点マスター）を送信中... ({len(products)}件)')
    for attempt in range(1, 4):
        try:
            resp = requests.post(gas_url, json=payload, timeout=300)
            resp.raise_for_status()
            result = resp.json()
            if result.get('success'):
                logging.info(f'✅ 適正在庫の書き込み成功: {result.get("count")}件 (試行{attempt}回目)')
                return True
            logging.warning(f'GASエラー応答 (試行{attempt}): {result.get("error")}')
        except Exception as e:
            logging.warning(f'通信エラー (試行{attempt}): {e}')
    return False


def post_proposals(gas_url, api_key, proposals, excess, dead, kpi, analyzed_at):
    payload = {
        'action': 'updateOrderProposals',
        'api_key': api_key,
        'analyzedAt': analyzed_at,
        'proposals': proposals,
        'excess': excess,
        'dead': dead,
        'kpi': kpi,
    }
    ref_count = sum(1 for x in proposals if x.get('refOnly'))
    logging.info(f'GASへ発注提案・過剰在庫・死蔵在庫・KPIを送信中... '
                 f'(提案{len(proposals) - ref_count}件＋参考{ref_count}件 / '
                 f'過剰在庫{len(excess)}件 / 死蔵在庫{len(dead)}件)')
    for attempt in range(1, 4):
        try:
            resp = requests.post(gas_url, json=payload, timeout=300)
            resp.raise_for_status()
            result = resp.json()
            if result.get('success'):
                logging.info(f'✅ 書き込み成功: 提案{result.get("count")}件 / '
                             f'過剰在庫{result.get("excessCount")}件 / '
                             f'死蔵在庫{result.get("deadCount")}件 (試行{attempt}回目)')
                return True
            logging.warning(f'GASエラー応答 (試行{attempt}): {result.get("error")}')
        except Exception as e:
            logging.warning(f'通信エラー (試行{attempt}): {e}')
    return False


def post_receipt_matches(gas_url, api_key, order_matches, posting_lags, analyzed_at):
    """発注×仕入の突合結果をGASへ書き戻す（Phase G, v1.12.0）

    「入荷待ちリストから除外すべき発注明細行」だけを送る。
    - 入荷済み: 仕入計上を確認できた（在庫に反映済み）
    - 打ち切り: 発注から ORDER_OPEN_CUTOFF_DAYS を超えても未計上（欠品・キャンセル扱い）
    未入荷（＝まだ待っている）行は送らない。GAS側でそのまま入荷待ちリストに残る
    """
    matches = []
    missing_order_no = 0
    for code, m in order_matches.items():
        for ln in m['lines']:
            if not ln['orderNo']:
                # 発注No が無い＝GASが旧バージョン（v1.24.0以前）。除外キーを作れない
                missing_order_no += 1
                continue
            if ln['received']:
                status = '入荷済み'
            elif ln['stale']:
                status = '打ち切り'
            else:
                continue
            matches.append({
                'orderNo': ln['orderNo'],
                'code': code,
                'status': status,
                'orderDate': ln['date'],
                'qty': ln['qty'],
            })

    if missing_order_no:
        logging.warning(
            f'⚠️ 発注実績{missing_order_no}行に発注Noが含まれていません。'
            f'GASが旧バージョン（v1.24.0以前）の可能性があります。'
            f'Code.gs v1.25.0以降を再デプロイしてください（入荷待ちの自動消し込みが効きません）')

    payload = {
        'action': 'updateReceiptMatches',
        'api_key': api_key,
        'analyzedAt': analyzed_at,
        'matches': matches,
        'postingLags': posting_lags,
    }
    logging.info(f'GASへ入荷突合結果を送信中... (除外対象{len(matches)}行 / 計上ラグ{len(posting_lags)}社)')
    for attempt in range(1, 4):
        try:
            resp = requests.post(gas_url, json=payload, timeout=300)
            resp.raise_for_status()
            result = resp.json()
            if result.get('success'):
                logging.info(f'✅ 入荷突合結果の書き込み成功: {result.get("count")}行 / '
                             f'計上ラグ{result.get("lagCount")}社 (試行{attempt}回目)')
                return True
            logging.warning(f'GASエラー応答 (試行{attempt}): {result.get("error")}')
        except Exception as e:
            logging.warning(f'通信エラー (試行{attempt}): {e}')
    return False


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--csv', help='売上データ明細表.CSV のパス')
    parser.add_argument('--products', help='商品.CSV のパス')
    parser.add_argument('--receipts', help='仕入データ明細表.CSV のパス')
    parser.add_argument('--dry-run', action='store_true', help='GASへ送信せずローカル出力のみ')
    args = parser.parse_args()

    log_file = setup_logger()
    logging.info('=' * 60)
    logging.info('Beaufield 需要分析・発注提案スクリプト v1.13.0 開始'
                 + ('（dry-run）' if args.dry_run else ''))

    config = load_config()
    sales_path = find_csv(args.csv or config.get('csv_path'), SALES_CSV_CANDIDATES, '売上データ明細表.CSV')
    products_path = find_csv(args.products, PRODUCT_CSV_CANDIDATES, '商品.CSV')
    receipts_path = find_csv(args.receipts, RECEIPT_CSV_CANDIDATES, '仕入データ明細表.CSV')

    # 在庫（商品.CSV）と仕入（仕入データ明細表.CSV）は同じ夜間バッチの出力である前提。
    # 片方だけ古いと突合がズレるため、両ファイルの更新日時をログに残す（§6 ガードレール）
    for label, p in (('商品.CSV', products_path), ('仕入データ明細表.CSV', receipts_path)):
        mtime = datetime.fromtimestamp(Path(p).stat().st_mtime)
        logging.info(f'  データ更新日時 {label}: {mtime:%Y-%m-%d %H:%M}')

    # ---- GASから設定取得（dry-runで失敗したら既定値で続行） ----
    suppliers_cfg = {}
    exclusions = set()
    eol_codes = set()
    lot_stats = {}
    lot_overrides = {}
    recent_orders = {}
    try:
        cfg = fetch_reorder_config(config['gas_url'], config['api_key'])
        for s in cfg.get('suppliers', []):
            key = normalize_supplier_code(s.get('code'))
            if key:
                suppliers_cfg[key] = s
        exclusions = {normalize_code(c) for c in cfg.get('exclusions', [])} - {None}
        eol_codes = {normalize_code(c) for c in cfg.get('eolCodes', [])} - {None}
        lot_stats = {normalize_code(k): v for k, v in cfg.get('lotStats', {}).items()
                     if normalize_code(k)}
        lot_overrides = {normalize_code(k): v for k, v in cfg.get('lotOverrides', {}).items()
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

    sales, last_sale_by_code = load_sales(sales_path, start_str, end_str)
    products = load_products(products_path)

    # ---- 発注×仕入の突き合わせ（Phase G, v1.12.0） ----
    receipt_cutoff = (date.today() - timedelta(days=RECEIPT_LOOKBACK_DAYS)).strftime('%Y%m%d')
    receipts, receipt_df = load_receipts(receipts_path, receipt_cutoff)
    posting_lags = estimate_posting_lags(receipt_df)
    order_matches = match_orders_to_receipts(recent_orders, receipts)

    _open_codes = sum(1 for m in order_matches.values() if m['open_qty'] > 0)
    _lines_all = sum(len(m['lines']) for m in order_matches.values())
    _lines_open = sum(1 for m in order_matches.values() for ln in m['lines']
                      if not ln['received'] and not ln['stale'])
    _lines_stale = sum(1 for m in order_matches.values() for ln in m['lines'] if ln['stale'])
    logging.info(f'発注×仕入 突合: 発注明細{_lines_all:,}行 → 入荷済み{_lines_all - _lines_open - _lines_stale:,}行 / '
                 f'未入荷{_lines_open:,}行（{_open_codes:,}商品）/ '
                 f'{ORDER_OPEN_CUTOFF_DAYS}日超で打ち切り{_lines_stale:,}行')

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

    # 直近トレンド判定用: 直近12ヶ月だけの需要統計を別途用意する
    recent_months = months[-DECLINE_TREND_MONTHS:]
    recent_start_ym = recent_months[0]
    recent_start_date = date(int(recent_start_ym[:4]), int(recent_start_ym[4:6]), 1)
    window_days_recent = (end_date - recent_start_date).days + 1

    sales_recent = sales[sales['ym'] >= recent_start_ym]
    monthly_sum_recent = sales_recent.groupby(['code', 'ym'])['qty'].sum()
    slip_sum_recent = sales_recent[sales_recent['slip'] != ''].groupby(['code', 'slip'])['qty'].sum()
    slip_sum_recent = slip_sum_recent[slip_sum_recent > 0]

    monthly_by_code_recent = {}
    for (code, ym), qty in monthly_sum_recent.items():
        monthly_by_code_recent.setdefault(code, {})[ym] = qty
    sizes_by_code_recent = {}
    for (code, _slip), qty in slip_sum_recent.items():
        sizes_by_code_recent.setdefault(code, []).append(qty)

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
        stat_full = compute_stats(g_month, sizes, months, window_days)

        if stat_full['mean_monthly'] <= 0:
            continue

        # 直近トレンド判定: 1年前の12ヶ月平均 vs 直近12ヶ月平均。大きく減っていたら
        # 統計を直近12ヶ月ベースに切り替える（古い高需要期の実績に引っ張られないように）
        older_months = months[:DECLINE_TREND_MONTHS]
        older_mean_monthly = sum(g_month.get(ym, 0.0) for ym in older_months) / len(older_months)
        g_month_recent = monthly_by_code_recent.get(code, {})
        sizes_recent = sizes_by_code_recent.get(code, [])
        stat_recent = compute_stats(g_month_recent, sizes_recent, recent_months, window_days_recent)
        is_declining = (older_mean_monthly >= MIN_MEAN_MONTHLY and
                        stat_recent['mean_monthly'] <= older_mean_monthly * DECLINE_RATIO_THRESHOLD)
        stat = stat_recent if is_declining else stat_full

        months_of_history = len([ym for ym in months if ym >= first_ym.get(code, months[0])])
        insufficient = months_of_history < MIN_MONTHS_DATA

        stock = float(prod['stock'])
        unit_cost = float(prod['cost'])
        # ABCランク分類だけに使う経済価値の代理指標。仕入単価が未設定(0円)の商品は
        # 売上単価で代用する（提案金額の表示にはこのフォールバックを使わず unit_cost のまま）
        value_basis = unit_cost or float(prod['sale_price'])
        lot = estimate_lot(lot_stats.get(code), lot_overrides.get(code))

        # 発注済み・未入荷分（Phase G, v1.12.0）: 仕入データとの突き合わせで
        # 「発注したが仕入計上されていない」数量を求め、在庫に加算して二重発注を防ぐ。
        # ⚠️ v1.11.0以前は「リードタイム以内に発注したもの」という時間窓の推測だったが、
        #   リードタイムが短い仕入先（千代田化学=1日）では、実物が届いていても納品書待ちで
        #   仕入入力されていない期間に在庫にもon_orderにも計上されない空白ができていた。
        #   詳細は 発注仕入突合_設計プラン.md §1.2・§3.2
        on_order = order_matches.get(code, {}).get('open_qty', 0.0)

        excluded = code in exclusions
        eol_flagged = code in eol_codes

        analyzed.append({
            'code': code, 'prod': prod, 'supp_key': supp_key, 'protect_days': protect_days,
            'stat': stat, 'is_declining': is_declining, 'stat_full_mean': stat_full['mean_monthly'],
            'older_mean': round(older_mean_monthly, 1),
            'insufficient': insufficient, 'stock': stock, 'unit_cost': unit_cost,
            'lot': lot, 'on_order': on_order, 'excluded': excluded, 'eol_flagged': eol_flagged,
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
        eol_flagged, is_declining = item['eol_flagged'], item['is_declining']
        older_mean = item['older_mean']

        z = SERVICE_Z_BY_CLASS[abc]
        if is_declining:
            # 減少トレンド商品はランクに関わらずP95フロアを適用しない（旧体制最後の
            # 大口注文1件がP95・stdを歪めるため）。代わりに専用の月数キャップで頭打ちにする
            recommended = compute_recommended(stat, protect_days, z,
                                               p95_mode='none', global_cap_months=CAP_MONTHS_DECLINING)
        elif abc == 'A':
            recommended = compute_recommended(stat, protect_days, z,
                                               p95_mode='full', global_cap_months=CAP_MONTHS_GLOBAL)
        elif abc == 'B':
            recommended = compute_recommended(stat, protect_days, z,
                                               p95_mode='capped', p95_cap_months=CAP_MONTHS_B,
                                               global_cap_months=CAP_MONTHS_GLOBAL)
        else:
            recommended = compute_recommended(stat, protect_days, z,
                                               p95_mode='none', global_cap_months=CAP_MONTHS_GLOBAL)

        # Cランクの間欠需要（散発型・まとめ買い型）は提案を出さず受注発注推奨とする
        mto_recommended = (abc == 'C') and (stat['pattern'] in ('まとめ買い型', '散発型'))

        shortage = recommended - (stock + on_order)
        stock_negative = stock < 0
        # 通常の対象条件（データ十分・Cランク間欠需要でない・月平均が閾値以上）
        normally_eligible = (not insufficient) and (not mto_recommended) and stat['mean_monthly'] >= min_mean
        # 手動の「🚫除外」「🔚終売」は最優先。提案にも参考表示にも一切出さない
        manually_hidden = excluded or eol_flagged
        # 在庫マイナスの商品は上記の自動除外条件を無視して救済表示する（手動の除外・終売のみ優先）
        eligible = (not manually_hidden) and (stock_negative or normally_eligible)
        forced_by_negative_stock = stock_negative and not normally_eligible
        # 参考表示（v1.13.0）: 自動条件で提案対象外だが在庫が推奨を下回っている商品。
        # 提案と同じ行として送るが refOnly を立て、アプリ側でチェックOFF・グレー表示にして
        # 発注金額の合計にも件数にも含めない（切れかけに気づけないのを防ぐ可視化のみが目的）
        ref_only = (not manually_hidden) and (not eligible) and shortage > 0
        proposed_qty = 0
        if (eligible or ref_only) and shortage > 0:
            proposed_qty = int(math.ceil(shortage / lot) * lot)

        # 参考表示になった理由（アプリの根拠メモに出す。複数該当しうるので全部並べる）
        ref_reason = ''
        if ref_only:
            _reasons = []
            if insufficient:
                _reasons.append(f'販売歴{MIN_MONTHS_DATA}ヶ月未満')
            if mto_recommended:
                _reasons.append('Cランクの間欠需要で受注発注推奨')
            if stat['mean_monthly'] < min_mean:
                _reasons.append(f'月平均{min_mean:g}個未満')
            ref_reason = '・'.join(_reasons)

        pattern_label = 'データ不足' if insufficient else stat['pattern']
        row = {
            'code': code,
            'name': prod['name'],
            'supplier_cd': supp_key or prod['supplier_cd'],
            'supplier': prod['supplier'],
            'pattern': pattern_label,
            'abc': abc,
            'mto_recommended': mto_recommended,
            'ref_only': ref_only,
            'excluded': excluded,
            'eol_flagged': eol_flagged,
            'is_declining': is_declining,
            'older_12mo_mean': older_mean,
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
        # stat['monthly'] は減少トレンド商品なら直近12ヶ月分、そうでなければ24ヶ月分
        # なので、対応する月ラベルもそれに合わせないとズレる（v1.7.0で修正）
        stat_months = recent_months if is_declining else months
        stats_json[code] = {**row, 'monthly': dict(zip(stat_months, stat['monthly']))}

        # 過剰在庫: 現在庫が推奨在庫のEXCESS_RATIO倍を超え、超過数量がEXCESS_MIN_QTY以上
        excess_qty = stock - recommended
        if stock > recommended * EXCESS_RATIO and excess_qty >= EXCESS_MIN_QTY:
            # 直近12ヶ月の需要が完全にゼロ（減少トレンド判定でmean_monthly=0）の場合は
            # 24ヶ月平均を使って在庫月数を算出する（0除算防止・少なくとも過去の実績ベースで表示）
            monthly_rate = stat['mean_monthly'] if stat['mean_monthly'] > 0 else item['stat_full_mean']
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
                'monthsOfStock': round(stock / monthly_rate, 1) if monthly_rate > 0 else None,
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
                'refOnly': ref_only,
                'note': build_note(stat, protect_days, lot, on_order, abc, is_declining, older_mean,
                                   forced_by_negative_stock, ref_reason),
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

    # 仕入先→提案が先・参考は後→提案数量の多い順で並べる（シートを直接見た時の分かりやすさ優先。
    # アプリ側は棚番順に並べ替えるのでここの並びは表示順には影響しない）
    proposals.sort(key=lambda x: (x['supplierName'], x['refOnly'], -x['proposedQty']))
    # 参考表示（refOnly）は「気づくため」の行であって発注提案そのものではないので、
    # KPI・通知・件数ガードレールは提案行だけを母集団にする
    real_proposals = [x for x in proposals if not x['refOnly']]
    ref_proposals  = [x for x in proposals if x['refOnly']]
    # 過剰在庫は金額の多い順（資金インパクトが大きい商品から見せる）
    excess_rows.sort(key=lambda x: -x['excessAmount'])

    # ---- 死蔵在庫検出（Phase E, v1.10.0） ----
    # tier1=完全死蔵: 在庫管理する・非廃番・在庫あり・分析期間(24ヶ月)の売上明細に1件も出現しない商品
    #   （上の1周目ループは sales['code'].unique() を起点に回るため、この層は一度もループに
    #    入らず提案にも過剰在庫にも出てこない不可視の在庫。ここで商品マスター起点に取り直す）
    # tier2=休眠    : 分析対象商品(results)のうち、直近12ヶ月の販売数量合計が0 かつ在庫あり
    #   （tier1と母集団が排他なので重複しない）
    sold_codes_period = set(sales['code'].unique())
    dead_rows = []

    def _last_sale_info(code):
        last_sale = last_sale_by_code.get(code)
        if not last_sale:
            return '', None, '未販売'
        last_sale_fmt = f'{last_sale[:4]}-{last_sale[4:6]}-{last_sale[6:8]}'
        months_since = round(
            (date.today() - date(int(last_sale[:4]), int(last_sale[4:6]), int(last_sale[6:8]))).days
            / DAYS_PER_MONTH, 1)
        return last_sale_fmt, months_since, None

    for code in products.index.unique():
        prod = products.loc[code]
        if isinstance(prod, pd.DataFrame):
            prod = prod.iloc[0]
        if prod['stock_mgmt'] != 'する' or prod['discontinued'] == '廃番':
            continue
        stock = float(prod['stock'])
        if stock <= 0 or code in sold_codes_period:
            continue
        unit_cost = float(prod['cost'])
        supp_key = normalize_supplier_code(prod['supplier_cd'])
        last_sale_fmt, months_since, unsold_reason = _last_sale_info(code)
        dead_rows.append({
            'code': code,
            'name': prod['name'],
            'supplierCode': supp_key or prod['supplier_cd'],
            'supplierName': prod['supplier'],
            'stock': stock,
            'unitCost': unit_cost,
            'deadAmount': round(stock * unit_cost),
            'lastSaleDate': last_sale_fmt,
            'monthsSinceLastSale': months_since,
            'tier': '完全死蔵',
            'reason': unsold_reason or f'{WINDOW_MONTHS}ヶ月販売ゼロ',
        })

    for row in results:
        code = row['code']
        if row['stock'] <= 0:
            continue
        recent_sum = sum(monthly_by_code_recent.get(code, {}).values())
        if recent_sum > 0:
            continue
        last_sale_fmt, months_since, unsold_reason = _last_sale_info(code)
        dead_rows.append({
            'code': code,
            'name': row['name'],
            'supplierCode': row['supplier_cd'],
            'supplierName': row['supplier'],
            'stock': row['stock'],
            'unitCost': row['unit_cost'],
            'deadAmount': round(row['stock'] * row['unit_cost']),
            'lastSaleDate': last_sale_fmt,
            'monthsSinceLastSale': months_since,
            'tier': '休眠',
            'reason': unsold_reason or f'直近{DECLINE_TREND_MONTHS}ヶ月販売ゼロ',
        })

    dead_rows.sort(key=lambda x: -x['deadAmount'])

    # ---- ローカル出力 ----
    OUTPUT_DIR.mkdir(exist_ok=True)
    stamp = date.today().strftime('%Y%m%d')
    analyzed_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    df_out = pd.DataFrame(results)
    csv_out = OUTPUT_DIR / f'demand_analysis_{stamp}.csv'
    df_out.to_csv(csv_out, index=False, encoding='utf-8-sig')

    excess_csv_out = OUTPUT_DIR / f'excess_stock_{stamp}.csv'
    pd.DataFrame(excess_rows).to_csv(excess_csv_out, index=False, encoding='utf-8-sig')

    dead_csv_out = OUTPUT_DIR / f'dead_stock_{stamp}.csv'
    pd.DataFrame(dead_rows).to_csv(dead_csv_out, index=False, encoding='utf-8-sig')

    json_out = OUTPUT_DIR / 'demand_stats.json'
    with open(json_out, 'w', encoding='utf-8') as f:
        json.dump({
            'generated_at': analyzed_at,
            'period': {'start': str(start_date), 'end': str(end_date), 'months': months},
            'params': {'service_z_by_class': SERVICE_Z_BY_CLASS, 'min_mean_monthly': min_mean,
                       'default_lead_time_days': DEFAULT_LEAD_TIME_DAYS,
                       'default_order_cycle_days': DEFAULT_ORDER_CYCLE_DAYS,
                       'cap_months_b': CAP_MONTHS_B, 'cap_months_global': CAP_MONTHS_GLOBAL,
                       'cap_months_declining': CAP_MONTHS_DECLINING,
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
    declining_count = int(df_out['is_declining'].sum())
    if declining_count:
        logging.info(f'  うち直近{DECLINE_TREND_MONTHS}ヶ月ベースに切り替え（減少トレンド検出）: {declining_count:,}件')
    eol_count = int(df_out['eol_flagged'].sum())
    total_amount = sum(p['amount'] for p in real_proposals)
    logging.info(f'発注提案: {len(real_proposals):,}件 / 合計 {total_amount:,.0f}円'
                 f'（月平均{min_mean}個以上・除外設定{len(exclusions)}件・終売設定{eol_count}件を反映）')
    if ref_proposals:
        ref_amount = sum(p['amount'] for p in ref_proposals)
        logging.info(f'  参考表示（自動提案の対象外だが在庫が推奨を下回る）: {len(ref_proposals):,}件 / '
                     f'参考額 {ref_amount:,.0f}円 ※KPI・通知には含めない')
    warn_if_proposal_count_swings(len(real_proposals))
    excess_total = sum(r['excessAmount'] for r in excess_rows)
    logging.info(f'過剰在庫: {len(excess_rows):,}件 / 過剰額合計 {excess_total:,.0f}円'
                 f'（推奨の{EXCESS_RATIO}倍超・超過{EXCESS_MIN_QTY}個以上が対象）')
    dead_total = sum(d['deadAmount'] for d in dead_rows)
    tier1_rows = [d for d in dead_rows if d['tier'] == '完全死蔵']
    tier2_rows = [d for d in dead_rows if d['tier'] == '休眠']
    logging.info(f'死蔵在庫: {len(dead_rows):,}件 / 在庫額合計 {dead_total:,.0f}円'
                 f'（完全死蔵 {len(tier1_rows):,}件 {sum(d["deadAmount"] for d in tier1_rows):,.0f}円 / '
                 f'休眠 {len(tier2_rows):,}件 {sum(d["deadAmount"] for d in tier2_rows):,.0f}円）')
    logging.info(f'出力: {csv_out}')
    logging.info(f'出力: {excess_csv_out}')
    logging.info(f'出力: {dead_csv_out}')
    logging.info(f'出力: {json_out}')

    # ---- 経営KPI ----
    stock_value = float((df_out['stock'].clip(lower=0) * df_out['unit_cost']).sum())
    monthly_cogs = float((df_out['mean_monthly'] * df_out['unit_cost']).sum())
    turnover_days = round(stock_value / monthly_cogs * DAYS_PER_MONTH, 1) if monthly_cogs > 0 else 0
    holding_cost_annual = round(stock_value * HOLDING_COST_RATE)
    kpi = {
        'date': date.today().strftime('%Y-%m-%d'),
        'stockValue': round(stock_value),
        'monthlyCogs': round(monthly_cogs),
        'turnoverDays': turnover_days,
        'excessAmount': excess_total,
        'excessCount': len(excess_rows),
        'proposalAmount': total_amount,
        'proposalCount': len(real_proposals),
        'holdingCostAnnual': holding_cost_annual,
        'deadAmount': dead_total,
        'deadCount': len(dead_rows),
    }
    logging.info(f'在庫金額: {stock_value:,.0f}円 / 月次売上原価: {monthly_cogs:,.0f}円 / '
                 f'回転日数: {turnover_days:.1f}日 / 年間保有コスト概算: {holding_cost_annual:,.0f}円')

    # ---- 新旧比較レポート（--dry-run時のみ・本適用前のレビュー用） ----
    if args.dry_run and comparison_rows:
        old_count = sum(1 for r in comparison_rows if r['proposed_old'] > 0)
        old_amount = sum(r['amount_old'] for r in comparison_rows)
        diff_pct = f'（{(total_amount - old_amount) / old_amount * 100:+.1f}%）' if old_amount else ''
        logging.info('=' * 60)
        logging.info('【新旧比較レポート】ABCランク差別化ロジック（Phase B）適用時の変化')
        logging.info(f'  提案件数: 旧{old_count:,}件 → 新{len(real_proposals):,}件')
        logging.info(f'  提案金額: 旧{old_amount:,.0f}円 → 新{total_amount:,.0f}円{diff_pct}')
        logging.info('  新ロジックでのランク別内訳:')
        by_rank = {}
        for p in real_proposals:
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
        if not post_proposals(config['gas_url'], config['api_key'], proposals, excess_rows, dead_rows, kpi, analyzed_at):
            logging.error('GASへの発注提案送信が3回すべて失敗しました。ログ: %s', log_file)
            sys.exit(1)
        if not post_reorder_points(config['gas_url'], config['api_key'], results, analyzed_at):
            logging.error('GASへの適正在庫送信が3回すべて失敗しました。ログ: %s', log_file)
            sys.exit(1)
        # 入荷突合結果の書き戻しは非致命的にする。提案・適正在庫の書き込みは既に成功しており、
        # GASが未デプロイ（updateReceiptMatches が無い）だとここだけ失敗するため。
        # 失敗しても提案の数値は正しい（on_orderはPython側で算出済み）。
        # 影響は「入荷待ちリストの自動消し込みが効かない」だけなので処理は続行する
        if not post_receipt_matches(config['gas_url'], config['api_key'],
                                    order_matches, posting_lags, analyzed_at):
            logging.error('⚠️ 入荷突合結果の送信に失敗しました（提案の更新自体は成功しています）。'
                          'Code.gs v1.25.0以降が再デプロイされているか確認してください。ログ: %s', log_file)

    logging.info('処理完了')


if __name__ == '__main__':
    main()
