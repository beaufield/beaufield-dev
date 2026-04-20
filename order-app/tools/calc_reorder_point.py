# ============================================================
# Beaufield 発注点算出スクリプト
# Version: v1.0.0
#
# 概要:
#   売上データ明細表.CSV（顧客情報含む）をローカルで処理し、
#   商品コード別の月平均出荷数（=発注点）だけを Google Sheets に書き込む。
#   顧客情報は一切外部に送信しない。
#
# 実行方法:
#   python calc_reorder_point.py
#
# 必要ライブラリ:
#   pip install pandas requests
#
# 設定ファイル:
#   同フォルダの config.json を参照（初回はconfig.json.exampleをコピーして編集）
#
# Windowsタスクスケジューラ設定（推奨）:
#   ・実行タイミング: 毎月1日 07:00
#   ・「タスクを実行できる最も早い時刻に実行する」を有効化
#   ・実行コマンド例:
#       python "D:\Dropbox\ClaudeWork\開発・自動化\order-app\tools\calc_reorder_point.py"
# ============================================================

import json
import logging
import os
import smtplib
import sys
from datetime import date, datetime
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd
import requests

# ============================================================
# 設定読み込み
# ============================================================
SCRIPT_DIR  = Path(__file__).parent
CONFIG_PATH = SCRIPT_DIR / 'config.json'
LOG_DIR     = SCRIPT_DIR / 'logs'

def load_config():
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(
            f'設定ファイルが見つかりません: {CONFIG_PATH}\n'
            'config.json.example をコピーして config.json を作成してください。'
        )
    with open(CONFIG_PATH, encoding='utf-8') as f:
        return json.load(f)

# ============================================================
# ログ設定
# ============================================================
def setup_logger():
    LOG_DIR.mkdir(exist_ok=True)
    log_file = LOG_DIR / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return log_file

# ============================================================
# メール通知
# ============================================================
def send_error_mail(config, subject, body):
    try:
        email_from     = config.get('email_from', '')
        email_to       = config.get('email_to', '')
        app_password   = config.get('email_app_password', '')
        if not (email_from and email_to and app_password):
            logging.warning('メール設定が不完全なため通知をスキップします')
            return
        msg            = MIMEText(body, 'plain', 'utf-8')
        msg['Subject'] = subject
        msg['From']    = email_from
        msg['To']      = email_to
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(email_from, app_password)
            smtp.send_message(msg)
        logging.info(f'エラーメール送信完了: {email_to}')
    except Exception as e:
        logging.warning(f'メール送信失敗（無視）: {e}')

# ============================================================
# 集計期間の計算（当月除く完全6ヶ月）
# 例: 実行日2026-04-20 → 2025-10-01〜2026-03-31
# ============================================================
def calc_period():
    today      = date.today()
    # 先月末
    end_month  = date(today.year, today.month, 1)  # 今月1日
    # 先月末 = 今月1日の前日
    end_date   = date(end_month.year, end_month.month, 1) - __import__('datetime').timedelta(days=1)
    # 6ヶ月前の月初
    m          = today.month - 7  # 先月の1ヶ月前からさらに5ヶ月前 = 今月 - 7
    y          = today.year
    if m <= 0:
        m += 12
        y -= 1
    start_date = date(y, m, 1)
    # YYYYMMDD文字列（CSVとの比較用）
    start_str  = start_date.strftime('%Y%m%d')
    end_str    = end_date.strftime('%Y%m%d')
    return start_str, end_str, start_date, end_date

# ============================================================
# メイン処理
# ============================================================
def main():
    log_file = setup_logger()
    logging.info('=' * 60)
    logging.info('Beaufield 発注点算出スクリプト v1.0.0 開始')

    config = load_config()
    csv_path = Path(config['csv_path'])

    # ---- CSVの存在確認 ----
    if not csv_path.exists():
        msg = f'売上CSVが見つかりません: {csv_path}'
        logging.error(msg)
        send_error_mail(config, '[発注点更新] エラー: CSVファイルが見つかりません', msg)
        sys.exit(1)

    # ---- 集計期間 ----
    start_str, end_str, start_date, end_date = calc_period()
    logging.info(f'集計期間: {start_date} 〜 {end_date}（完全6ヶ月）')

    # ---- CSV読み込み（3列のみ・高速） ----
    logging.info(f'CSV読み込み開始: {csv_path}')
    t0 = datetime.now()
    try:
        # usecols で必要な列のみ読み込む（メモリ・速度を節約）
        # 列インデックス: 0=売上日, 24=商品コード(Y列), 26=数量
        df = pd.read_csv(
            csv_path,
            encoding='cp932',
            header=0,
            usecols=[0, 24, 26],
            names=['date', 'code', 'qty'],
            dtype=str,
            on_bad_lines='skip'
        )
    except Exception as e:
        msg = f'CSV読み込みエラー: {e}'
        logging.error(msg)
        send_error_mail(config, '[発注点更新] エラー: CSV読み込み失敗', msg)
        sys.exit(1)

    elapsed = (datetime.now() - t0).total_seconds()
    logging.info(f'CSV読み込み完了: {len(df):,}行 ({elapsed:.1f}秒)')

    # ---- 前処理 ----
    # 日付フィルタ（YYYYMMDD文字列比較）
    df['date'] = df['date'].fillna('').str.strip()
    df = df[(df['date'] >= start_str) & (df['date'] <= end_str)]
    logging.info(f'期間フィルタ後: {len(df):,}行')

    # 商品コード変換: '  6177  ' → '006177'
    df['code'] = df['code'].fillna('').str.strip()
    def normalize_code(c):
        try:
            return str(int(c)).zfill(6)
        except (ValueError, TypeError):
            return None
    df['code'] = df['code'].apply(normalize_code)
    df = df[df['code'].notna()]

    # 数量変換（カンマ除去・float変換、失敗はスキップ）
    df['qty'] = df['qty'].fillna('0').str.replace(',', '', regex=False)
    df['qty'] = pd.to_numeric(df['qty'], errors='coerce')
    df = df[df['qty'].notna()]

    logging.info(f'変換後の有効行数: {len(df):,}行  ユニーク商品コード数: {df["code"].nunique():,}件')

    # ---- 月別集計 ----
    # 年月キー（YYYYMM）を追加
    df['ym'] = df['date'].str[:6]

    # 商品コード × 年月 でグループ集計
    monthly = df.groupby(['code', 'ym'])['qty'].sum().reset_index()

    # 商品コード別に6ヶ月合計を計算して ÷6（0件月も分母に含む）
    total_by_code = monthly.groupby('code')['qty'].sum()
    avg_by_code   = (total_by_code / 6).round(1)

    # 発注点=0の商品は除外（売上合計がマイナスまたはゼロのケース）
    avg_by_code = avg_by_code[avg_by_code > 0]

    logging.info(f'発注点算出件数: {len(avg_by_code):,}件')

    # ---- GAS WebApp へ POST ----
    updated_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    products   = [
        {'code': code, 'reorderPoint': float(avg), 'updatedAt': updated_at}
        for code, avg in avg_by_code.items()
    ]

    gas_url = config['gas_url']
    api_key = config['api_key']
    payload = {
        'action':   'updateReorderPoints',
        'api_key':  api_key,
        'products': products
    }

    logging.info(f'GAS WebApp へ送信中... ({len(products)}件)')
    success = False
    for attempt in range(1, 4):
        try:
            resp = requests.post(gas_url, json=payload, timeout=120)
            resp.raise_for_status()
            result = resp.json()
            if result.get('success'):
                logging.info(f'✅ GAS書き込み成功: {result.get("count")}件 (試行{attempt}回目)')
                success = True
                break
            else:
                logging.warning(f'GASエラー応答 (試行{attempt}): {result.get("error")}')
        except Exception as e:
            logging.warning(f'通信エラー (試行{attempt}): {e}')

    if not success:
        msg = 'GAS WebApp への送信が3回すべて失敗しました。'
        logging.error(msg)
        send_error_mail(
            config,
            '[発注点更新] エラー: GAS送信失敗',
            f'{msg}\n\nログファイル: {log_file}'
        )
        sys.exit(1)

    # ---- 完了通知 ----
    elapsed_total = (datetime.now() - t0).total_seconds()
    summary = (
        f'集計期間: {start_date} 〜 {end_date}\n'
        f'処理行数: {len(df):,}行\n'
        f'発注点更新件数: {len(products):,}件\n'
        f'所要時間: {elapsed_total:.0f}秒\n'
        f'ログ: {log_file}'
    )
    logging.info('処理完了\n' + summary)

if __name__ == '__main__':
    main()
