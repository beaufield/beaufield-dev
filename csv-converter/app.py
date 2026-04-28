"""
app.py - CSV変換アプリ Flask メインアプリ v1.0.0

起動方法:
    python app.py
    または start.bat をダブルクリック
"""

import io
import json
import os
from datetime import datetime

from flask import Flask, jsonify, render_template, request, send_file

from converter import (
    convert_ocr_csv,
    load_counter,
    load_product_master,
    parse_filename,
    rows_to_csv_bytes,
    save_counter,
)

# ─── パス設定 ────────────────────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
MASTER_PATH = os.path.join(BASE_DIR, 'master.json')
COUNTER_PATH= os.path.join(BASE_DIR, 'counter.json')
OUTPUT_DIR  = os.path.join(BASE_DIR, 'output')
PRODUCT_MASTER_PATH = os.path.join(
    BASE_DIR, '..', 'ocr-invoice', '商品_utf8.csv'
)

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ─── Flask アプリ初期化 ──────────────────────────────────────────────────────
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB

# 商品マスター（起動時に一度だけ読み込む）
product_master = load_product_master(PRODUCT_MASTER_PATH)


# ─── ユーティリティ ──────────────────────────────────────────────────────────

def load_maker_master() -> dict:
    """メーカーIDマスター（master.json）を読み込む。"""
    if os.path.exists(MASTER_PATH):
        with open(MASTER_PATH, encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_maker_master(data: dict) -> None:
    """メーカーIDマスター（master.json）を保存する。"""
    with open(MASTER_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ─── ルート ──────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    """メイン画面を表示する。"""
    return render_template('index.html')


# ── マスター管理 API ──────────────────────────────────────────────────────────

@app.route('/api/master', methods=['GET'])
def api_master_get():
    """メーカーIDマスターを返す。"""
    return jsonify(load_maker_master())


@app.route('/api/master', methods=['POST'])
def api_master_add():
    """メーカーを追加または上書き登録する。"""
    data = request.get_json(silent=True) or {}
    name     = (data.get('name') or '').strip()
    maker_id = (data.get('id')   or '').strip()

    if not name or not maker_id:
        return jsonify({'error': 'メーカー名とIDを入力してください'}), 400

    master = load_maker_master()
    master[name] = maker_id
    save_maker_master(master)
    return jsonify({'ok': True, 'master': master})


@app.route('/api/master/<path:name>', methods=['DELETE'])
def api_master_delete(name):
    """メーカーを削除する。"""
    master = load_maker_master()
    if name in master:
        del master[name]
        save_maker_master(master)
    return jsonify({'ok': True, 'master': master})


# ── 変換 API ─────────────────────────────────────────────────────────────────

@app.route('/api/convert', methods=['POST'])
def api_convert():
    """
    OCR CSV を受け取り、61列フォーマットの CSV を返す。

    Form fields:
        file       : OCR CSV ファイル（multipart）
        maker_id   : メーカーID（必須）
        trade_type : 取引区分（'11' or '31'、デフォルト '11'）
        order_no   : 発注番号（任意）
    """
    # ファイルチェック
    if 'file' not in request.files or not request.files['file'].filename:
        return jsonify({'error': 'ファイルが選択されていません'}), 400

    file = request.files['file']

    # フォームパラメータ
    maker_id   = (request.form.get('maker_id')   or '').strip()
    trade_type = (request.form.get('trade_type') or '11').strip()
    order_no   = (request.form.get('order_no')   or '').strip()

    if not maker_id:
        return jsonify({'error': 'メーカーを選択してください'}), 400

    # ファイル読み込み（BOM付き UTF-8 対応）
    try:
        content = file.read().decode('utf-8-sig')
    except UnicodeDecodeError:
        return jsonify({'error': 'ファイルの文字コードが UTF-8 ではありません'}), 400

    # 変換処理
    counter = load_counter(COUNTER_PATH)
    rows = convert_ocr_csv(
        csv_content    = content,
        filename       = file.filename,
        maker_id       = maker_id,
        trade_type     = trade_type,
        order_no       = order_no,
        product_master = product_master,
        counter        = counter,
    )

    if not rows:
        return jsonify({'error': 'CSVにデータが見つかりませんでした'}), 400

    # カウンターを保存
    save_counter(COUNTER_PATH, counter)

    # 出力ファイル名を決定
    _, date_str, slip_no = parse_filename(file.filename)
    if date_str and slip_no:
        out_filename = f'import_{date_str}_{slip_no}.csv'
    elif date_str:
        out_filename = f'import_{date_str}.csv'
    else:
        out_filename = f'import_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'

    # CP932 で CSV バイト列を生成
    csv_bytes = rows_to_csv_bytes(rows)

    # output フォルダにも保存（バックアップ）
    out_path = os.path.join(OUTPUT_DIR, out_filename)
    with open(out_path, 'wb') as f:
        f.write(csv_bytes)

    # ブラウザにダウンロードさせる
    return send_file(
        io.BytesIO(csv_bytes),
        mimetype='text/csv; charset=cp932',
        as_attachment=True,
        download_name=out_filename,
    )


# ─── 起動 ────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    print('=' * 50)
    print('  CSV Converter v1.0.0  Starting...')
    print('  Open http://127.0.0.1:5000 in your browser')
    if not os.path.exists(PRODUCT_MASTER_PATH):
        print(f'  [!] Product master not found: {PRODUCT_MASTER_PATH}')
        print('      JAN codes will be blank.')
    else:
        print(f'  [OK] Product master: {len(product_master)} items loaded.')
    print('=' * 50)
    app.run(debug=False, host='127.0.0.1', port=5000)
