"""
converter.py - CSV変換ロジック v1.0.0

OCR出力の5列CSVを、販売管理システム取込フォーマット（61列・CP932）に変換する。
"""

import csv
import io
import json
import os
import re

# ─── 固定値 ────────────────────────────────────────────────────────────────
DEALER_ID        = 'A0020800'
TORIHIKI_NAME    = 'ｶﾌﾞｼｷｶﾞｲｼｬ ﾋﾞｭｰﾌｨｰﾙﾄﾞ'
HASSOU_MOTO_CODE = 'A0020800'

# ─── 出力ヘッダー（61列・PDF仕様準拠）──────────────────────────────────────
OUTPUT_HEADER = [
    '仕入データＩＤ', 'メーカＩＤ', 'ディーラＩＤ', '出荷日', '取引区分',
    '発注番号',       '発注日',     '分類コード',   '伝票区分',  '口座',
    '年月日',         '伝票番号',   '請求日',       '取引先名',  'お届け先区分',
    '訂正区分',       '発注元コード', 'ご帳合先コード', 'ご帳合先名',
    'お届け先コード', 'お届け先名', 'お届け先郵便番号', 'お届け先住所',
    'ケース数合計',   '金額合計',   '値引額合計',   '差引合計',
    '摘要Ａ',         '摘要１行目（発注番号）', '摘要１行目（コメント）',
    '数量引率',       '値引額３',   '値引額４',     '値引額５',  'Ｔ２データ有無',
    '摘要２行目',     '摘要３行目',
    '消費税区分',     '消費税率',   '消費税額',
    '仕入明細ＩＤ',   '行ＮＯ',     '商品コード区分', '商品コード', '商品名',
    '入数',           'ケース',     '発注数量',     '有償バラ数', '景品バラ数',
    '単価',           '単価単位区分', '金額',        'セット分解区分', '荷合わせ数',
    '備考下段',       '備考上段',
    '消費税区分',     '消費税率',   '消費税額',
    'トレーサビリティ情報',
]

# ─── 全角カタカナ→半角カタカナ 変換テーブル ────────────────────────────────
# 濁点・半濁点は2文字（例: ガ→ｶﾞ）になるため dict で管理する
_KANA_MAP = {
    'ァ': 'ｧ', 'ア': 'ｱ', 'ィ': 'ｨ', 'イ': 'ｲ', 'ゥ': 'ｩ', 'ウ': 'ｳ',
    'ェ': 'ｪ', 'エ': 'ｴ', 'ォ': 'ｫ', 'オ': 'ｵ',
    'カ': 'ｶ',  'ガ': 'ｶﾞ', 'キ': 'ｷ',  'ギ': 'ｷﾞ',
    'ク': 'ｸ',  'グ': 'ｸﾞ', 'ケ': 'ｹ',  'ゲ': 'ｹﾞ',
    'コ': 'ｺ',  'ゴ': 'ｺﾞ',
    'サ': 'ｻ',  'ザ': 'ｻﾞ', 'シ': 'ｼ',  'ジ': 'ｼﾞ',
    'ス': 'ｽ',  'ズ': 'ｽﾞ', 'セ': 'ｾ',  'ゼ': 'ｾﾞ',
    'ソ': 'ｿ',  'ゾ': 'ｿﾞ',
    'タ': 'ﾀ',  'ダ': 'ﾀﾞ', 'チ': 'ﾁ',  'ヂ': 'ﾁﾞ',
    'ッ': 'ｯ',  'ツ': 'ﾂ',  'ヅ': 'ﾂﾞ', 'テ': 'ﾃ',  'デ': 'ﾃﾞ',
    'ト': 'ﾄ',  'ド': 'ﾄﾞ',
    'ナ': 'ﾅ',  'ニ': 'ﾆ',  'ヌ': 'ﾇ',  'ネ': 'ﾈ',  'ノ': 'ﾉ',
    'ハ': 'ﾊ',  'バ': 'ﾊﾞ', 'パ': 'ﾊﾟ',
    'ヒ': 'ﾋ',  'ビ': 'ﾋﾞ', 'ピ': 'ﾋﾟ',
    'フ': 'ﾌ',  'ブ': 'ﾌﾞ', 'プ': 'ﾌﾟ',
    'ヘ': 'ﾍ',  'ベ': 'ﾍﾞ', 'ペ': 'ﾍﾟ',
    'ホ': 'ﾎ',  'ボ': 'ﾎﾞ', 'ポ': 'ﾎﾟ',
    'マ': 'ﾏ',  'ミ': 'ﾐ',  'ム': 'ﾑ',  'メ': 'ﾒ',  'モ': 'ﾓ',
    'ャ': 'ｬ',  'ヤ': 'ﾔ',  'ュ': 'ｭ',  'ユ': 'ﾕ',  'ョ': 'ｮ',  'ヨ': 'ﾖ',
    'ラ': 'ﾗ',  'リ': 'ﾘ',  'ル': 'ﾙ',  'レ': 'ﾚ',  'ロ': 'ﾛ',
    'ワ': 'ﾜ',  'ヲ': 'ｦ',  'ン': 'ﾝ',
    'ヴ': 'ｳﾞ', 'ヵ': 'ｶ',  'ヶ': 'ｹ',
    'ー': 'ｰ',  # 長音符（カタカナ）
    '。': '｡',  '「': '｢',  '」': '｣',  '、': '､',  '・': '･',
    '　': ' ',  # 全角スペース
}


def to_hankaku(text: str, max_len: int = 25) -> str:
    """
    全角文字を半角に変換し、max_len 文字以内に切り詰める。
    - 全角カタカナ → 半角カタカナ（濁点は2文字になる）
    - 全角英数字・記号 → 半角英数字・記号
    """
    if not text:
        return ''
    result = []
    for ch in text:
        code = ord(ch)
        if ch in _KANA_MAP:
            result.append(_KANA_MAP[ch])
        elif 0xFF01 <= code <= 0xFF5E:          # 全角！～ → 半角
            result.append(chr(code - 0xFEE0))
        elif code == 0x3000:                     # 全角スペース
            result.append(' ')
        else:
            result.append(ch)
    return ''.join(result)[:max_len]


def load_product_master(master_path: str) -> dict:
    """
    商品_utf8.csv を読み込み {コード: JANCD} の辞書を返す。
    ファイルが存在しない場合は空 dict を返す。
    """
    result = {}
    if not os.path.exists(master_path):
        return result
    try:
        with open(master_path, encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                code  = row.get('コード', '').strip()
                jancd = row.get('JANCD', '').strip()
                if code:
                    result[code] = jancd
    except Exception:
        pass
    return result


def parse_filename(filename: str) -> tuple:
    """
    ファイル名から (maker_name, date_yyyymmdd, slip_no) を抽出する。
    形式: {メーカー名}_{YYYYMMDD}_{伝票番号}.csv
    """
    base  = os.path.splitext(os.path.basename(filename))[0]
    parts = base.split('_')

    date_str  = ''
    date_idx  = -1
    for i, p in enumerate(parts):
        if re.match(r'^\d{8}$', p):
            date_str = p
            date_idx = i
            break

    if date_idx < 0:
        return (base, '', '')

    maker_name = '_'.join(parts[:date_idx])
    slip_no    = '_'.join(parts[date_idx + 1:])
    return (maker_name, date_str, slip_no)


def load_counter(counter_path: str) -> dict:
    """連番カウンターを読み込む。ファイルがなければ初期値を返す。"""
    if os.path.exists(counter_path):
        with open(counter_path, encoding='utf-8') as f:
            return json.load(f)
    return {'shiire_data_id': 10000001, 'shiire_meisai_id': 90000001}


def save_counter(counter_path: str, counter: dict) -> None:
    """連番カウンターをファイルに保存する。"""
    with open(counter_path, 'w', encoding='utf-8') as f:
        json.dump(counter, f, ensure_ascii=False, indent=2)


def convert_ocr_csv(
    csv_content : str,
    filename    : str,
    maker_id    : str,
    trade_type  : str,
    order_no    : str,
    product_master: dict,
    counter     : dict,
) -> list:
    """
    OCR CSV（5列・UTF-8）を 61列フォーマットの行リストに変換する。

    Args:
        csv_content:    CSVファイルのテキスト内容（BOM除去済み）
        filename:       元ファイル名（日付・伝票番号の抽出に使用）
        maker_id:       メーカーID（マスターから選択済み）
        trade_type:     取引区分（'11' or '31'）
        order_no:       発注番号（空欄可）
        product_master: {コード: JANCD} の辞書
        counter:        連番管理 dict（インプレースで更新される）

    Returns:
        61要素のリストを行数分格納したリスト
    """
    _, date_str, slip_no = parse_filename(filename)
    if not re.match(r'^\d{8}$', date_str):
        date_str = ''

    # 仕入データID（伝票単位）を取得・更新
    shiire_data_id = counter['shiire_data_id']
    counter['shiire_data_id'] += 1

    # 入力CSV読み込み
    rows_in = []
    reader  = csv.DictReader(io.StringIO(csv_content))
    for row in reader:
        code  = (row.get('コード')  or '').strip()
        name  = (row.get('商品名')  or '').strip()
        qty   = (row.get('数量')    or '0').strip()
        price = (row.get('単価')    or '0').strip()
        total = (row.get('合計')    or '0').strip()
        # 空行はスキップ
        if not code and not name:
            continue
        rows_in.append((code, name, qty, price, total))

    if not rows_in:
        return []

    # 金額合計（合計列の和）
    kin_total = 0
    for _, _, _, _, total in rows_in:
        try:
            kin_total += int(total)
        except (ValueError, TypeError):
            pass

    output_rows = []
    row_no = 1

    for (code, name, qty, price, total) in rows_in:
        # 仕入明細ID（行単位）を取得・更新
        meisai_id = counter['shiire_meisai_id']
        counter['shiire_meisai_id'] += 1

        # JANコード取得（コード=1 は未照合のため空欄）
        jancd = ''
        if code and code != '1' and code in product_master:
            jancd = product_master[code]

        # 商品名を半角カナに変換・25文字以内
        name_han = to_hankaku(name, 25)

        row = [
            str(shiire_data_id),   #  1: 仕入データID
            maker_id,              #  2: メーカーID
            DEALER_ID,             #  3: ディーラーID
            date_str,              #  4: 出荷日（YYYYMMDD）
            trade_type,            #  5: 取引区分
            order_no,              #  6: 発注番号（空欄可）
            date_str,              #  7: 発注日（= 出荷日）
            '',                    #  8: 分類コード
            '',                    #  9: 伝票区分
            '',                    # 10: 口座
            date_str,              # 11: 年月日（= 出荷日）
            slip_no,               # 12: 伝票番号
            '',                    # 13: 請求日
            TORIHIKI_NAME,         # 14: 取引先名（半角カナ固定）
            '1',                   # 15: お届け先区分
            '',                    # 16: 訂正区分
            HASSOU_MOTO_CODE,      # 17: 発注元コード
            '',                    # 18: ご帳合先コード
            '',                    # 19: ご帳合先名
            '',                    # 20: お届け先コード
            '',                    # 21: お届け先名
            '',                    # 22: お届け先郵便番号
            '',                    # 23: お届け先住所
            '0',                   # 24: ケース数合計
            str(kin_total),        # 25: 金額合計
            '0',                   # 26: 値引額合計
            str(kin_total),        # 27: 差引合計
            '',                    # 28: 摘要A
            '',                    # 29: 摘要1行目（発注番号）
            '',                    # 30: 摘要1行目（コメント）
            '0',                   # 31: 数量引率
            '0',                   # 32: 値引額3
            '0',                   # 33: 値引額4
            '0',                   # 34: 値引額5
            '1',                   # 35: T2データ有無
            '',                    # 36: 摘要2行目
            '',                    # 37: 摘要3行目
            '1',                   # 38: 消費税区分（1=外税）
            '0.1',                 # 39: 消費税率
            '0',                   # 40: 消費税額
            str(meisai_id),        # 41: 仕入明細ID
            str(row_no),           # 42: 行NO
            'J',                   # 43: 商品コード区分（J=JAN）
            jancd,                 # 44: 商品コード（JAN13桁）
            name_han,              # 45: 商品名（半角カナ・25文字以内）
            '',                    # 46: 入数（マスター未対応のため空欄）
            '0',                   # 47: ケース
            qty,                   # 48: 発注数量
            qty,                   # 49: 有償バラ数（= 発注数量）
            '0',                   # 50: 景品バラ数
            price,                 # 51: 単価
            '',                    # 52: 単価単位区分（空=バラ）
            total,                 # 53: 金額
            '0',                   # 54: セット分解区分（0=通常）
            '0',                   # 55: 荷合わせ数
            '',                    # 56: 備考下段
            '',                    # 57: 備考上段
            '',                    # 58: 消費税区分（明細）
            '',                    # 59: 消費税率（明細）
            '',                    # 60: 消費税額（明細）
            '',                    # 61: トレーサビリティ情報（未使用）
        ]

        output_rows.append(row)
        row_no += 1

    return output_rows


def rows_to_csv_bytes(rows: list) -> bytes:
    """
    行リスト（ヘッダー含む）を CP932 エンコードの CSV バイト列に変換する。
    CP932 で表現できない文字は '?' に置換する。
    """
    buf = io.StringIO()
    writer = csv.writer(buf, lineterminator='\r\n')
    writer.writerow(OUTPUT_HEADER)
    writer.writerows(rows)
    return buf.getvalue().encode('cp932', errors='replace')
