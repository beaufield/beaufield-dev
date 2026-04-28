"""
商品_utf8.csv の最終更新日時を確認するスクリプト
「納品書スキャンして」実行時にClaude Codeが呼び出す
"""
import os
from datetime import datetime, timedelta

script_dir = os.path.dirname(os.path.abspath(__file__))
csv_path = os.path.join(script_dir, "商品_utf8.csv")

if not os.path.exists(csv_path):
    print("WARNING: 商品_utf8.csvが存在しません。csv変換.batを実行してください。")
else:
    mtime = datetime.fromtimestamp(os.path.getmtime(csv_path))
    now = datetime.now()
    age = now - mtime
    days = age.days
    hours = age.seconds // 3600

    if days >= 1:
        print(f"WARNING|{days}日{hours}時間前")
    else:
        print(f"OK|{hours}時間前")
