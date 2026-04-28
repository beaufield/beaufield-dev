import os
import shutil
import sys

# パス設定
script_dir = os.path.dirname(os.path.abspath(__file__))
onedrive_src = r'D:\OneDrive - Beaufield\PowerBI\Data\商品.CSV'
local_csv    = os.path.join(script_dir, '商品.CSV')
utf8_csv     = os.path.join(script_dir, '商品_utf8.csv')

# ── Step 1: OneDrive からコピー ──────────────────────────
print('[1/2] Copying...')
if not os.path.exists(onedrive_src):
    print('ERROR: Source file not found.')
    print(f'  {onedrive_src}')
    print('OneDrive が同期完了しているか確認してください。')
    sys.exit(1)

shutil.copy2(onedrive_src, local_csv)
print('  OK')

# ── Step 2: CP932 → UTF-8-sig 変換 ──────────────────────
print('[2/2] Converting to UTF-8...')
with open(local_csv, 'r', encoding='cp932') as f:
    content = f.read()

with open(utf8_csv, 'w', encoding='utf-8-sig') as f:
    f.write(content)

print('  OK')
print()
print('Done! 商品_utf8.csv を更新しました。')
