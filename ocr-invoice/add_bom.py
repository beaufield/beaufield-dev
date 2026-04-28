"""
results フォルダ内の全CSVにBOMを付与する（Excel文字化け対策）
使い方：このファイルをダブルクリック or add_bom.bat を実行
"""
import os
import glob

script_dir = os.path.dirname(os.path.abspath(__file__))
results_dir = os.path.join(script_dir, 'results')

csv_files = glob.glob(os.path.join(results_dir, '*.csv'))

if not csv_files:
    print('CSVファイルが見つかりませんでした')
else:
    for path in csv_files:
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()
        # すでにBOM付きの場合はスキップ
        if content.startswith('\ufeff'):
            print(f'スキップ（BOM済み）: {os.path.basename(path)}')
            continue
        with open(path, 'w', encoding='utf-8-sig') as f:
            f.write(content)
        print(f'BOM付与完了: {os.path.basename(path)}')

print('\n処理完了。Enterキーで閉じます。')
input()
