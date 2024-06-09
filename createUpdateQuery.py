"""
.csvからExcelを作成する。
    csv転記(ファイル数に応じてシート追加)
    sql生成（excel式埋め込み）
    クライアントシステムでSQL一括実行するためのコピペフィールド生成(excel式埋め込み)
    nullの警告
"""
import shutil
import re
import pandas as pd
# from openpyxl import Workbook
from pathlib import Path

'''------------------------------
前準備
------------------------------'''
root = Path('.')
input_dir = root / 'in_updateQuery'
input_files = sorted(list(input_dir.glob('**/*.csv')))
output_dir = root / 'out_updateQuery'
output_file = f'{output_dir}/test.xlsx'

# 入力元: 存在チェック
exit_msg = '処理を終了します。'
err_reasons = {
    'dir_exists': f'{input_dir}が存在しません。ディレクトリを作成し、入力元となるcsvファイルを格納してください。',
    'file_exists': 'csvファイルが存在しません。入力元となるcsvファイルを格納してください。'
}

if not input_dir.exists():
    print(f'{err_reasons['dir_exists']} \n {exit_msg}')
    exit()
elif not input_files:
    print(f'{err_reasons['file_exists']} \n {exit_msg}')
    exit()

# 出力先: 初期化
if output_dir.exists():
    shutil.rmtree(output_dir)
output_dir.mkdir()


'''------------------------------
csv >> excel書き出し(シート追記)
------------------------------'''


def createSheetTitle(csvFile):
    replace_ptn = re.compile(r'(^.+/|\.csv$)')
    ws_title = re.sub(replace_ptn, '', str(csvFile))
    return ws_title


ws_list = []
for idx, file in enumerate(input_files):
    title = createSheetTitle(file)
    ws_list.append(title)

for idx, file in enumerate(input_files):

    # 1行目から新URLを取得
    df_head = pd.read_csv(file, header=None, nrows=1)
    new_url = df_head.iloc[0, 0]

    # 2行目以降からヘッダーとデータを取得
    df = pd.read_csv(file, header=1)
    df_trimmed = df.replace(
        r'(^[\'|\"]{1}[\s|\t]*|[\s|\t]*[\'|\"]{1}$)', '', regex=True)

    if idx == 0:
        df_trimmed.to_excel(output_file, sheet_name=ws_list[idx])
    else:
        with pd.ExcelWriter(
            output_file,
            engine='openpyxl',
            mode='a', if_sheet_exists='replace'
        ) as writer:
            df_trimmed.to_excel(writer, sheet_name=ws_list[idx])

'''------------------------------
sql生成
------------------------------'''

'''------------------------------
samarry生成
------------------------------'''

'''------------------------------
nullチェック
------------------------------'''

# wb.save(output_file)
