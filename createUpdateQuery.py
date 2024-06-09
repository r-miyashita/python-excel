"""
.csvからExcelを作成する。
シート作成
csv転記
sql生成（excel式）
samarry生成
nullチェック
"""
import shutil
import re
import pandas as pd
# from openpyxl import Workbook
from pathlib import Path


root = Path('.')
input_dir = root / 'in_updateQuery'
inputs = sorted(list(input_dir.glob('**/*.csv')))
output_dir = root / 'out_updateQuery'
output_file = f'{output_dir}/test.xlsx'

# input 存在チェック
exit_msg = '処理を終了します。'
err_reasons = {
    'dir_exists': f'{input_dir}が存在しません。ディレクトリを作成し、入力元となるcsvファイルを格納してください。',
    'file_exists': 'csvファイルが存在しません。入力元となるcsvファイルを格納してください。'
}

if not input_dir.exists():
    print(f'{err_reasons['dir_exists']} \n {exit_msg}')
    exit()
elif not inputs:
    print(f'{err_reasons['file_exists']} \n {exit_msg}')
    exit()

# 出力先の初期化
if output_dir.exists():
    shutil.rmtree(output_dir)
output_dir.mkdir()


'''------------------------------
csv転記
------------------------------'''


def createSheetTitle(csvFile):
    replace_ptn = re.compile(r'(^.+/|\.csv$)')
    ws_title = re.sub(replace_ptn, '', str(csvFile))
    return ws_title


ws_list = []
for idx, file in enumerate(inputs):
    title = createSheetTitle(file)
    ws_list.append(title)

df_list = []
new_url_list = []
for idx, file in enumerate(inputs):

    # 1行目から新URLを取得
    df_head = pd.read_csv(file, header=None, nrows=1)
    new_url = df_head.iloc[0, 0]
    new_url_list.append(new_url)

    # 2行目以降からヘッダーとデータを取得
    df = pd.read_csv(file, header=1)
    df_trimmed = df.replace(
        r'(^[\'|\"]{1}[\s|\t]*|[\s|\t]*[\'|\"]{1}$)', '', regex=True)
    df_list.append(df_trimmed)

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
