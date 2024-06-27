"""--------------------------------------------------
createUpdateQuery.py

    csvからUPDATEクエリ付きのEXCELファイルを作成する。
    @in:
        *.csv
        settings.update.json
    @out: 任意名.xlsx
--------------------------------------------------"""
# 外部ライブラリ
import shutil
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from pathlib import Path
# カスタム関数
import functions as cf
# クラス
from modules import UploadManager


'''========================================
設定ファイル読み込み
========================================'''
jsn = 'settings.update.json'

try:
    inputNum = int(input(f'{jsn}のキー番号を選んでください: '))
    key = str(inputNum)
except ValueError:
    # Enter an integer: abc
    exit(f'input error: {jsn}のキー番号を確認してください。')

'''========================================
前準備
========================================'''
params = cf.getParamsByJson(key, jsn)

root = Path('.')
input_dir = root / params['input_dir']
input_files = sorted(list(input_dir.glob('**/*.csv')))
output_dir = root / params['output_dir']
output_file = f'{output_dir}/{params['output_file']}'

table = params['table']
key_val_dict = params['update_key_val']

# 特定のテーブル用の処理
if inputNum == 1:
    um = UploadManager(input_files)
    key_val_dict['upload_file_url'] = um.getUrlByFiles()
    key_val_dict['upload_filename'] = \
        um.getFileNameByUrls(key_val_dict['upload_file_url'])

# 入力元: 存在チェック
try:
    if not input_dir.exists():
        raise FileNotFoundError
except FileNotFoundError:
    exit(f'{input_dir} が存在しません。')

try:
    if not input_files:
        raise FileNotFoundError
except FileNotFoundError:
    exit(f'{input_dir} にcsvファイルが存在しません。')

# 出力先: 初期化
if output_dir.exists():
    shutil.rmtree(output_dir)
output_dir.mkdir()


'''========================================
csv >> excel書き出し
========================================'''

ws_list = []

for i in input_files:
    df_head = pd.read_csv(i, header=None, nrows=1)

for idx, file in enumerate(input_files):

    offset_num = cf.applyOffsetNum(table)
    df = pd.read_csv(file, header=offset_num)
    df_trimmed = df.replace(
        r'(^[\'|\"|\s]{1}[\s|\t]*|[\s|\t]*[\'|\"]{1}$)', '', regex=True)

    df_result = cf.duplicateDf(df_trimmed, params['sortkey'])

    # ファイル名をシート名にする
    ws_title = re.sub('.csv', '', cf.getFileName(file))
    ws_list.append(ws_title)

    if idx == 0:
        df_result.to_excel(output_file, sheet_name=ws_title, index=False)
    else:
        with pd.ExcelWriter(
            output_file,
            engine='openpyxl',
            mode='a',
            if_sheet_exists='replace'
        ) as writer:
            df_result.to_excel(writer, sheet_name=ws_title, index=False)

'''========================================
excel処理
========================================'''

wb = load_workbook(output_file)

key_idxs = cf.getColumnIndex(
    list(df_result.columns), key_val_dict.keys())
keys = cf.getColumnNames(key_val_dict)
vals_per_ws = cf.getUpdtSrcList(ws_list, key_val_dict)

for ws_idx, ws in enumerate(ws_list):
    ws = wb[ws]

    key_addrs = []
    val_addrs = []
    for i in range(1, len(keys) + 1):
        key_addrs.append(ws.cell(row=i, column=1).coordinate)
        val_addrs.append(ws.cell(row=i, column=2).coordinate)

    fill_color1 = PatternFill(fgColor='7AF5D8', fill_type='solid')
    fill_color2 = PatternFill(fgColor='F5E7EE', fill_type='solid')

    emphasis_font_color = Font(color='FF0000')

    # 更新用の値分だけ上から行追加していく。key_addrsに値を設定する。
    for idx, val in enumerate(vals_per_ws[ws_idx]):
        ws.insert_rows(idx + 1)
        ws[key_addrs[idx]].value = keys[idx]
        ws[key_addrs[idx]].fill = fill_color1
        ws[val_addrs[idx]].value = val

    # src行分下から表ループ開始
    start_row = len(key_addrs) + 2
    interval = 2

    # 一定間隔でloop: interval間隔で表を走査
    for i in range(start_row, ws.max_row+1, interval):

        # 新規行の対象セルを上書きする
        for row in ws.iter_rows(min_row=i, max_row=i):
            for cell in row:
                cell.fill = fill_color2

                # 現在のセルが更新キー列か判定 >> true: srcセルへの参照を埋め込む
                # key_idxs: 更新対象となるキー列の番号を格納
                # key_addrs: 参照元とするセル番地を格納
                # キー列番号とセル番地(list)の長さ・序列は対応している
                for idx, key in enumerate(key_idxs):
                    if cell.column == key:
                        cell.value = f'={val_addrs[idx]}'
                        cell.font = emphasis_font_color

    '''------------------------------
    SQL生成スタート
    ------------------------------'''
    # 表の最後にSQL用の列を追加
    append_col_no = ws.max_column+1

    # SET句で使うカラム(セル番地)を準備
    set_key_addrs = []
    keys_info = list(key_val_dict.keys())
    for row in ws.iter_rows(min_row=start_row-1, max_row=start_row-1):
        for cell in row:
            for key in keys_info:
                if cell.value == key:
                    set_key_addrs.append(cell.coordinate)

    '''********************
    UPDATE句
    ********************'''
    row_count = 1
    for row in ws.iter_rows(
        min_row=start_row,
        max_row=ws.max_row,
        min_col=append_col_no,
        max_col=append_col_no
    ):
        # sql_head 切り戻し用の行はコメントアウトしておく()
        if row_count % 2 == 0:
            sql_u_head = f'="-- UPDATE `{table}` SET"'
        else:
            sql_u_head = f'="UPDATE `{table}` SET"'
        row_count += 1

        sql_u_body = ''
        sql_u_condition = '&"WHERE"&'
        for cell in row:

            # query_body
            for idx, key in enumerate(key_idxs):
                set_key_addr = set_key_addrs[idx]
                set_val_addr = ws.cell(row=cell.row, column=key).coordinate
                value = ws[set_val_addr].value

                # ループ 1st: ループ回数判定(最終回か否か)
                # ループ 2nd: NULL判定
                if idx == len(key_idxs) - 1:
                    if re.fullmatch('NULL', value, flags=re.IGNORECASE):
                        sql_u_body += \
                            f'&" `"&{set_key_addr}&"` = "&{set_val_addr}&" "'
                    else:
                        sql_u_body += \
                            f'&" `"&{set_key_addr}&"` = \'"&{
                                set_val_addr}&"\' "'
                else:
                    if re.fullmatch('NULL', value, flags=re.IGNORECASE):
                        sql_u_body += \
                            f'&" `"&{set_key_addr}&"` = "&{set_val_addr}&","'
                    else:
                        sql_u_body += \
                            f'&" `"&{set_key_addr}&"` = \'"&{
                                set_val_addr}&"\',"'

            # query_condition
            cond_key_addr = ws.cell(row=start_row-1, column=1).coordinate
            cond_val_addr = ws.cell(row=cell.row, column=1).coordinate

            sql_u_condition += \
                f'" `"&{cond_key_addr}&"` = \'"&{cond_val_addr}&"\';"'

            # cell.valueに埋め込み
            cell.value = (sql_u_head + sql_u_body + sql_u_condition)

    '''********************
    SELECT句
    ********************'''
    s_condition_key = ws.cell(row=start_row-1, column=1).coordinate
    s_condition_vals = []

    # 1列目を走査し、条件カラム値を取得
    for i in range(start_row, ws.max_row):

        for row in ws.iter_rows(min_row=i, max_row=i, min_col=1, max_col=1):
            for cell in row:
                s_condition_vals.append(cell.value)

    s_condition_vals = sorted(list(set(s_condition_vals)))

    sql_s_head = f'= "SELECT * FROM `{table}` WHERE `"&{
        s_condition_key}&"` IN('

    sql_s_body = ''
    for idx, val in enumerate(s_condition_vals):
        if idx == len(s_condition_vals)-1:
            sql_s_body += f' \'{val}\' );"'
        else:
            sql_s_body += f' \'{val}\', '

    ws.cell(row=ws.max_row+2, column=1).value = '確認用クエリ'
    ws.cell(row=ws.max_row+1, column=1).value = sql_s_head + sql_s_body

    # ヘッダー行にフィルター設定
    ws.auto_filter.ref = f'{ws.cell(row=start_row - 1, column=1).coordinate}:{
        ws.cell(row=start_row - 1, column=ws.max_column - 1).coordinate}'

wb.save(output_file)
