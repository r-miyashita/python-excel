# 参照用セルA列はカラム名を入れる
# select のid条件を改善する(？)

"""
.csvからExcelを作成する。
    csv転記(ファイル数に応じてシート追加)
    sql生成（excel式埋め込み）
"""
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

try:
    inputNum = int(input('settings.jsonのキー番号を選んでください: '))
    key = str(inputNum)
except ValueError:
    # Enter an integer: abc
    print('input error: settings.jsonのキー番号を確認してください。')
    exit()

params = cf.getParamsByJson(key, 'settings.json')

'''------------------------------
前準備
------------------------------'''
root = Path('.')
input_dir = root / params['input_dir']
input_files = sorted(list(input_dir.glob('**/*.csv')))
output_dir = root / params['output_dir']
output_file = f'{output_dir}/{params['output_file']}'

table = params['table']
updt_clmns = params['update_key_val']

if inputNum == 1:
    um = UploadManager(input_files)
    updt_clmns['upload_file_url'] = um.getUrlByFiles()
    updt_clmns['upload_filename'] = \
        um.getFileNameByUrls(updt_clmns['upload_file_url'])

err_reasons = {
    'dir_err': f'{input_dir}が存在しません。ディレクトリを作成し、入力元となるcsvファイルを格納してください。',
    'file_err': 'csvファイルが存在しません。入力元となるcsvファイルを格納してください。'
}

# 入力元: 存在チェック

if not input_dir.exists():
    print(f'{err_reasons['dir_err']}')
    exit()
elif not input_files:
    print(f'{err_reasons['file_err']}')
    exit()


# 出力先: 初期化
if output_dir.exists():
    shutil.rmtree(output_dir)
output_dir.mkdir()


'''------------------------------
csv >> excel書き出し(シート追記)
------------------------------'''

ws_list = []

for i in input_files:
    df_head = pd.read_csv(i, header=None, nrows=1)
    # new_url = df_head.iloc[0, 0]

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


# excel処理
wb = load_workbook(output_file)

updt_clmns_idx = cf.getColumnIndex(list(df_result.columns), updt_clmns.keys())
updt_clmn_names = cf.getColumnNames(updt_clmns)
updt_src = cf.getUpdtSrcList(ws_list, updt_clmns)

for ws_idx, ws in enumerate(ws_list):
    ws = wb[ws]

    updt_src_cells = []
    for i in range(1, len(updt_clmn_names) + 1):
        updt_src_cells.append(ws.cell(row=i, column=1).coordinate)

    new_row_fill = PatternFill(fgColor='F5E7EE', fill_type='solid')
    emphasis_font_color = Font(color='FF0000')

    # 更新用の値分だけ上から行追加していく。updt_src_cellsに値を設定する。
    for idx, val in enumerate(updt_src[ws_idx]):
        ws.insert_rows(idx + 1)
        ws[updt_src_cells[idx]].value = val

    # src行分下から表ループ開始
    start_row = len(updt_src_cells) + 2
    interval = 2

    # 一定間隔でloop: interval間隔で表を走査
    for i in range(start_row, ws.max_row+1, interval):

        # 新規行の対象セルを上書きする
        for row in ws.iter_rows(min_row=i, max_row=i):
            for cell in row:
                cell.fill = new_row_fill

                # 現在のセルが更新キー列か判定 >> true: srcセルへの参照を埋め込む
                # updt_clmns_idx: 更新対象となるキー列の番号を格納
                # updt_src_cells: 参照元とするセル番地を格納
                # キー列番号とセル番地(list)の長さ・序列は対応している
                for idx, key in enumerate(updt_clmns_idx):
                    if cell.column == key:
                        cell.value = f'={updt_src_cells[idx]}'
                        cell.font = emphasis_font_color

    # 最終列にSQLを追加する
    # 奇数行： 更新用
    # 偶数行： 切り戻し用
    append_column_no = ws.max_column+1

    # SET句で使うカラム(セル番地)を準備
    colname_cells = []
    colnames = list(updt_clmns.keys())
    for row in ws.iter_rows(min_row=start_row-1, max_row=start_row-1):
        for cell in row:
            for colname in colnames:
                if cell.value == colname:
                    colname_cells.append(cell.coordinate)

    # queryを作成
    row_count = 1
    for row in ws.iter_rows(
        min_row=start_row,
        max_row=ws.max_row,
        min_col=append_column_no,
        max_col=append_column_no
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
            for idx, key in enumerate(updt_clmns_idx):
                col_pos = colname_cells[idx]
                val_pos = ws.cell(row=cell.row, column=key).coordinate
                value = ws[val_pos].value

                # ループ 1st: ループ回数判定(最終回か否か)
                # ループ 2nd: NULL判定
                if idx == len(updt_clmns_idx) - 1:
                    if re.fullmatch('NULL', value, flags=re.IGNORECASE):
                        sql_u_body += \
                            f'&" `"&{col_pos}&"` = "&{val_pos}&" "'
                    else:
                        sql_u_body += \
                            f'&" `"&{col_pos}&"` = \'"&{val_pos}&"\' "'
                else:
                    if re.fullmatch('NULL', value, flags=re.IGNORECASE):
                        sql_u_body += \
                            f'&" `"&{col_pos}&"` = "&{val_pos}&","'
                    else:
                        sql_u_body += \
                            f'&" `"&{col_pos}&"` = \'"&{val_pos}&"\',"'

            # query_condition
            condition_col = ws.cell(row=start_row-1, column=1).coordinate
            condition_val = ws.cell(row=cell.row, column=1).coordinate

            sql_u_condition += f'" `"&{
                condition_col}&"` = \'"&{condition_val}&"\';"'

            # cell.valueに埋め込み
            cell.value = (sql_u_head + sql_u_body + sql_u_condition)

    # 確認用 select句作成

    # キー列、値(セル番地)を取得する
    # 1列目をプライマリーキー列として決め打ちしている
    s_condition_key = ws.cell(row=start_row-1, column=1).coordinate
    s_condition_vals = []

    # 重複がないように、行飛ばしで走査
    for i in range(start_row, ws.max_row, interval):

        for row in ws.iter_rows(min_row=i, max_row=i, min_col=1, max_col=1):
            for cell in row:
                s_condition_vals.append(cell.coordinate)

    sql_s_head = f'= "SELECT * FROM `{table}` WHERE `"&{
        s_condition_key}&"` IN("'

    sql_s_body = ''
    for idx, val in enumerate(s_condition_vals):
        if idx == len(s_condition_vals)-1:
            sql_s_body += f'&" \'"&{val}&"\' );"'
        else:
            sql_s_body += f'&" \'"&{val}&"\', "'

    ws.cell(row=ws.max_row+2, column=1).value = '確認用クエリ'
    ws.cell(row=ws.max_row+1, column=1).value = sql_s_head + sql_s_body

    # ヘッダー行にフィルター設定
    ws.auto_filter.ref = f'{ws.cell(row=start_row - 1, column=1).coordinate}:{
        ws.cell(row=start_row - 1, column=ws.max_column - 1).coordinate}'

wb.save(output_file)
