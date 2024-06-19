"""
.csvからExcelを作成する。
    csv転記(ファイル数に応じてシート追加)
    sql生成（excel式埋め込み）
"""
# 外部ライブラリ
import shutil
import re
import datetime as dt
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from pathlib import Path
# カスタム関数
import functions as cf

'''------------------------------
前準備
------------------------------'''
root = Path('.')
input_dir = root / 'in_updateQuery'
input_files = sorted(list(input_dir.glob('**/*.csv')))
output_dir = root / 'out_updateQuery'
output_file = f'{output_dir}/test.xlsx'

table = 'table'
update_user = 'todays_user'
update_datetime = dt.datetime.now()

exit_msg = '処理を終了します。'
err_reasons = {
    'dir_exists': f'{input_dir}が存在しません。ディレクトリを作成し、入力元となるcsvファイルを格納してください。',
    'file_exists': 'csvファイルが存在しません。入力元となるcsvファイルを格納してください。'
}

# 後続のループで利用するために入れ物だけ用意
wb = None

# 入力元: 存在チェック

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


ws_list = []
new_url_list = []

# 6/18 追加（未置換） dest_column_numsと置き換える
updt_columns_idx = [4, 5, 6, 7]  # paramで受け取る
updt_columns = {}
##

for idx, file in enumerate(input_files):

    # 1行目から新URLを取得
    df_head = pd.read_csv(file, header=None, nrows=1)
    new_url = df_head.iloc[0, 0]
    new_url_list.append(new_url)

    # 2行目以降からヘッダーとデータを取得
    df = pd.read_csv(file, header=1)
    df_trimmed = df.replace(
        r'(^[\'|\"|\s]{1}[\s|\t]*|[\s|\t]*[\'|\"]{1}$)', '', regex=True)

    # 6/18 追加
    # 初回にデータフレームから更新対象列を取得する
    if idx == 0:
        for i in updt_columns_idx:
            updt_columns[i] = list(df_trimmed)[i-1]
    print(list(range(1, len(updt_columns) + 1)))
    # iter_count分データを複製する
    df_concat = []
    df_result = df_trimmed
    sortkey_dict = {'id': True}  # sortkey: isAscending(True or False)
    iter_count = 1
    if iter_count:
        for i in range(iter_count+1):
            df_concat.append(df_trimmed)
        df_result = pd.concat(df_concat).sort_values(
            by=list(sortkey_dict.keys()),
            ascending=list(sortkey_dict.values()))

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

for idx, ws in enumerate(ws_list):
    ws = wb[ws]

    '''
    上書き対象のカラム番号、参照元セルを設定
        4: upload_filename
        5: upload_file_url
        6: update_user
        7: update_datetime
    '''
    updt_src_vals = [
        cf.getFileName(new_url_list[idx]),
        new_url_list[idx],
        update_user,
        update_datetime
    ]

    updt_src_cells = []
    for i in range(1, len(updt_columns) + 1):
        updt_src_cells.append(ws.cell(row=i, column=1).coordinate)

    new_row_fill = PatternFill(fgColor='F5E7EE', fill_type='solid')
    emphasis_font_color = Font(color='FF0000')

    # 更新用の値分だけ上から行追加していく。updt_src_cellsに値を設定する。
    for idx, val in enumerate(updt_src_vals):
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
                # column_keys: 更新対象となるキー列の番号を格納
                # updt_src_cells: 参照元とするセル番地を格納
                # キー列番号とセル番地(list)の長さ・序列は対応している
                column_keys = list(updt_columns.keys())
                for idx, key in enumerate(column_keys):
                    if cell.column == key:
                        cell.value = f'={updt_src_cells[idx]}'
                        cell.font = emphasis_font_color

    # 最終列にSQLを追加する
    # 奇数行： 更新用
    # 偶数行： 切り戻し用
    append_column_no = ws.max_column+1

    # SET句で使うカラム(セル番地)を準備
    colname_cells = []
    colnames = list(updt_columns.values())
    for row in ws.iter_rows(min_row=start_row-1, max_row=start_row-1):
        for cell in row:
            for colname in colnames:
                if cell.value == colname:
                    colname_cells.append(cell.coordinate)

    # queryを作成
    count = 1
    for row in ws.iter_rows(
        min_row=start_row,
        max_row=ws.max_row,
        min_col=append_column_no,
        max_col=append_column_no
    ):
        # sql_head 切り戻し用の行はコメントアウトしておく()
        if count % 2 == 0:
            sql_u_head = f'="-- UPDATE {table} SET"'
        else:
            sql_u_head = f'="UPDATE {table} SET"'
        count += 1

        sql_u_body = ''
        sql_u_condition = '&"WHERE"&'
        for cell in row:

            # query_body
            column_keys = list(updt_columns.keys())
            for idx, key in enumerate(column_keys):
                col_pos = colname_cells[idx]
                val_pos = ws.cell(row=cell.row, column=key).coordinate
                value = ws[val_pos].value

                if idx == len(column_keys) - 1:
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
            condition_column = ws.cell(row=start_row-1, column=1).coordinate
            condition_val = ws.cell(row=cell.row, column=1).coordinate

            sql_u_condition += f'" `"&{
                condition_column}&"` = \'"&{condition_val}&"\';"'

            # cell.valueに埋め込み
            cell.value = (sql_u_head + sql_u_body + sql_u_condition)

    # 確認用 select句作成

    # キー列、値(セル番地)を取得する
    # 1列目をプライマリーキー列として決め打ちしている
    s_condition_key = ws.cell(row=start_row-1, column=1).coordinate
    s_condition_values = []

    # 重複がないように、行飛ばしで走査
    for i in range(start_row, ws.max_row, interval):

        for row in ws.iter_rows(min_row=i, max_row=i, min_col=1, max_col=1):
            for cell in row:
                s_condition_values.append(cell.coordinate)

    sql_s_head = f'= "SELECT * FROM `{table}` WHERE `"&{
        s_condition_key}&"` IN("'

    sql_s_body = ''
    for idx, val in enumerate(s_condition_values):
        if idx == len(s_condition_values)-1:
            sql_s_body += f'&" \'"&{val}&"\' );"'
        else:
            sql_s_body += f'&" \'"&{val}&"\', "'

    ws.cell(row=ws.max_row+2, column=1).value = '確認用クエリ'
    ws.cell(row=ws.max_row+1, column=1).value = sql_s_head + sql_s_body

    # ヘッダー行にフィルター設定
    ws.auto_filter.ref = f'{ws.cell(row=start_row - 1, column=1).coordinate}:{
        ws.cell(row=start_row - 1, column=ws.max_column - 1).coordinate}'

wb.save(output_file)
