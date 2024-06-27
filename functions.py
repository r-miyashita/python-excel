import re
import json
import urllib.parse as ul
import pandas as pd

'''------------------------------------------------------------
ファイルからパラメータを取得する
@param
    key: str
    file: str
@return
    params: dict
------------------------------------------------------------'''


def getParamsByJson(key, file):
    try:
        with open(file) as f:
            return json.load(f)[key]
    except FileNotFoundError as fe:
        print(f'file error: 設定ファイルが存在しません：{fe}')
        exit()
    except KeyError as ke:
        print(f'key error: キー番号が{file} に存在しません：{ke}')
        exit()


'''------------------------------------------------------------
データフレームのヘッダーカラムから更新カラムのインデックスを特定し取得する
@param
    header: list
    updt_clmns: list
@return
    idx_list: list
------------------------------------------------------------'''


def getColumnIndex(header, updt_clmns):
    idx_list = []
    for clmn in updt_clmns:
        idx_list.append(header.index(clmn) + 1)
    return idx_list


'''------------------------------------------------------------
更新情報から更新カラム名を取得する
@param
    updt_src: dict
@return
    name_list: list
------------------------------------------------------------'''


def getColumnNames(updt_src):
    name_list = []
    for i in updt_src.keys():
        name_list.append(i)
    return name_list


'''------------------------------------------------------------
ヘッダーオフセット行を取得する( file読み込みの開始行 )
@param
    table: str
@return
    offset_num: num
------------------------------------------------------------'''


def applyOffsetNum(table):
    list = [
        't_product_technical_manager',
        'pm_t_upload_manager'
    ]
    if table in list:
        return 1
    else:
        return 0


'''------------------------------------------------------------
ファイルパスからファイル名を取り出す
(%エンコードはデコードする)
@param
    filePath: str
    encoding: str
        'shift-jis'
        ''utf-8': デフォルト値
@return
    file_name_decoded: str
ex)
    in: xxxx/xxxx/zzzz/file.csv
    out: file.csv
------------------------------------------------------------'''


def getFileName(filePath, encoding='utf-8'):
    replace_ptn = re.compile(r'(^.+/)')
    file_name = re.sub(replace_ptn, '', str(filePath))
    file_name_decoded = ul.unquote(file_name, encoding=encoding)
    return file_name_decoded


'''------------------------------------------------------------
データフレームを複製する
@param
    df: DataFrame
    sortOpt: dict: {sortkey(str): isAscending(bool)}
    iter_count: int ※デフォルト値
@return
    df_result: Dataframe
------------------------------------------------------------'''


def duplicateDf(df, sortOpt, iter_count=1):
    df_concat = []
    if iter_count:
        for i in range(iter_count + 1):
            df_concat.append(df)

        df_result = pd.concat(df_concat).sort_values(
            by=list(sortOpt.keys()),
            ascending=list(sortOpt.values())
        )
        return df_result
    else:
        return df


'''------------------------------------------------------------
更新用の情報をワークシート単位に切り分け、リスト化する。
@param
    ws_list: list
    updt_src: dict
@return
    src_list: list
------------------------------------------------------------'''


def getUpdtSrcList(ws_list, updt_src):
    src_list = []
    for i in range(len(ws_list)):
        items = []
        for j in updt_src.values():
            items.append(j[i]) if isinstance(j, list) else items.append(j)
        src_list.append(items)
    return src_list
