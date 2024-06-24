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


def getParams(key, file):
    try:
        with open('./settings.json') as f:
            return json.load(f)[key]
    except FileNotFoundError as fe:
        print(f'file error: ファイルが存在しません：{fe}')
        exit()
    except KeyError as ke:
        print(f'key error: キーが存在しません：{ke}')
        exit()


'''------------------------------------------------------------
srcからkeyのインデックスを取得する
@param
    src: list
    keys: list
@return
    idx_list: list
------------------------------------------------------------'''


def getIndex(src, keys):
    results = []
    for key in keys:
        results.append(src.index(key) + 1)
    return results


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
    iter_count: int
    sortOpt: dict: {sortkey(str): isAscending(bool)}
@return
    df_result: Dataframe
------------------------------------------------------------'''


def duplicateDf(df, iter_count, sortOpt):
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
update情報から、値セットを作成する。
@param
    df: DataFrame
    iter_count: int
    sortOpt: dict: {sortkey(str): isAscending(bool)}
@return
    df_result: Dataframe
------------------------------------------------------------'''
