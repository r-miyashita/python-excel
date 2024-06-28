import re
import json
import urllib.parse as ul
import pandas as pd

'''------------------------------------------------------------
getParamsByJson
    ファイルからパラメータを取得する
        @param
            key: str: 取り出したいテーブルのインデックス
            jsn: str: 設定ファイル
        @return
            params: dict: 設定ファイル内の当該テーブル情報
------------------------------------------------------------'''


def getParamsByJson(key, jsn):
    try:
        with open(jsn) as f:
            return json.load(f)[key]
    except FileNotFoundError as fe:
        print(f'file error: 設定ファイルが存在しません：{fe}')
        exit()
    except KeyError as ke:
        print(f'key error: キー番号が{jsn} に存在しません：{ke}')
        exit()


'''------------------------------------------------------------
getColumnIndex
    データフレームのヘッダーカラムから更新カラムのインデックスを特定し取得する
        @param
            header: list: ヘッダー情報
            updt_clmns: list: 更新カラム名
        @return
            idx_list: list: ヘッダー内更新カラムのインデックス
------------------------------------------------------------'''


def getColumnIndex(header, updt_clmns):
    idx_list = []
    for clmn in updt_clmns:
        idx_list.append(header.index(clmn) + 1)
    return idx_list


'''------------------------------------------------------------
getColumnNames
    更新情報から更新カラム名を取得する
        @param
            key_val_dict: dict: 「カラム名:値」 を格納した辞書
        @return
            name_list: list: カラム名リスト
------------------------------------------------------------'''


def getColumnNames(key_val_dict):
    name_list = []
    for i in key_val_dict.keys():
        name_list.append(i)
    return name_list


'''------------------------------------------------------------
applyOffsetNum
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
getFileName
    ファイルパスからファイル名を取り出す (%エンコードはデコードする)
    ex)
        in: xxxx/xxxx/zzzz/file.csv
        out: file.csv

        @param
            filePath: str
            encoding: str
                'shift-jis'
                ''utf-8': デフォルト値
        @return
            file_name_decoded: str
------------------------------------------------------------'''


def getFileName(filePath, encoding='utf-8'):
    replace_ptn = re.compile(r'(^.+/)')
    file_name = re.sub(replace_ptn, '', str(filePath))
    file_name_decoded = ul.unquote(file_name, encoding=encoding)
    return file_name_decoded


'''------------------------------------------------------------
duplicateDf
    データフレームを複製する
        @param
            df: DataFrame: 基底データ
            sortOpt: dict: ソートキー名と昇降順設定 {sortkey(str): isAscending(bool)}
            iter_count: int: 複製数制御( 現状は1セット複製のみ取り扱う )
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
getUpdtSrcList
    更新用の情報をワークシート単位に切り分け、リスト化する。
        @param
            ws_list: list: ワークシート名のリスト
            key_val_dict: dict: 「カラム名:値」 を格納した辞書
        @return
            val_list: list: ワークシートに対応する値セット
------------------------------------------------------------'''


def getUpdtSrcList(ws_list, key_val_dict):
    src_list = []
    for i in range(len(ws_list)):
        items = []
        for val in key_val_dict.values():
            if isinstance(val, list):
                items.append(val[i])
            else:
                items.append(val)

        src_list.append(items)
    return src_list
