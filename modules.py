import pandas as pd
import re
import urllib.parse as ul

'''============================================================
UploadManager

    S3管理用のテーブル更新の際はINデータ(.csv)にオブジェクトURL情報を追記しておく運用
    ファイルを受け取り、URL情報を加工するためのクラス
============================================================'''

'''---------------------------------------------
getUrlByFiles
    インプットファイル数分のオブジェクトURLを取得
        @return
            url_list: list

getFileNameByUrls
    URLの末尾ファイル名のみ取得
        @return
            filename_list: list
---------------------------------------------'''


class UploadManager:

    def __init__(self, files):
        self.files = files

    def getUrlByFiles(self, replace_ptn):
        url_list = []
        r_tgt = replace_ptn['tgt']
        r_src = replace_ptn['src']

        for f in self.files:
            df_head = pd.read_csv(f, header=None, nrows=1)
            url = df_head.iloc[0, 0]
            url_list.append(re.sub(r_tgt, r_src, url))
        return url_list

    def getFileNameByUrls(self, urls, encoding='utf-8'):
        filename_list = []

        for i in urls:
            ptn = re.compile(r'(^.+/)')
            f_name = re.sub(ptn, '', str(i))
            f_name_decoded = ul.unquote(f_name, encoding)
            filename_list.append(f_name_decoded)
        return filename_list
