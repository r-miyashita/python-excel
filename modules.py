import pandas as pd
import re
import urllib.parse as ul


class UploadManager:

    def __init__(self, files):
        self.files = files

    def getUrlByFiles(self):
        url_list = []
        for f in self.files:
            df_head = pd.read_csv(f, header=None, nrows=1)
            new_url = df_head.iloc[0, 0]
            url_list.append(new_url)
        return url_list

    def getFileNameByUrls(self, urls, encoding='utf-8'):
        filename_list = []
        for i in urls:
            ptn = re.compile(r'(^.+/)')
            f_name = re.sub(ptn, '', str(i))
            f_name_decoded = ul.unquote(f_name, encoding)
            filename_list.append(f_name_decoded)
        return filename_list
