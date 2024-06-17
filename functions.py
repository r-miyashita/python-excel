import re
import urllib.parse as ul

'''------------------------------
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
------------------------------'''


def getFileName(filePath, encoding='utf-8'):
    replace_ptn = re.compile(r'(^.+/)')
    file_name = re.sub(replace_ptn, '', str(filePath))
    file_name_decoded = ul.unquote(file_name, encoding=encoding)
    return file_name_decoded
