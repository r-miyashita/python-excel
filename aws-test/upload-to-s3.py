import os
import boto3
from botocore.exceptions import NoCredentialsError, PartialCredentialsError
from dotenv import load_dotenv
from urllib import parse, request
import time
import mimetypes


# .env ファイルから環境変数を読み込む
load_dotenv()

# s3 アクセス情報を設定する
AWS_ACCESS_KEY_ID = os.getenv('AWS_ACCESS_KEY_ID')
AWS_SECRET_ACCESS_KEY = os.getenv('AWS_SECRET_ACCESS_KEY')
AWS_REGION = os.getenv('AWS_REGION')
BUCKET_NAME = os.getenv('BUCKET_NAME')

# envチェック
if not all(
        [AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_REGION, BUCKET_NAME]):
    raise ValueError(
        'AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_REGION, BUCKET_NAME '
        'を.envに設定してください。'
    )

# s3 クライアントを作成
s3_client = boto3.client(
    's3',
    aws_access_key_id=AWS_ACCESS_KEY_ID,
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
    region_name=AWS_REGION
)

# ファイルリスト リソースフォルダ の指定
UPLOAD_FILE_LIST_PATH = './upload_file_list.txt'
FILES_DIR = './files/'

# エンコード除外キーワード
EXCLUDE_CHARS = '/-_.~'


def check_url_accessible(url, retries=3, delay=2):
    # 疎通失敗しても規定回数リトライする
    for attempt in range(retries):
        try:
            with request.urlopen(url) as response:
                if response.status == 200:
                    # 疎通成功
                    return True
        except Exception as e:
            print(f'疎通確認失敗（{attempt + 1}/{retries}）: {e}')
            if attempt < retries - 1:
                time.sleep(delay)
    return False


def write_results_to_file(
        success_list,
        failure_list,
        output_file='./upload_results.txt'
):
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('成功 ####\n')
            for item in success_list:
                f.write(f'{item['file_name']},{item['url']}\n')

            f.write('\n')

            f.write('失敗 ####\n')
            for item in failure_list:
                f.write(f'{item['file_name']}:{item['reason']}\n')

        print(f'結果を{output_file}に出力しました')
    except Exception as e:
        print(f'結果ファイルの出力に失敗しました。: {e}')


def upload_files_to_s3():
    success_list = []
    failure_list = []
    try:
        # ファイルリスト読み込み
        with open(UPLOAD_FILE_LIST_PATH, 'r', encoding='utf-8') as file_list:
            # 1行ずつ ファイルパスを取得
            for line in file_list:
                s3_key = line.strip()
                if not s3_key:
                    continue

                # ファイル名を抽出してローカルパスを生成
                file_name = os.path.basename(s3_key)
                local_file_path = os.path.join(FILES_DIR, file_name)
                if not os.path.isfile(local_file_path):
                    failure_list.append({
                        'file_name': file_name,
                        'reason': 'ローカルファイルが見つかりません'
                    })
                    continue

                # MIMEタイプを取得 #mime_type の例: 'image/jpeg'
                mime_type, _ = mimetypes.guess_type(local_file_path)

                # ExtraArgsを作成
                extra_args = {
                    'ACL': 'public-read',
                    # MIMEタイプ不明な場合は 'application/octet-stream'とする
                    'ContentType': mime_type if mime_type
                    else 'application/octet-stream',
                }

                try:
                    # ファイルをS3にアップロード
                    s3_client.upload_file(
                        Filename=local_file_path,
                        Bucket=BUCKET_NAME,
                        Key=s3_key,
                        ExtraArgs=extra_args
                    )

                    # アップロードしたファイルのパブリックURLを生成
                    encoded_key = parse.quote(s3_key, safe=EXCLUDE_CHARS)
                    public_url = (
                        f'https://{BUCKET_NAME}.s3.{AWS_REGION}.amazonaws.com/'
                        f'{encoded_key}'
                    )

                    # 結果をファイルに書き込み
                    if check_url_accessible(public_url):
                        success_list.append({
                            'file_name': file_name,
                            'url': public_url
                        })
                    else:
                        failure_list.append({
                            'file_name': file_name,
                            'reason': 'URL疎通確認失敗'
                        })
                except Exception as e:
                    failure_list.append({
                        'file_name': file_name,
                        'reason': f'S3アップロード失敗: {e}'
                    })
        # ファイル出力
        write_results_to_file(success_list, failure_list)

    except FileNotFoundError as e:
        print(f'指定されたファイルが見つかりません: {e}')
    except NoCredentialsError:
        print('AWS認証情報が見つかりません。')
    except PartialCredentialsError:
        print('AWS認証情報が不完全です。')
    except Exception as e:
        print(f'エラーが発生しました: {e}')


# スクリプト実行
if __name__ == '__main__':
    upload_files_to_s3()
