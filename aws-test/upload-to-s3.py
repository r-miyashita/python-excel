import os
import boto3
from botocore.exceptions import NoCredentialsError, PartialCredentialsError
from dotenv import load_dotenv
import urllib.parse
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

# ファイルリスト.txt と リソースフォルダ の指定
UPLOAD_FILE_LIST_PATH = './upload_file_list.txt'
FILES_DIR = './files/'


def upload_files_to_s3():
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
                    print(f'ローカルファイルが見つかりません：{local_file_path}')
                    continue

                # MIMEタイプを取得 #mime_type の例: 'image/jpeg'
                mime_type, _ = mimetypes.guess_type(local_file_path)

                # ExtraArgsを作成
                extra_args = {'ACL': 'public-read'}

                # MIMEタイプを ContentType として指定し、 ExtraArgs に追加
                if mime_type:
                    extra_args['ContentType'] = mime_type
                else:
                    extra_args['ContentType'] = 'application/octet-stream'

                # ファイルをS3にアップロード
                s3_client.upload_file(
                    Filename=local_file_path,
                    Bucket=BUCKET_NAME,
                    Key=s3_key,
                    ExtraArgs=extra_args
                )

                # アップロードしたファイルのパブリックURLを生成
                encoded_key = urllib.parse.quote(s3_key)
                public_url = (
                    f"https://{BUCKET_NAME}.s3.{AWS_REGION}.amazonaws.com/"
                    f"{encoded_key}"
                )
                print(f"アップロード成功: {public_url}")

                #
                # encode は safe を使ってs3仕様に近い仕様にする
                # encoded_key = urllib.parse.quote(key, safe='/-_.~')
                #

    except FileNotFoundError as e:
        print(
            f"指定されたファイルが見つかりません: {e}"
        )
    except NoCredentialsError:
        print("AWS認証情報が見つかりません。")
    except PartialCredentialsError:
        print("AWS認証情報が不完全です。")
    except Exception as e:
        print(f"エラーが発生しました: {e}")


# スクリプト実行
if __name__ == '__main__':
    upload_files_to_s3()
