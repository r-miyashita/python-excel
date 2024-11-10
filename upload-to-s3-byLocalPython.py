import boto3
from botocore.exceptions import NoCredentialsError, PartialCredentialsError

# IAMユーザーのアクセスキーとシークレットキーを指定
aws_access_key_id = '<あなたのアクセスキー>'
aws_secret_access_key = '<あなたのシークレットキー>'

# S3のバケット名とアップロードするファイルのパスを設定
bucket_name = 'mybucket-prod'  # S3バケット名
file_path = 'sample.png'       # アップロードするローカルファイルのパス
object_name = 'public/sample.png'  # S3上でのオブジェクト名

# S3クライアントの作成（認証情報を直接設定）
s3_client = boto3.client(
    's3',
    aws_access_key_id=aws_access_key_id,
    aws_secret_access_key=aws_secret_access_key
)

# ファイルをS3バケットにアップロード
try:
    s3_client.upload_file(
        file_path,
        bucket_name,
        object_name,
        ExtraArgs={
            'ACL': 'public-read',
            'ContentType': 'image/png'
        }
    )
    print(f'ファイル {file_path} を S3バケット {bucket_name} にアップロードしました。')
except FileNotFoundError:
    print(f'指定されたファイル {file_path} が見つかりません。')
except NoCredentialsError:
    print('AWSの認証情報が見つかりません。')
except PartialCredentialsError:
    print('AWS認証情報が不完全です。')
except Exception as e:
    print(f'エラーが発生しました: {str(e)}')
