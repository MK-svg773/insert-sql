FROM python:3.11-slim

# 作業ディレクトリの作成
WORKDIR /app

# カレントディレクトリの内容をコンテナへコピー
COPY . /app

# 依存パッケージのインストール
RUN pip install --no-cache-dir pandas openpyxl

# スクリプトを実行（Insert.py をデフォルトの実行コマンドに設定）
CMD ["python", "Insert.py"]
