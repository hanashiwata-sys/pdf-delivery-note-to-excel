# PDF Delivery Note to Excel

Pythonで納品書PDFを解析し、
Excel・CSVへの出力と商品別集計を行うツールです。

## 機能

- 複数PDFの一括処理
- 日付抽出
- 会社名抽出
- 商品情報抽出
- Excel出力
- CSV出力
- 商品別集計

## サンプルデータ

sample_pdfフォルダに動作確認用のサンプルPDFを格納しています。

これらのPDFをPDF(納品書)フォルダへ配置して実行できます。

## 使用ライブラリ

- pdfplumber
- pandas
- openpyxl

## 実行方法

### 1. 必要なライブラリをインストール

'''bash
pip install -r requirements.txt
'''

### 2. PDFを配置

PDF(納品書)フォルダの中に納品書PDFを入れます。

### 3. プログラム実行

'''bash
python main.py
'''

### 4. 出力ファイル

実行後、以下のファイルが作成されます。

- 納品書一覧.xlsx
- 納品書一覧.csv
- 商品別集計.xlsx
