# Excel CSV Master Data Matcher

注文データと商品マスタを照合し，マスタ情報を補完した一覧と，不一致データを出力するPythonツールです．

`input` フォルダに配置した注文データを読み込み，`master` フォルダに配置した商品マスタと照合します．  
商品コードがマスタに存在する行は，商品名，カテゴリ，標準単価などを補完して `enriched_orders.xlsx` に出力します．  
商品コードがマスタに存在しない行や，注文データ自体に不備がある行は，`unmatched_orders.xlsx` に分離します．  
また，照合件数やエラー件数を確認できる `match_summary.xlsx` も出力します．

このツールは，単なるExcel出力ツールではなく，Excel業務でよく行われるマスタ照合，データ補完，不一致抽出を自動化するためのツールです．

---

## 主な機能

* 注文CSVファイルの読み込み
* 注文Excelファイルの読み込み
* 商品マスタCSVファイルの読み込み
* 商品マスタExcelファイルの読み込み
* 注文データと商品マスタの照合
* マスタ情報の商品名・カテゴリ・標準単価の補完
* 数量と標準単価による金額計算
* マスタに存在しない商品コードの抽出
* 注文データの必須項目チェック
* 注文データの数値チェック
* 注文データの日付形式変換
* 商品マスタの必須項目チェック
* 商品マスタの商品コード重複チェック
* 商品マスタの商品コード空欄チェック
* 補完済みデータのExcel出力
* 不一致データのExcel出力
* 照合結果サマリーのExcel出力
* 初回実行用サンプルファイルの自動作成
* 日本語CSV向けの文字コード対応

---

## 使用技術

* Python
* pandas
* openpyxl

---

## 想定用途

このツールは，以下のような作業を自動化するためのものです．

* 注文データと商品マスタを照合する
* 商品コードから商品名やカテゴリを自動補完する
* 商品マスタに存在しない商品コードを抽出する
* VLOOKUP / XLOOKUP のような照合作業を自動化する
* EC注文データに商品マスタ情報を付与する
* 売上明細に商品分類や標準単価を付与する
* 取引データとマスタデータの不一致を確認する
* 手作業のマスタ照合ミスを減らす
* CSV・Excel業務自動化案件の補助ツールとして使う

---

## ディレクトリ構成

```text
excel-csv-master-data-matcher/
├─ excel_csv_master_data_matcher.py
├─ config.json
├─ requirements.txt
├─ README.md
├─ input/
│  └─ sample_orders.xlsx
├─ master/
│  └─ product_master.xlsx
└─ output/
   ├─ enriched_orders.xlsx
   ├─ unmatched_orders.xlsx
   └─ match_summary.xlsx
```

### フォルダの役割

| フォルダ | 内容 |
|---|---|
| `input` | 照合対象の注文CSV・Excelファイルを配置するフォルダ |
| `master` | 商品マスタCSV・Excelファイルを配置するフォルダ |
| `output` | 照合結果，不一致データ，サマリーの出力先 |

---

## 実行方法

### 1. リポジトリを取得

```bash
git clone <repository_url>
```

### 2. ディレクトリへ移動

```bash
cd excel-csv-master-data-matcher
```

### 3. 必要ライブラリをインストール

```bash
pip install -r requirements.txt
```

### 4. 注文データを配置

照合したい注文データを `input` フォルダに配置します．

対応形式は以下です．

```text
.csv
.xlsx
```

例：

```text
input/sample_orders.xlsx
```

CSVファイルでも使用できます．

```text
input/sample_orders.csv
```

### 5. 商品マスタを配置

商品マスタを `master` フォルダに配置します．

例：

```text
master/product_master.xlsx
```

CSVファイルでも使用できます．

```text
master/product_master.csv
```

### 6. config.json を確認

`config.json` で，注文ファイル名，商品マスタファイル名，照合キー，出力列，チェック対象列などを設定します．

### 7. プログラムを実行

```bash
python excel_csv_master_data_matcher.py
```

または

```bash
py excel_csv_master_data_matcher.py
```

### 8. 出力ファイルを確認

実行後，`output` フォルダに以下のファイルが出力されます．

```text
output/
├─ enriched_orders.xlsx
├─ unmatched_orders.xlsx
└─ match_summary.xlsx
```

---

## 初回実行について

初回実行時，以下のファイルやフォルダが存在しない場合は自動作成されます．

```text
config.json
input/sample_orders.xlsx
master/product_master.xlsx
output/
```

そのため，最初は自分で入力ファイルを用意しなくても，以下のコマンドだけで動作確認できます．

```bash
python excel_csv_master_data_matcher.py
```

サンプルデータには，正常に照合できる行，商品マスタに存在しない行，商品コードが空欄の行，数量が数値ではない行が含まれています．  
これにより，補完済みデータと不一致データが分かれて出力される動作を確認できます．

---

## 入力データの形式

## 注文データ

注文データは，CSVまたはExcelの表形式を想定しています．

* 1行目に列名があること
* 2行目以降に注文データがあること
* Excelファイルの場合，最初のシートを読み込むこと
* `config.json` の `order_required_columns` に指定した列が存在すること
* `order_key` に指定した列が存在すること

サンプルの注文データ：

| order_id | order_date | product_code | quantity | note |
|---|---|---|---:|---|
| ORD-001 | 2026/06/12 | P-1001 | 2 | 正常データ |
| ORD-002 | 2026/06/13 | P-1002 | 1 | 正常データ |
| ORD-003 | 2026/06/14 | P-9999 | 3 | マスタに存在しない商品コード |
| ORD-004 | 2026/06/15 |  | 1 | 商品コードが空欄 |
| ORD-005 | 2026/06/16 | P-1003 | abc | 数量が数値ではない |
| ORD-006 | 2026/06/17 | P-1004 | 2 | 正常データ |

---

## 商品マスタ

商品マスタは，CSVまたはExcelの表形式を想定しています．

* 1行目に列名があること
* 2行目以降に商品マスタがあること
* Excelファイルの場合，最初のシートを読み込むこと
* `config.json` の `master_required_columns` に指定した列が存在すること
* `master_key` に指定した列が存在すること
* `master_key` の値が空欄でないこと
* `master_key` の値が重複していないこと

サンプルの商品マスタ：

| product_code | product_name | category | standard_price |
|---|---|---|---:|
| P-1001 | ワイヤレスマウス | PC周辺機器 | 1800 |
| P-1002 | USBキーボード | PC周辺機器 | 3200 |
| P-1003 | Webカメラ | PC周辺機器 | 4500 |
| P-1004 | 外付けSSD | ストレージ | 9800 |

---

## config.json

動作設定は `config.json` で変更できます．

### サンプル設定

```json
{
  "order_file": "sample_orders.xlsx",
  "master_file": "product_master.xlsx",
  "enriched_file": "enriched_orders.xlsx",
  "unmatched_file": "unmatched_orders.xlsx",
  "summary_file": "match_summary.xlsx",
  "order_key": "product_code",
  "master_key": "product_code",
  "order_required_columns": [
    "order_id",
    "order_date",
    "product_code",
    "quantity"
  ],
  "master_required_columns": [
    "product_code",
    "product_name",
    "category",
    "standard_price"
  ],
  "master_value_columns": [
    "product_name",
    "category",
    "standard_price"
  ],
  "order_numeric_columns": [
    "quantity"
  ],
  "master_numeric_columns": [
    "standard_price"
  ],
  "order_date_columns": [
    "order_date"
  ],
  "date_format": "%Y-%m-%d",
  "amount_calculation": {
    "enabled": true,
    "quantity_column": "quantity",
    "price_column": "standard_price",
    "output_column": "expected_amount"
  },
  "enriched_output_columns": [
    "order_id",
    "order_date",
    "product_code",
    "product_name",
    "category",
    "quantity",
    "standard_price",
    "expected_amount"
  ],
  "unmatched_output_columns": [
    "row_number",
    "order_id",
    "order_date",
    "product_code",
    "quantity",
    "error_type",
    "error_message"
  ]
}
```

---

## 設定項目

| 項目 | 内容 |
|---|---|
| `order_file` | `input` フォルダ内で読み込む注文データファイル |
| `master_file` | `master` フォルダ内で読み込む商品マスタファイル |
| `enriched_file` | マスタ情報を補完した結果を出力するExcelファイル |
| `unmatched_file` | 不一致データやエラー行を出力するExcelファイル |
| `summary_file` | 照合結果サマリーを出力するExcelファイル |
| `order_key` | 注文データ側の照合キー列 |
| `master_key` | 商品マスタ側の照合キー列 |
| `order_required_columns` | 注文データで空欄を許可しない列 |
| `master_required_columns` | 商品マスタで空欄を許可しない列 |
| `master_value_columns` | 商品マスタから補完する列 |
| `order_numeric_columns` | 注文データで数値として扱う列 |
| `master_numeric_columns` | 商品マスタで数値として扱う列 |
| `order_date_columns` | 注文データで日付として扱う列 |
| `date_format` | 日付の出力形式 |
| `amount_calculation` | 金額計算の設定 |
| `enriched_output_columns` | 補完済みデータの出力列順 |
| `unmatched_output_columns` | 不一致データの出力列順 |

---

## 照合キーについて

このツールでは，`order_key` と `master_key` に指定した列を使って，注文データと商品マスタを照合します．

標準設定では，どちらも `product_code` です．

```json
{
  "order_key": "product_code",
  "master_key": "product_code"
}
```

この場合，以下のように照合します．

```text
注文データ.product_code
        ↓
商品マスタ.product_code
```

商品コードが一致した場合，商品マスタから `product_name`，`category`，`standard_price` などの情報を補完します．

---

## master_value_columns について

`master_value_columns` は，商品マスタから注文データへ補完する列を指定する設定です．

例：

```json
{
  "master_value_columns": [
    "product_name",
    "category",
    "standard_price"
  ]
}
```

この場合，商品コードが一致した注文行に対して，商品マスタの以下の情報が追加されます．

```text
product_name
category
standard_price
```

---

## amount_calculation について

`amount_calculation` は，数量と単価から金額を計算する設定です．

例：

```json
{
  "amount_calculation": {
    "enabled": true,
    "quantity_column": "quantity",
    "price_column": "standard_price",
    "output_column": "expected_amount"
  }
}
```

この設定では，以下の計算を行います．

```text
expected_amount = quantity × standard_price
```

例：

| quantity | standard_price | expected_amount |
|---:|---:|---:|
| 2 | 1800 | 3600 |
| 1 | 3200 | 3200 |
| 2 | 9800 | 19600 |

金額計算が不要な場合は，以下のように `enabled` を `false` にします．

```json
{
  "amount_calculation": {
    "enabled": false
  }
}
```

---

## 出力ファイル

## enriched_orders.xlsx

商品マスタと一致した注文データを出力するExcelファイルです．

例：

```text
output/enriched_orders.xlsx
```

出力例：

| order_id | order_date | product_code | product_name | category | quantity | standard_price | expected_amount |
|---|---|---|---|---|---:|---:|---:|
| ORD-001 | 2026-06-12 | P-1001 | ワイヤレスマウス | PC周辺機器 | 2 | 1800 | 3600 |
| ORD-002 | 2026-06-13 | P-1002 | USBキーボード | PC周辺機器 | 1 | 3200 | 3200 |
| ORD-006 | 2026-06-17 | P-1004 | 外付けSSD | ストレージ | 2 | 9800 | 19600 |

このファイルは，商品マスタの情報を補完した完成データとして使うことを想定しています．

---

## unmatched_orders.xlsx

商品マスタに存在しなかった行や，注文データ自体に不備がある行を出力するExcelファイルです．

例：

```text
output/unmatched_orders.xlsx
```

出力例：

| row_number | order_id | order_date | product_code | quantity | error_type | error_message |
|---:|---|---|---|---|---|---|
| 4 | ORD-003 | 2026-06-14 | P-9999 | 3 | master_not_found | product_code not found in master |
| 5 | ORD-004 | 2026-06-15 |  | 1 | validation_error | product_code is required |
| 6 | ORD-005 | 2026-06-16 | P-1003 | abc | validation_error | quantity must be numeric |

### 各列の意味

| 列名 | 内容 |
|---|---|
| `row_number` | 入力ファイル上の行番号 |
| `error_type` | エラーの種類 |
| `error_message` | エラー内容 |
| その他の列 | 注文データの内容 |

`row_number` は，入力ファイル上の行番号です．  
1行目はヘッダーとして扱うため，データ1行目は `row_number = 2` になります．

---

## error_type の種類

| error_type | 内容 |
|---|---|
| `validation_error` | 注文データ自体に不備がある |
| `master_not_found` | 注文データの商品コードが商品マスタに存在しない |

### validation_error の例

```text
product_code is required
quantity must be numeric
order_date must be a valid date
```

### master_not_found の例

```text
product_code not found in master
```

---

## match_summary.xlsx

照合結果の概要を出力するExcelファイルです．

例：

```text
output/match_summary.xlsx
```

このファイルには，以下のシートが含まれます．

| シート名 | 内容 |
|---|---|
| `summary` | 入力ファイル名，出力ファイル名，照合件数，開始終了時刻 |
| `settings` | 照合キー，必須列，数値列，日付列などの設定 |
| `master_check` | 商品マスタの行数，キー空欄，キー重複の確認結果 |

### summary シートの例

| item | value |
|---|---|
| order_file | sample_orders.xlsx |
| master_file | product_master.xlsx |
| enriched_file | enriched_orders.xlsx |
| unmatched_file | unmatched_orders.xlsx |
| summary_file | match_summary.xlsx |
| total_order_rows | 6 |
| matched_rows | 3 |
| unmatched_rows | 3 |
| validation_error_rows | 2 |
| master_not_found_rows | 1 |

---

## 照合結果の考え方

このツールでは，注文データを以下のように分類します．

```text
注文データに問題がなく，商品マスタにも存在する行
        → enriched_orders.xlsx

注文データ自体に不備がある行
        → unmatched_orders.xlsx

注文データに問題はないが，商品マスタに存在しない行
        → unmatched_orders.xlsx
```

たとえば，サンプルデータでは以下の結果になります．

```text
注文データ件数: 6
一致件数: 3
不一致件数: 3
入力データ不備: 2
マスタ未登録: 1
```

---

## CSV文字コードについて

CSV入力時は，以下の文字コードを順番に試します．

```text
utf-8-sig
utf-8
cp932
```

これにより，日本語を含むCSVでも読み込める可能性が高くなります．

---

## 注意事項

* 現在は，`config.json` で指定した1つの注文ファイルと1つの商品マスタを処理します．
* 入力ファイルは `.csv` / `.xlsx` に対応しています．
* 出力ファイルは `.xlsx` 形式です．
* `.xls` 形式には対応していません．
* パスワード付きExcelや壊れたExcelファイルには対応していません．
* Excel入力ファイルは最初のシートを読み込みます．
* 入力ファイルの1行目は列名として扱います．
* 注文データの照合キー列は，`order_required_columns` に含める必要があります．
* 商品マスタの照合キー列は，`master_required_columns` に含める必要があります．
* 商品マスタの照合キーに空欄や重複がある場合，処理全体を停止します．
* 商品マスタは正しい基準データとして扱います．
* `master_value_columns` に `master_key` と同じ列を含めることはできません．
* `amount_calculation` を有効にする場合，数量列は `order_numeric_columns`，単価列は `master_numeric_columns` に含める必要があります．
* `output` フォルダ内の生成ファイルは実行結果なので，Git管理対象から外しても問題ありません．
* `config.json` やサンプルファイルが既に存在する場合，自動作成処理では上書きされません．

---

## 活用方法

* 注文データと商品マスタの照合
* 商品コードから商品名を自動補完
* 商品分類や標準単価の自動付与
* 商品マスタ未登録データの抽出
* 売上明細と商品マスタの突き合わせ
* EC注文データの整形
* 在庫管理用データの作成
* 物流システム向けデータの事前確認
* VLOOKUP / XLOOKUP 作業の自動化
* CSV・Excel業務自動化案件のポートフォリオ

---

## Requirements

```text
pandas
openpyxl
```

---

# English

## Overview

Excel CSV Master Data Matcher is a Python tool that matches order data with product master data and exports enriched rows and unmatched rows.

The tool reads an order CSV or Excel file from the `input` folder and a product master CSV or Excel file from the `master` folder.  
Rows whose product codes exist in the master data are enriched with product name，category，and standard price.  
Rows with validation errors or product codes not found in the master data are exported separately.

---

## Features

* Read order CSV files
* Read order Excel files
* Read product master CSV files
* Read product master Excel files
* Match order data with product master data
* Enrich matched rows with master values
* Extract product codes not found in master data
* Validate required fields
* Validate numeric fields
* Normalize date fields
* Check blank master keys
* Check duplicated master keys
* Calculate expected amount
* Export enriched rows to Excel
* Export unmatched rows to Excel
* Export matching summary to Excel
* Create sample files on first run
* Support common Japanese CSV encodings

---

## Directory Structure

```text
excel-csv-master-data-matcher/
├─ excel_csv_master_data_matcher.py
├─ config.json
├─ requirements.txt
├─ README.md
├─ input/
│  └─ sample_orders.xlsx
├─ master/
│  └─ product_master.xlsx
└─ output/
   ├─ enriched_orders.xlsx
   ├─ unmatched_orders.xlsx
   └─ match_summary.xlsx
```

---

## How to Use

### 1. Install requirements

```bash
pip install -r requirements.txt
```

### 2. Place order data

Place an order CSV or Excel file in the `input` folder.

```text
input/sample_orders.xlsx
```

### 3. Place master data

Place a product master CSV or Excel file in the `master` folder.

```text
master/product_master.xlsx
```

### 4. Edit config.json

Set the order file name，master file name，matching keys，required columns，numeric columns，date columns，and output columns.

### 5. Run the script

```bash
python excel_csv_master_data_matcher.py
```

---

## Output

Generated files are saved in the `output` folder.

```text
output/enriched_orders.xlsx
output/unmatched_orders.xlsx
output/match_summary.xlsx
```

### enriched_orders.xlsx

This file contains rows successfully matched with product master data.

### unmatched_orders.xlsx

This file contains rows with validation errors or product codes not found in the master data.

### match_summary.xlsx

This file contains a summary of the matching result，settings，and master data checks.

---

## Limitations

* This version processes one order file and one master file specified in `config.json`.
* Input files must be `.csv` or `.xlsx`.
* `.xls` files are not supported.
* Password-protected or broken Excel files are not supported.
* For Excel input，only the first sheet is read.
* The master key must not be blank.
* The master key must not be duplicated.

---

## Use Cases

* Match order data with product master data
* Enrich product names and categories
* Extract unknown product codes
* Automate VLOOKUP-like Excel tasks
* Prepare sales or inventory data
* Clean and validate business data
* CSV / Excel workflow automation
