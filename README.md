# AI Description Generator

## 概要
このリポジトリは、**AIを活用した商品説明文生成ツール**です。  
ISBNやJANコードを基に書誌情報を取得し、OpenAI APIを使って **自然な英語タイトルやHTML商品説明文** を自動生成します。  
特に eBay などの越境EC向け商品登録で活用できます。

## 主な機能
- **書誌情報取得**  
  - openBD / NDL APIから書籍データを取得  
- **AI翻訳・説明文生成**  
  - OpenAI APIを利用し、自然な英語タイトルを80文字以内で翻訳  
  - HTML形式の説明文を自動生成（カラフルなテンプレート対応）  
- **Google Sheets連携**  
  - 翻訳結果や説明文をスプレッドシートへ直接出力  
- **バッチ処理**  
  - 未処理データのみを効率的に処理  

## 使用技術
- Python 3.10+  
- gspread / Google Sheets API  
- OpenAI API  

## 成果・効果
- 手作業での翻訳・説明文作成を自動化  
- 国際マーケット対応（越境EC向けの販売支援）  
- 説明文の品質を均一化し、効率的に商品ページを作成可能  

## 実行方法
1. Google Service Account の認証JSONを準備  
2. `descriptiongen.py` または `titlegen.py` を実行  

```bash
python descriptiongen.py
