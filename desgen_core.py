# -*- coding: utf-8 -*-
# コア処理クラス: 商品説明生成
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import openai
import time
from google.oauth2 import service_account
from googleapiclient.discovery import build
import threading

class DescriptionGeneratorCore:
    """
    商品説明生成のコアロジックを管理するクラス。
    通常モードと書籍モードの処理を担当します。
    """
    def __init__(self, credentials_file, openai_api_key, log_callback=None):
        """
        クラスの初期化。
        """
        self.credentials_file = credentials_file
        self.openai_api_key = openai_api_key
        self.log_callback = log_callback or print
        
        openai.api_key = self.openai_api_key
        
        self.scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        self.credentials = service_account.Credentials.from_service_account_file(
            credentials_file, scopes=self.scopes
        )
        self.sheets_service = build("sheets", "v4", credentials=self.credentials)
        
        self.client = gspread.authorize(
            ServiceAccountCredentials.from_json_keyfile_name(credentials_file, self.scopes)
        )
        
        self.batch_size = 20
        self.translation_delay = 1.0
        self.batch_delay = 5.0
        self.stop_flag = False
        
        self.log("✅ Google API & OpenAI サービスの初期化が完了しました。")
    
    def log(self, message):
        if self.log_callback:
            self.log_callback(message)
    
    def stop_processing(self):
        self.stop_flag = True
        self.log("🛑 処理の中断リクエストを受け付けました。")
    
    def column_letter_to_number(self, column_letter):
        column_letter = column_letter.upper()
        result = 0
        for char in column_letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result
    
    def translate_text(self, text, context=""):
        """
        OpenAI APIを使用してテキストを翻訳します。
        """
        if not text or not text.strip():
            return ""
        
        try:
            self.log(f"🔄 翻訳を開始 ({context}): {text[:30]}...")
            time.sleep(self.translation_delay)
            
            system_prompt = f"You are a professional translator. Convert the following Japanese text into natural, fluent English. This text is a {context if context else 'product description'}. Return only the translated text itself, without any additional comments or explanations."
            
            response = openai.ChatCompletion.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"Please translate this into English: {text}"}
                ],
                temperature=0.2,
                max_tokens=1500
            )
            
            result = response["choices"][0]["message"]["content"].strip()
            self.log(f"✅ 翻訳完了 ({context}): {result[:30]}...")
            return result
            
        except Exception as e:
            self.log(f"❌ 翻訳中にエラーが発生 ({context} - {text[:20]}...): {e}")
            time.sleep(2)
            return f"Error: Translation failed. Original text: {text}"

    # --- 共通処理ループ ---
    def _processing_loop(self, spreadsheet_id, sheet_name, column_settings, start_row, data_fetcher, batch_processor):
        """
        両方のモードで共通の処理ループ。
        """
        self.stop_flag = False
        start_time = time.time()
        current_row = start_row
        total_success, total_failed = 0, 0
        consecutive_empty_batches = 0

        while not self.stop_flag:
            batch_data = data_fetcher(spreadsheet_id, sheet_name, current_row, self.batch_size, column_settings)
            
            if not batch_data:
                consecutive_empty_batches += 1
                self.log(f"INFO: 有効なデータが見つかりませんでした。({consecutive_empty_batches}/3)")
                if consecutive_empty_batches >= 3:
                    self.log("🛑 3回連続で有効なデータがなかったため、処理を自動終了します。")
                    break
                current_row += self.batch_size
                continue
            
            consecutive_empty_batches = 0
            results = batch_processor(batch_data, spreadsheet_id, sheet_name, column_settings)
            
            batch_success = sum(1 for r in results if r == "success")
            total_success += batch_success
            total_failed += len(results) - batch_success
            
            current_row += self.batch_size
            
            if not self.stop_flag and batch_data:
                self.log(f"⏳ 次のバッチ処理まで {self.batch_delay}秒 待機します...")
                time.sleep(self.batch_delay)

        total_duration = time.time() - start_time
        self.log("\n" + "="*60 + "\n🎉 全ての処理が完了しました。\n" + "="*60)
        self.log(f"✅ 成功件数: {total_success} 件")
        self.log(f"❌ 失敗件数: {total_failed} 件")
        self.log(f"⏱️  総処理時間: {total_duration:.1f} 秒")
    
    # --- 通常モード ---
    def process_product_descriptions(self, spreadsheet_id, sheet_name, column_settings, start_row=2):
        self.log("="*60 + "\n🚀 [通常モード] 商品説明の生成処理を開始します。\n" + "="*60)
        self._processing_loop(
            spreadsheet_id, sheet_name, column_settings, start_row,
            self.get_normal_mode_batch_data, self.process_normal_mode_batch
        )

    def get_normal_mode_batch_data(self, spreadsheet_id, sheet_name, start_row, batch_size, column_settings):
        end_row = start_row + batch_size - 1
        ranges = [
            f"{sheet_name}!{column_settings['input_col']}{start_row}:{column_settings['input_col']}{end_row}",
            f"{sheet_name}!{column_settings['translated_name_col']}{start_row}:{column_settings['translated_name_col']}{end_row}",
            f"{sheet_name}!{column_settings['jan_code_col']}{start_row}:{column_settings['jan_code_col']}{end_row}",
            f"{sheet_name}!{column_settings['description_col']}{start_row}:{column_settings['description_col']}{end_row}"
        ]
        self.log(f"📊 [通常] スプレッドシートからデータを取得します... (行: {start_row}-{end_row})")
        return self._get_batch_data(spreadsheet_id, ranges, start_row, ['trigger_value', 'translated_name', 'jan_code', 'description'])

    def process_normal_mode_batch(self, batch_data, spreadsheet_id, sheet_name, column_settings):
        sheet = self.client.open_by_key(spreadsheet_id).worksheet(sheet_name)
        output_col_num = self.column_letter_to_number(column_settings['output_col'])
        results = []

        for i, item in enumerate(batch_data):
            if self.stop_flag: break
            self.log(f"\n📄 [通常] {i+1}/{len(batch_data)} 件目 (シート {item['row']} 行目)...")
            
            translated_description = self.translate_text(item["description"], "product description")
            html_content = f"""<p><strong>Product Title:</strong> {item.get('translated_name', 'N/A')}</p>
<p><strong>JAN Code:</strong> {item.get('jan_code', 'N/A')}</p><br>
<p><strong>Description:</strong></p><p>{translated_description}</p>"""
            final_html = self.combine_with_template(html_content)
            
            sheet.update_cell(item["row"], output_col_num, final_html)
            self.log(f"✅ {item['row']}行目の更新完了。")
            results.append("success")
        return results

    # --- 書籍モード ---
    def process_book_descriptions(self, spreadsheet_id, sheet_name, column_settings, start_row=2):
        self.log("="*60 + "\n📚 [書籍モード] 商品説明の生成処理を開始します。\n" + "="*60)
        self._processing_loop(
            spreadsheet_id, sheet_name, column_settings, start_row,
            self.get_book_mode_batch_data, self.process_book_mode_batch
        )

    def get_book_mode_batch_data(self, spreadsheet_id, sheet_name, start_row, batch_size, column_settings):
        end_row = start_row + batch_size - 1
        # GUIから渡されるキーとスプレッドシートの列をマッピング
        col_map = {
            'trigger': 'trigger', 'product_name': 'product_name', 'author': 'author',
            'publisher': 'publisher', 'release_date': 'release_date', 'language': 'language',
            'pages': 'pages', 'isbn10': 'isbn10', 'isbn13': 'isbn13', 'dimensions': 'dimensions'
        }
        ranges = [f"{sheet_name}!{column_settings[key]}{start_row}:{column_settings[key]}{end_row}" for key in col_map]
        self.log(f"📊 [書籍] スプレッドシートからデータを取得します... (行: {start_row}-{end_row})")
        return self._get_batch_data(spreadsheet_id, ranges, start_row, list(col_map.values()))

    def process_book_mode_batch(self, batch_data, spreadsheet_id, sheet_name, column_settings):
        sheet = self.client.open_by_key(spreadsheet_id).worksheet(sheet_name)
        output_col_num = self.column_letter_to_number(column_settings['output'])
        results = []

        for i, item in enumerate(batch_data):
            if self.stop_flag: break
            self.log(f"\n📖 [書籍] {i+1}/{len(batch_data)} 件目 (シート {item['row']} 行目)...")
            
            # 必要な項目を並行して翻訳
            item['translated_author'] = self.translate_text(item.get('author', ''), 'author name')
            item['translated_publisher'] = self.translate_text(item.get('publisher', ''), 'publisher name')
            item['translated_release_date'] = self.translate_text(item.get('release_date', ''), 'date')
            item['translated_language'] = self.translate_text(item.get('language', ''), 'language')
            item['translated_pages'] = self.translate_text(item.get('pages', ''), 'page count')
            
            html_content = self.generate_book_html_description(item)
            final_html = self.combine_with_template(html_content)

            sheet.update_cell(item["row"], output_col_num, final_html)
            self.log(f"✅ {item['row']}行目の更新完了。")
            results.append("success")
        return results

    def generate_book_html_description(self, item_data):
        """書籍データからHTML商品説明を生成します。"""
        # データを安全に取得
        details = {
            "Product Name": item_data.get('product_name', 'N/A'),
            "Author": item_data.get('translated_author', 'N/A'),
            "Publisher": item_data.get('translated_publisher', 'N/A'),
            "Release Date": item_data.get('translated_release_date', 'N/A'),
            "Language": item_data.get('translated_language', 'N/A'),
            "Pages": item_data.get('translated_pages', 'N/A'),
            "ISBN-10": item_data.get('isbn10', 'N/A'),
            "ISBN-13": item_data.get('isbn13', 'N/A'),
            "Dimensions": item_data.get('dimensions', 'N/A'),
        }
        
        html = "<h3>Product Details</h3>\n<table border=\"1\" cellpadding=\"5\" cellspacing=\"0\" style=\"border-collapse: collapse; width: 100%;\">\n"
        for key, value in details.items():
            if value and value.strip() != 'N/A':
                html += f"  <tr>\n    <td style=\"width: 30%; background-color: #f2f2f2;\"><strong>{key}</strong></td>\n    <td>{value}</td>\n  </tr>\n"
        html += "</table>"
        
        self.log(f"✅ [書籍] HTMLを生成しました: {details['Product Name'][:30]}...")
        return html

    # --- 共通ヘルパー ---
    def _get_batch_data(self, spreadsheet_id, ranges, start_row, keys):
        try:
            response = self.sheets_service.spreadsheets().values().batchGet(
                spreadsheetId=spreadsheet_id, ranges=ranges
            ).execute()
            
            value_ranges = response.get('valueRanges', [])
            if not value_ranges: return []

            max_rows = max(len(vr.get('values', [])) for vr in value_ranges)
            batch_data = []

            for i in range(max_rows):
                def get_cell_value(range_index):
                    try:
                        return value_ranges[range_index]['values'][i][0]
                    except (IndexError, KeyError):
                        return ""
                
                # トリガー列(最初の列)に値がある場合のみ処理対象とする
                trigger_value = get_cell_value(0)
                if trigger_value and trigger_value.strip():
                    row_data = {'row': start_row + i}
                    for j, key in enumerate(keys):
                        row_data[key] = get_cell_value(j).strip()
                    batch_data.append(row_data)
            
            self.log(f"✅ データ取得完了。{len(batch_data)}件の有効なデータが見つかりました。")
            return batch_data
            
        except Exception as e:
            self.log(f"❌ データの一括取得中にエラーが発生しました: {e}")
            return []
    
    def combine_with_template(self, description_html):
        """商品説明を指定のHTMLテンプレートに埋め込みます。"""
        return f'''<div id="ds_div">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<p class="sub_tit" style="width:780px; line-height:30px; padding:10px;background-color:#330000; color:#FFFFFF; font-size:35px; font-weight:bold;">Description</p>
{description_html}
<p class="Payment" style="width:780px;line-height:40px;padding:10px;background-color:#330000;color:#FFFFFF;font-size:35px;font-weight:bold;">Payment</p>
<p class="sub_text" style="width:660px;padding:0px 10px;">Please pay within 5 days after the auction closed.</p>
<p class="Shipping" style="width:780px;line-height:40px;padding:10px;background-color:#330000;color:#FFFFFF;font-size:35px;font-weight:bold;">Shipping</p>
<p class="sub_text" style="width:780px;padding:0px 10px;">The products are shipped mainly to the United States.<br>If you wish to ship to other regions, please contact us.</p>
<p class="sub_tit" style="width:780px; line-height:40px;padding:10px;background-color:#330000;color:#FFFFFF;font-size:35px;font-weight:bold;">Returns</p>
<p class="sub_text" style="width:780px;padding:0px 10px;">Returns will ONLY be accepted if the item is not as described.</p>
<p class="sub_tit" style="width:780px;line-height:40px;padding:10px;background-color:#330000;color:#FFFFFF;font-size:35px;font-weight:bold;">International Buyers - Please Note:</p>
<p class="sub_text" style="width:780px;padding:0px 10px;">* Import duties, taxes and charges are not included in the item price or shipping charges. These charges are the buyer's responsibility.<br>* Please check with your country's customs office to determine what these additional costs will be prior to bidding/buying.<br>* These charges are normally collected by the delivering freight (shipping) company or when you pick the item up - do not confuse them for additional shipping charges.<br>* We do not mark merchandise values below value or mark items as "gifts" - US and International government regulations prohibit such behavior.</p>
</div>'''
