# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import threading
import os
import json
from datetime import datetime
# 修正後のコアファイル 'desgen_core.py' からクラスをインポート
from desgen_core import DescriptionGeneratorCore

class DescriptionGeneratorGUI:
    """
    商品説明生成ツールのためのTkinter GUIアプリケーション。
    通常モードと書籍モードの2つの機能を提供します。
    """
    def __init__(self, root):
        self.root = root
        self.root.title("商品説明生成ツール (v3.1)")
        self.root.geometry("850x750")

        self.processor = None
        self.is_processing = False
        # 新しいバージョン用の設定ファイル
        self.config_file = "description_generator_config_v3.json"

        self.create_widgets()
        self.load_config()

    def create_widgets(self):
        """GUIのウィジェットをすべて作成・配置します。"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # --- API & スプレッドシート設定 (共通) ---
        api_frame = self._create_section(main_frame, "共通設定", 0)
        self.credentials_var = tk.StringVar()
        self._create_file_input(api_frame, "Google認証ファイル:", 0, self.credentials_var)
        self.openai_api_key_var = tk.StringVar()
        self._create_entry(api_frame, "OpenAI APIキー:", 1, self.openai_api_key_var, show="*")
        self.spreadsheet_id_var = tk.StringVar()
        self._create_entry(api_frame, "スプレッドシートID:", 2, self.spreadsheet_id_var)
        self.sheet_name_var = tk.StringVar(value="集計")
        self._create_entry(api_frame, "シート名:", 3, self.sheet_name_var, width=30)

        # --- モード選択タブ ---
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=1, column=0, sticky="ew", pady=10)

        self.normal_mode_frame = ttk.Frame(self.notebook, padding="10")
        self.book_mode_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.normal_mode_frame, text="通常モード")
        self.notebook.add(self.book_mode_frame, text="書籍モード")

        # --- 各モードのウィジェット作成 ---
        self._create_normal_mode_widgets()
        self._create_book_mode_widgets()

        # --- 処理設定 (共通) ---
        processing_frame = self._create_section(main_frame, "処理設定", 2)
        # gridレイアウト用に列の重みを設定
        processing_frame.columnconfigure((1, 3, 5, 7), weight=1)

        self.start_row_var = tk.StringVar(value="2")
        self._create_entry(processing_frame, "開始行:", 0, self.start_row_var, width=8)
        self.batch_size_var = self._create_combobox(processing_frame, "バッチサイズ:", 2, ["10", "15", "20", "25", "30"], "20")
        self.translation_delay_var = self._create_combobox(processing_frame, "翻訳間隔(秒):", 4, ["0.5", "1.0", "1.5", "2.0"], "1.0")
        self.batch_delay_var = self._create_combobox(processing_frame, "バッチ間隔(秒):", 6, ["3.0", "5.0", "10.0"], "5.0")


        # --- ボタン (共通) ---
        button_frame = ttk.Frame(main_frame, padding=(0, 10))
        button_frame.grid(row=3, column=0, sticky="ew")
        self.start_button = ttk.Button(button_frame, text="処理開始", command=self.start_processing, style="Accent.TButton")
        self.start_button.pack(side="left", padx=5)
        self.stop_button = ttk.Button(button_frame, text="処理停止", command=self.stop_processing, state="disabled")
        self.stop_button.pack(side="left", padx=5)
        ttk.Button(button_frame, text="設定保存", command=self.save_config).pack(side="left", padx=5)
        ttk.Button(button_frame, text="設定読込", command=self.load_config).pack(side="left", padx=5)
        ttk.Button(button_frame, text="接続テスト", command=self.test_connection).pack(side="left", padx=5)

        # --- ログエリア (共通) ---
        log_frame = self._create_section(main_frame, "ログ", 4)
        main_frame.rowconfigure(4, weight=1)
        self.log_text = scrolledtext.ScrolledText(log_frame, width=90, height=15, wrap=tk.WORD, relief="solid", bd=1)
        self.log_text.pack(expand=True, fill="both", padx=5, pady=5)
        ttk.Button(log_frame, text="ログクリア", command=self.clear_log).pack(anchor="w", padx=5, pady=(0,5))

        style = ttk.Style()
        style.configure("Accent.TButton", foreground="white", background="#0078D7")

    def _create_normal_mode_widgets(self):
        """通常モードの列設定ウィジェットを作成します。"""
        frame = ttk.LabelFrame(self.normal_mode_frame, text="列設定", padding="10")
        frame.pack(fill="x")
        
        self.nm_input_col_var = tk.StringVar(value="A")
        self._create_entry(frame, "トリガー列:", 0, self.nm_input_col_var, width=10)
        self.nm_translated_name_col_var = tk.StringVar(value="K")
        self._create_entry(frame, "翻訳済商品名列:", 1, self.nm_translated_name_col_var, width=10)
        self.nm_jan_code_col_var = tk.StringVar(value="L")
        self._create_entry(frame, "JANコード列:", 2, self.nm_jan_code_col_var, width=10)
        self.nm_description_col_var = tk.StringVar(value="I")
        self._create_entry(frame, "商品説明列(翻訳元):", 3, self.nm_description_col_var, width=10)
        self.nm_output_col_var = tk.StringVar(value="Q")
        self._create_entry(frame, "出力列:", 4, self.nm_output_col_var, width=10)

    def _create_book_mode_widgets(self):
        """書籍モードの列設定ウィジェットを作成します。"""
        frame = ttk.LabelFrame(self.book_mode_frame, text="列設定", padding="10")
        frame.pack(fill="x")

        # 2列レイアウト用の親フレーム
        left_frame = ttk.Frame(frame)
        left_frame.grid(row=0, column=0, sticky="ns", padx=5)
        right_frame = ttk.Frame(frame)
        right_frame.grid(row=0, column=1, sticky="ns", padx=5)

        # 変数とウィジェットの定義
        self.bm_vars = {
            "trigger": (tk.StringVar(value="A"), "トリガー列:", left_frame, 0),
            "product_name": (tk.StringVar(value="B"), "商品名列:", left_frame, 1),
            "author": (tk.StringVar(value="C"), "著者列 (翻訳元):", left_frame, 2),
            "publisher": (tk.StringVar(value="D"), "出版社列 (翻訳元):", left_frame, 3),
            "release_date": (tk.StringVar(value="E"), "発売日列 (翻訳元):", left_frame, 4),
            "language": (tk.StringVar(value="F"), "言語列 (翻訳元):", right_frame, 0),
            "pages": (tk.StringVar(value="G"), "ページ数 (翻訳元):", right_frame, 1),
            "isbn10": (tk.StringVar(value="H"), "ISBN-10列:", right_frame, 2),
            "isbn13": (tk.StringVar(value="I"), "ISBN-13列:", right_frame, 3),
            "dimensions": (tk.StringVar(value="J"), "寸法列:", right_frame, 4),
            "output": (tk.StringVar(value="K"), "出力列:", right_frame, 5),
        }

        for key, (var, label, parent, row) in self.bm_vars.items():
            self._create_entry(parent, label, row, var, width=15)

    def _create_section(self, parent, text, row):
        frame = ttk.LabelFrame(parent, text=text, padding="10")
        frame.grid(row=row, column=0, sticky="ew", pady=5, padx=5)
        return frame

    def _create_entry(self, parent, label_text, row, var, **kwargs):
        ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky="w", pady=2, padx=5)
        ttk.Entry(parent, textvariable=var, **kwargs).grid(row=row, column=1, sticky="w", pady=2, padx=5)

    def _create_file_input(self, parent, label_text, row, var):
        ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky="w", pady=2, padx=5)
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=1, sticky="ew")
        frame.columnconfigure(0, weight=1)
        entry = ttk.Entry(frame, textvariable=var)
        entry.grid(row=0, column=0, sticky="ew")
        ttk.Button(frame, text="参照...", command=lambda: self.browse_credentials(var)).grid(row=0, column=1, padx=5)

    def _create_combobox(self, parent, label_text, col, values, default_value):
        """
        Comboboxウィジェットを作成し、gridで配置します。
        """
        var = tk.StringVar(value=default_value)
        ttk.Label(parent, text=label_text).grid(row=0, column=col, sticky="w", padx=(10, 2))
        combo = ttk.Combobox(parent, textvariable=var, values=values, width=6, state="readonly")
        combo.grid(row=0, column=col + 1, sticky="w", padx=(0, 10))
        return var

    def browse_credentials(self, var):
        filename = filedialog.askopenfilename(title="Google API認証ファイルを選択", filetypes=[("JSON files", "*.json")])
        if filename: var.set(filename)

    def log_message(self, message):
        self.root.after(0, self._append_log, f"[{datetime.now():%H:%M:%S}] {message}\n")

    def _append_log(self, message):
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def validate_settings(self):
        if not os.path.exists(self.credentials_var.get()): return "Google認証ファイルが見つかりません。"
        if not self.openai_api_key_var.get(): return "OpenAI APIキーを入力してください。"
        if not self.spreadsheet_id_var.get(): return "スプレッドシートIDを入力してください。"
        try: int(self.start_row_var.get())
        except ValueError: return "開始行は半角数値を入力してください。"
        return None

    def start_processing(self):
        error = self.validate_settings()
        if error:
            messagebox.showerror("設定エラー", error)
            return
        if self.is_processing:
            messagebox.showwarning("処理中", "既に処理が実行中です。")
            return
        
        self.is_processing = True
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.clear_log()
        
        threading.Thread(target=self._run_processing, daemon=True).start()

    def stop_processing(self):
        if self.processor:
            self.processor.stop_processing()
        self.stop_button.config(state="disabled")

    def _run_processing(self):
        try:
            self.processor = DescriptionGeneratorCore(
                self.credentials_var.get(), self.openai_api_key_var.get(), self.log_message
            )
            self.processor.batch_size = int(self.batch_size_var.get())
            self.processor.translation_delay = float(self.translation_delay_var.get())
            self.processor.batch_delay = float(self.batch_delay_var.get())
            
            # 選択中のタブに応じて処理を分岐
            selected_tab_index = self.notebook.index(self.notebook.select())
            
            if selected_tab_index == 0: # 通常モード
                self.log_message("🚀 通常モードで処理を開始します。")
                column_settings = {
                    'input_col': self.nm_input_col_var.get(),
                    'translated_name_col': self.nm_translated_name_col_var.get(),
                    'jan_code_col': self.nm_jan_code_col_var.get(),
                    'description_col': self.nm_description_col_var.get(),
                    'output_col': self.nm_output_col_var.get()
                }
                self.processor.process_product_descriptions(
                    self.spreadsheet_id_var.get(), self.sheet_name_var.get(),
                    column_settings, int(self.start_row_var.get())
                )
            else: # 書籍モード
                self.log_message("📚 書籍モードで処理を開始します。")
                column_settings = {key: var.get() for key, (var, _, _, _) in self.bm_vars.items()}
                self.processor.process_book_descriptions(
                    self.spreadsheet_id_var.get(), self.sheet_name_var.get(),
                    column_settings, int(self.start_row_var.get())
                )

        except Exception as e:
            self.log_message(f"❌ 致命的なエラーが発生しました: {e}")
            messagebox.showerror("実行時エラー", f"処理中に予期せぬエラーが発生しました:\n{e}")
        finally:
            self.root.after(0, self._reset_ui)

    def _reset_ui(self):
        self.is_processing = False
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.processor = None

    def test_connection(self):
        error = self.validate_settings()
        if error:
            messagebox.showerror("設定エラー", error)
            return
        
        self.log_message("🔍 接続テストを開始します...")
        def run_test():
            try:
                core = DescriptionGeneratorCore(self.credentials_var.get(), self.openai_api_key_var.get(), self.log_message)
                self.log_message("...Google Sheetsに接続中...")
                core.client.open_by_key(self.spreadsheet_id_var.get()).worksheet(self.sheet_name_var.get())
                self.log_message("✅ Google Sheets接続: 成功")
                
                self.log_message("...OpenAI APIに接続中...")
                if core.translate_text("これは接続テストです。"):
                    self.log_message("✅ OpenAI API接続: 成功")
                else: self.log_message("❌ OpenAI API接続: 失敗。APIキーまたはネットワーク設定を確認してください。")
            except Exception as e:
                self.log_message(f"❌ 接続テスト失敗: {e}")
        threading.Thread(target=run_test, daemon=True).start()

    def get_config_as_dict(self):
        """現在のGUI設定をすべて辞書として取得します。"""
        config = {
            'credentials_var': self.credentials_var.get(),
            'openai_api_key_var': self.openai_api_key_var.get(),
            'spreadsheet_id_var': self.spreadsheet_id_var.get(),
            'sheet_name_var': self.sheet_name_var.get(),
            'start_row_var': self.start_row_var.get(),
            'batch_size_var': self.batch_size_var.get(),
            'translation_delay_var': self.translation_delay_var.get(),
            'batch_delay_var': self.batch_delay_var.get(),
            'selected_tab': self.notebook.index(self.notebook.select()),
            'normal_mode': {
                'input_col': self.nm_input_col_var.get(),
                'translated_name_col': self.nm_translated_name_col_var.get(),
                'jan_code_col': self.nm_jan_code_col_var.get(),
                'description_col': self.nm_description_col_var.get(),
                'output_col': self.nm_output_col_var.get(),
            },
            'book_mode': {key: var.get() for key, (var, _, _, _) in self.bm_vars.items()}
        }
        return config

    def save_config(self):
        config = self.get_config_as_dict()
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
            messagebox.showinfo("成功", f"設定を {self.config_file} に保存しました。")
        except Exception as e:
            messagebox.showerror("エラー", f"設定の保存に失敗しました:\n{e}")

    def load_config(self):
        if not os.path.exists(self.config_file): return
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)

            # 共通設定
            self.credentials_var.set(config.get('credentials_var', ''))
            self.openai_api_key_var.set(config.get('openai_api_key_var', ''))
            self.spreadsheet_id_var.set(config.get('spreadsheet_id_var', ''))
            self.sheet_name_var.set(config.get('sheet_name_var', '集計'))
            self.start_row_var.set(config.get('start_row_var', '2'))
            self.batch_size_var.set(config.get('batch_size_var', '20'))
            self.translation_delay_var.set(config.get('translation_delay_var', '1.0'))
            self.batch_delay_var.set(config.get('batch_delay_var', '5.0'))
            
            # 通常モード設定
            nm_config = config.get('normal_mode', {})
            self.nm_input_col_var.set(nm_config.get('input_col', 'A'))
            self.nm_translated_name_col_var.set(nm_config.get('translated_name_col', 'K'))
            self.nm_jan_code_col_var.set(nm_config.get('jan_code_col', 'L'))
            self.nm_description_col_var.set(nm_config.get('description_col', 'I'))
            self.nm_output_col_var.set(nm_config.get('output_col', 'Q'))

            # 書籍モード設定
            bm_config = config.get('book_mode', {})
            for key, (var, _, _, _) in self.bm_vars.items():
                var.set(bm_config.get(key, ''))
            
            # 最後にタブを選択
            self.notebook.select(config.get('selected_tab', 0))

            self.log_message(f"📁 設定ファイル {self.config_file} を読み込みました。")
        except Exception as e:
            messagebox.showerror("エラー", f"設定の読み込みに失敗しました:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DescriptionGeneratorGUI(root)
    root.mainloop()
