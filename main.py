"""
日本語テキスト分析ツール
WordCloudと共起ネットワークを生成するGUIアプリケーション

必要なライブラリ:
pip install tkinter pillow wordcloud mecab-python3 networkx matplotlib japanize-matplotlib

※MeCabのインストールも必要です
Windows: https://github.com/ikegami-yukino/mecab/releases
Mac: brew install mecab mecab-ipadic
Linux: sudo apt-get install mecab libmecab-dev mecab-ipadic-utf8
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import MeCab
import ipadic
import re
from functools import lru_cache
from pathlib import Path
from typing import Optional
from collections import Counter
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import networkx as nx
import itertools
from matplotlib import cm
import csv
import io




class JapaneseTextAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("日本語テキスト分析ツール")
        self.root.geometry("1400x900")

        # ============================
        # フォント設定（Meiryo固定）
        # ============================
        self.font_path = r"C:\Windows\Fonts\meiryo.ttc"

        if Path(self.font_path).exists():
            self.font_prop = font_manager.FontProperties(fname=self.font_path)
            plt.rcParams["font.family"] = self.font_prop.get_name()
            plt.rcParams["axes.unicode_minus"] = False
        else:
            messagebox.showwarning(
                "警告",
                f"Meiryo フォントが見つかりませんでした: {self.font_path}\n"
                "Windows環境であることを確認してください。"
            )
            self.font_prop = None

        # MeCab (ipadic 設定を利用)
        try:
            self.mecab = MeCab.Tagger(f"{ipadic.MECAB_ARGS} -Ochasen")
        except Exception:
            messagebox.showerror("警告", "MeCabが見つかりません")
            self.mecab = None


        # データ保持
        self.original_text = ""
        self.tokens = []
        self.word_freq = Counter()
        self.pos_cache = []
        self.original_lines = []  # 【新機能】行情報を保持

        # ストップワード
        self.stop_words = set([
            'の', 'に', 'は', 'を', 'た', 'が', 'で', 'て', 'と', 'し', 'れ', 'さ',
            'ある', 'いる', 'も', 'する', 'から', 'な', 'こと', 'として', 'い',
            'や', 'れる', 'など', 'なっ', 'ない', 'この', 'ため', 'その', 'あっ',
            'よう', 'また', 'もの', 'という', 'あり', 'まで', 'られ', 'なる',
            'へ', 'か', 'だ', 'これ', 'によって', 'により', 'おり', 'より', 'による',
            'ず', 'なり', 'られる', 'において', 'ば', 'なかっ', 'なく', 'しかし',
            'について', 'せ', 'だっ', 'その後', 'できる', 'それ', 'う', 'ので',
            'なお', 'のみ', 'でき', 'き', 'つ', 'における', 'および', 'いう',
            'さらに', 'でも', 'ら', 'たり', 'その他', 'に関する', 'たち', 'ます',
            'ん', 'なら', 'に対して', '特に', 'せる', 'あるいは', 'まし',
            'ながら', 'ただし', 'かつて', 'ください', 'なし', 'これら', 'それら'
        ])

        self.setup_ui()
        self.refresh_stopword_list()

    def setup_ui(self):
        # =============================
        # ttk の日本語フォント設定（Meiryo）
        # =============================
        font_path = r"C:\Windows\Fonts\meiryo.ttc"

        if Path(font_path).exists():
            # 全 ttk ウィジェットに Meiryo を適用
            style = ttk.Style()
            style.configure(".", font=("Meiryo", 11))
            style.configure("TLabel", font=("Meiryo", 11))
            style.configure("TButton", font=("Meiryo", 11))
            style.configure("TEntry", font=("Meiryo", 11))
            style.configure("Treeview", font=("Meiryo", 10))
            style.configure("Treeview.Heading", font=("Meiryo", 10, "bold"))
            style.configure("TNotebook.Tab", font=("Meiryo", 10))

        else:
            messagebox.showwarning(
                "警告",
                f"Meiryo フォントが見つかりません: {font_path}"
            )
            
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # ノートブック（タブ）
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # タブ1: テキスト入力
        self.setup_input_tab()

        # タブ2: 単語編集
        self.setup_edit_tab()

        # タブ3: 可視化
        self.setup_visualize_tab()

        # ウィンドウのリサイズ設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)

    def setup_input_tab(self):
        input_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(input_frame, text="1. テキスト入力")

        # ボタンフレーム
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)

        ttk.Button(btn_frame, text="ファイルから読み込み", command=self.load_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="サンプルテキスト", command=self.load_sample).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="クリア", command=self.clear_text).pack(side=tk.LEFT, padx=5)

        # テキストエリア
        ttk.Label(input_frame, text="解析したいテキストを入力してください:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.text_area = scrolledtext.ScrolledText(input_frame, width=100, height=25, wrap=tk.WORD)
        self.text_area.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        # ストップワード編集（分かち書き前に調整可能）
        stop_frame = ttk.LabelFrame(input_frame, text="ストップワード", padding=5)
        stop_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5)

        stop_list_frame = ttk.Frame(stop_frame); stop_list_frame.pack(fill=tk.BOTH, expand=True)
        stop_scroll = ttk.Scrollbar(stop_list_frame); stop_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.stopword_listbox = tk.Listbox(stop_list_frame, height=6, yscrollcommand=stop_scroll.set)
        self.stopword_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        stop_scroll.config(command=self.stopword_listbox.yview)

        stop_ctrl = ttk.Frame(stop_frame); stop_ctrl.pack(fill=tk.X, pady=4)
        self.stopword_entry = ttk.Entry(stop_ctrl, width=20); self.stopword_entry.pack(side=tk.LEFT, padx=2)
        ttk.Button(stop_ctrl, text="追加", command=self.add_stop_word).pack(side=tk.LEFT, padx=2)
        ttk.Button(stop_ctrl, text="削除", command=self.remove_selected_stop_word).pack(side=tk.LEFT, padx=2)
        ttk.Button(stop_ctrl, text="適用", command=self.apply_stop_words).pack(side=tk.LEFT, padx=2)

        # 分かち書きボタン
        ttk.Button(input_frame, text="分かち書き実行", command=self.tokenize_text,
                   style="Accent.TButton").grid(row=4, column=0, pady=10)

        input_frame.columnconfigure(0, weight=1)
        input_frame.rowconfigure(2, weight=1)

    def setup_edit_tab(self):
        edit_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(edit_frame, text="2. 単語編集")

        # 左右分割
        left_frame = ttk.Frame(edit_frame)
        left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)

        right_frame = ttk.Frame(edit_frame)
        right_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)

        # 左側: 単語リスト
        ttk.Label(left_frame, text="単語頻度リスト", font=("Meiryo", 12, "bold")).pack(pady=5)

        # 検索フレーム
        search_frame = ttk.Frame(left_frame)
        search_frame.pack(fill=tk.X, pady=5)
        ttk.Label(search_frame, text="検索:").pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.filter_word_list)
        ttk.Entry(search_frame, textvariable=self.search_var, width=20).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 単語リスト
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.word_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=20)
        self.word_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.word_listbox.yview)

        # ボタン
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame, text="選択した単語を削除", command=self.delete_selected_word).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="すべてリフレッシュ", command=self.refresh_word_list).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="品詞で削除", command=self.delete_by_pos).pack(side=tk.LEFT, padx=5)

        # 右側: 編集エリア
        ttk.Label(right_frame, text="単語編集（スペース区切り）", font=("", 12, "bold")).pack(pady=5)

        edit_control_frame = ttk.Frame(right_frame)
        edit_control_frame.pack(fill=tk.X, pady=5)

        ttk.Label(edit_control_frame, text="置換:").pack(side=tk.LEFT, padx=5)
        self.replace_from = ttk.Entry(edit_control_frame, width=15)
        self.replace_from.pack(side=tk.LEFT, padx=2)
        ttk.Label(edit_control_frame, text="→").pack(side=tk.LEFT, padx=2)
        self.replace_to = ttk.Entry(edit_control_frame, width=15)
        self.replace_to.pack(side=tk.LEFT, padx=2)
        ttk.Button(edit_control_frame, text="置換", command=self.replace_word).pack(side=tk.LEFT, padx=5)

        self.edit_area = scrolledtext.ScrolledText(right_frame, width=60, height=20, wrap=tk.WORD)
        self.edit_area.pack(fill=tk.BOTH, expand=True, pady=5)

        # パラメータと実行ボタン
        param_frame = ttk.Frame(right_frame)
        param_frame.pack(fill=tk.X, pady=5)

        # 左側にパラメータ、右側にアクションを分けて配置
        param_grid = ttk.Frame(param_frame)
        param_grid.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        ttk.Label(param_grid, text="最小出現回数:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.min_freq_var = tk.IntVar(value=2)
        ttk.Spinbox(param_grid, from_=1, to=20, textvariable=self.min_freq_var, width=7).grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)

        ttk.Label(param_grid, text="共起ウィンドウ:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.window_var = tk.IntVar(value=5)
        ttk.Spinbox(param_grid, from_=2, to=20, textvariable=self.window_var, width=7).grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)

        ttk.Label(param_grid, text="ノード色パレット:").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        self.network_cmap_var = tk.StringVar(value="Pastel1")
        ttk.Combobox(
            param_grid,
            values=["Pastel1", "Pastel2", "Set3", "Accent", "tab20"],
            textvariable=self.network_cmap_var,
            state="readonly",
            width=12
        ).grid(row=2, column=1, padx=5, pady=2, sticky=tk.W)

        ttk.Label(param_frame, text="ノード色パレット:").pack(side=tk.LEFT, padx=5)
        self.network_cmap_var = tk.StringVar(value="Pastel1")
        ttk.Combobox(
            param_frame,
            values=["Pastel1", "Pastel2", "Set3", "Accent", "tab20"],
            textvariable=self.network_cmap_var,
            state="readonly",
            width=10
        ).pack(side=tk.LEFT, padx=2)

        # 追加: WordCloud 画像サイズ設定（タブで事前指定）
        ttk.Label(param_grid, text="WordCloud 幅:").grid(row=3, column=0, padx=5, pady=2, sticky=tk.W)
        self.wc_width_var = tk.IntVar(value=1000)
        ttk.Spinbox(param_grid, from_=100, to=5000, textvariable=self.wc_width_var, width=7).grid(row=3, column=1, padx=5, pady=2, sticky=tk.W)

        ttk.Label(param_grid, text="WordCloud 高さ:").grid(row=3, column=2, padx=5, pady=2, sticky=tk.W)
        self.wc_height_var = tk.IntVar(value=600)
        ttk.Spinbox(param_grid, from_=100, to=5000, textvariable=self.wc_height_var, width=7).grid(row=3, column=3, padx=5, pady=2, sticky=tk.W)

        # 共起ネットワーク出力サイズ
        ttk.Label(param_grid, text="ネットワーク幅:").grid(row=4, column=0, padx=5, pady=2, sticky=tk.W)
        self.net_width_var = tk.IntVar(value=1200)
        ttk.Spinbox(param_grid, from_=200, to=5000, textvariable=self.net_width_var, width=7).grid(row=4, column=1, padx=5, pady=2, sticky=tk.W)

        ttk.Label(param_grid, text="ネットワーク高さ:").grid(row=4, column=2, padx=5, pady=2, sticky=tk.W)
        self.net_height_var = tk.IntVar(value=800)
        ttk.Spinbox(param_grid, from_=200, to=5000, textvariable=self.net_height_var, width=7).grid(row=4, column=3, padx=5, pady=2, sticky=tk.W)

        # 追加: 共起ネットワーク表示組数
        ttk.Label(param_grid, text="ネットワーク表示組数:").grid(row=5, column=0, padx=5, pady=2, sticky=tk.W)
        self.net_edge_count_var = tk.IntVar(value=50)
        ttk.Spinbox(param_grid, from_=10, to=500, textvariable=self.net_edge_count_var, width=7).grid(row=5, column=1, padx=5, pady=2, sticky=tk.W)

        # 【新機能】自己回帰ネットワーク制御
        ttk.Label(param_grid, text="自己ループ:").grid(row=6, column=0, padx=5, pady=2, sticky=tk.W)
        self.self_loop_var = tk.StringVar(value="remove")
        loop_frame = ttk.Frame(param_grid)
        loop_frame.grid(row=6, column=1, padx=5, pady=2, sticky=tk.W)
        ttk.Radiobutton(loop_frame, text="削除", variable=self.self_loop_var, value="remove").pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(loop_frame, text="描画", variable=self.self_loop_var, value="keep").pack(side=tk.LEFT, padx=2)

        # 【新機能】共起ウィンドウ形式
        ttk.Label(param_grid, text="共起ウィンドウ形式:").grid(row=7, column=0, padx=5, pady=2, sticky=tk.W)
        self.window_mode_var = tk.StringVar(value="sliding")
        mode_frame = ttk.Frame(param_grid)
        mode_frame.grid(row=7, column=1, columnspan=3, padx=5, pady=2, sticky=tk.W)
        ttk.Radiobutton(mode_frame, text="スライディングウィンドウ", variable=self.window_mode_var, value="sliding").pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(mode_frame, text="行ごと", variable=self.window_mode_var, value="line").pack(side=tk.LEFT, padx=2)

        # グラフ生成ボタンを縦方向に配置して見切れを防止
        action_frame = ttk.Frame(param_frame)
        action_frame.pack(side=tk.RIGHT, padx=5)
        ttk.Button(action_frame, text="WordCloud生成", command=self.on_generate_wordcloud).grid(row=0, column=0, pady=2, sticky=tk.EW)
        ttk.Button(action_frame, text="共起ネットワーク生成", command=self.on_generate_network).grid(row=1, column=0, pady=2, sticky=tk.EW)
        ttk.Button(action_frame, text="頻度グラフ生成", command=self.on_generate_frequency_chart).grid(row=2, column=0, pady=2, sticky=tk.EW)

        edit_frame.columnconfigure(0, weight=1)
        edit_frame.columnconfigure(1, weight=2)
        edit_frame.rowconfigure(0, weight=1)

    def setup_visualize_tab(self):
        vis_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(vis_frame, text="3. 可視化")

        # サブタブ
        self.vis_notebook = ttk.Notebook(vis_frame)
        self.vis_notebook.pack(fill=tk.BOTH, expand=True)

        # WordCloud
        self.wordcloud_frame = ttk.Frame(self.vis_notebook)
        self.vis_notebook.add(self.wordcloud_frame, text="WordCloud")

        # 共起ネットワーク
        self.network_frame = ttk.Frame(self.vis_notebook)
        self.vis_notebook.add(self.network_frame, text="共起ネットワーク")

        # 頻度グラフ
        self.freq_frame = ttk.Frame(self.vis_notebook)
        self.vis_notebook.add(self.freq_frame, text="頻度グラフ")

    def load_file(self):
        filepath = filedialog.askopenfilename(
            filetypes=[
                ("テキストファイル", "*.txt"),
                ("CSVファイル", "*.csv"),
                ("すべてのファイル", "*.*")
            ]
        )
        if filepath:
            try:
                # 拡張子判定は小文字で比較
                if filepath.lower().endswith('.csv'):
                    self.load_csv_file(filepath)
                else:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        content = f.read()
                        self.text_area.delete(1.0, tk.END)
                        self.text_area.insert(1.0, content)
                        # 行情報を保持
                        self.original_text = content
                        self.original_lines = content.split('\n')
            except Exception as e:
                messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました: {e}")

    def load_csv_file(self, filepath):
        """CSVファイルを読み込み、指定列のテキストを結合（エンコーディング/区切り検出付き）"""
        try:
            # バイナリ読み取りしてエンコーディング候補でデコードを試みる
            with open(filepath, 'rb') as bf:
                raw = bf.read()

            enc_candidates = [
                'utf-8-sig', 'utf-8', 'cp932', 'shift_jis', 'euc_jp',
                'utf-16', 'utf-16-le', 'utf-16-be'
            ]
            decoded = None
            used_enc = None
            for enc in enc_candidates:
                try:
                    decoded = raw.decode(enc)
                    used_enc = enc
                    break
                except Exception:
                    continue

            if decoded is None:
                # デコードできない場合でも中身を確認できるように置換デコードでフォールバック
                decoded = raw.decode('utf-8', errors='replace')
                used_enc = 'utf-8 (replace)'

            # 改行を統一（Sniffer が失敗しやすいケースを減らす）
            decoded = decoded.replace('\r\n', '\n').replace('\r', '\n')

            # 区切り文字を推定（先頭最大4KBをサンプルにして推定）
            sample = decoded[:4096]
            delimiter = ','
            dialect = None
            try:
                dialect = csv.Sniffer().sniff(sample)
                delimiter = dialect.delimiter
            except Exception:
                # フォールバック候補（カンマ、タブ、セミコロン）
                for cand in [',', '\t', ';']:
                    try:
                        # 簡易チェック：cand が行内に含まれるか
                        if cand in sample:
                            delimiter = cand
                            break
                    except Exception:
                        continue

            # CSVを読み込み（io.StringIO 経由）
            if dialect:
                reader = csv.reader(io.StringIO(decoded), dialect=dialect)
            else:
                reader = csv.reader(io.StringIO(decoded), delimiter=delimiter)
            rows = list(reader)
            if not rows:
                messagebox.showwarning("警告", "CSVファイルが空です。")
                return

            # 列選択ダイアログ
            col_window = tk.Toplevel(self.root)
            col_window.title("CSV列選択")
            col_window.geometry("420x360")

            ttk.Label(col_window, text=f"検出エンコーディング: {used_enc}   推定区切り文字: '{delimiter}'", wraplength=380).pack(pady=6, padx=10)
            ttk.Label(col_window, text="結合する列を選択してください（複数選択可）:", wraplength=380).pack(pady=6, padx=10)

            # ヘッダー行の判定（ユーザに確認）
            sniff_header = False
            try:
                sniff_header = csv.Sniffer().has_header(sample)
            except Exception:
                sniff_header = False

            header_prompt = "最初の行はヘッダーですか？"
            if sniff_header:
                header_prompt += " (推定: ヘッダーあり)"

            has_header = messagebox.askyesno("CSVヘッダ", header_prompt)
            header_row = rows[0] if has_header else [f"列{i+1}" for i in range(len(rows[0]))]
            data_start = 1 if has_header else 0

            # チェックボックスリスト（スクロール対応）
            check_frame = ttk.Frame(col_window)
            check_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

            scrollbar = ttk.Scrollbar(check_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            canvas = tk.Canvas(check_frame, yscrollcommand=scrollbar.set, height=220)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=canvas.yview)

            inner_frame = ttk.Frame(canvas)
            canvas_window = canvas.create_window((0, 0), window=inner_frame, anchor=tk.NW)

            col_vars = []
            for i, col_name in enumerate(header_row):
                var = tk.BooleanVar(value=False)
                col_vars.append(var)
                ttk.Checkbutton(
                    inner_frame,
                    text=f"列{i+1}: {str(col_name)[:60]}",
                    variable=var
                ).pack(anchor=tk.W, pady=2)

            def _on_config(event):
                # canvas の幅に合わせてウィンドウ幅を更新（スクロール横幅問題の解消）
                canvas.itemconfigure(canvas_window, width=event.width)
                canvas.config(scrollregion=canvas.bbox("all"))
            inner_frame.bind("<Configure>", _on_config)
            canvas.bind("<Configure>", lambda e: canvas.itemconfigure(canvas_window, width=e.width))

            def apply_selection():
                selected_indices = [i for i, var in enumerate(col_vars) if var.get()]
                if not selected_indices:
                    messagebox.showwarning("警告", "最低1つの列を選択してください。")
                    return

                # CSV行ごとに選択列を結合（空セルは空文字扱い）
                text_lines = []
                for row in rows[data_start:]:
                    # 安全にインデックス参照し、空白とBOM除去
                    parts = []
                    for i in selected_indices:
                        val = row[i] if i < len(row) else ""
                        if isinstance(val, str):
                            parts.append(val.strip())
                        else:
                            parts.append(str(val).strip())
                    row_text = " ".join([p for p in parts if p])
                    if row_text:  # 空行は無視
                        text_lines.append(row_text)

                combined_text = "\n".join(text_lines).strip()
                if not combined_text:
                    messagebox.showwarning("警告", "選択列の結合結果が空でした。別の列を選択してください。")
                    return

                # テキストエリアに確実に挿入
                self.text_area.delete(1.0, tk.END)
                self.text_area.insert(1.0, combined_text)
                # 行情報と元テキストを保持
                self.original_text = combined_text
                self.original_lines = text_lines

                col_window.destroy()
                messagebox.showinfo("完了", f"{len(selected_indices)}列を結合しました。")

            ttk.Button(col_window, text="適用", command=apply_selection).pack(pady=8)
            ttk.Button(col_window, text="キャンセル", command=col_window.destroy).pack(pady=2)
        except Exception as e:
            messagebox.showerror("エラー", f"CSVファイルの読み込みに失敗しました: {e}")

    def load_sample(self):
        sample = """人工知能は現代社会において重要な技術となっています。機械学習やディープラーニングの発展により、
画像認識や自然言語処理などの分野で大きな進歩がありました。これらの技術は医療診断、自動運転、
音声認識など様々な応用分野で活用されています。今後も人工知能技術の発展により、
社会の様々な課題解決に貢献することが期待されています。データ分析の重要性も高まっており、
ビッグデータを活用した意思決定が多くの企業で行われています。テクノロジーの進化は
私たちの生活を大きく変えつつあります。人工知能の発展は目覚ましく、機械学習アルゴリズムの
改善により精度が向上しています。自然言語処理技術も進歩し、より自然な対話が可能になりました。"""
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(1.0, sample)
        # 【改善】サンプル用に行情報を初期化
        self.original_lines = sample.split('\n')

    def clear_text(self):
        self.text_area.delete(1.0, tk.END)

    def parse_with_pos(self, text: str):
        tagger = MeCab.Tagger(f"{ipadic.MECAB_ARGS} -Ochasen")
        parsed = tagger.parse(text)
        lines = [l for l in parsed.splitlines() if l and l != "EOS"]
        surfaces, pos_list = [], []
        for line in lines:
            parts = line.split("\t")
            if len(parts) >= 4:
                surfaces.append(parts[0])
                pos_list.append(parts[3].split("-")[0])
        return surfaces, pos_list
    
    def tokenize_text(self):
        if not self.mecab:
            messagebox.showerror("エラー", "MeCabが初期化されていません。")
            return

        text = self.text_area.get(1.0, tk.END).strip()
        if not text:
            messagebox.showwarning("警告", "テキストを入力してください。")
            return

        self.original_text = text

        # 【改善】行情報をトークン化済みで保持（行ごとの共起計算用）
        self.original_lines = []
        for raw_line in text.split('\n'):
            line_surfaces, _ = self.parse_with_pos(raw_line)
            line_tokens = [s for s in line_surfaces if s not in self.stop_words and len(s) > 1]
            if line_tokens:
                self.original_lines.append(" ".join(line_tokens))

        surfaces, pos_list = self.parse_with_pos(text)
        if not surfaces:
            messagebox.showerror("エラー", "MeCabの解析結果を取得できませんでした。")
            return

        self.tokens = [s for s in surfaces if s not in self.stop_words and len(s) > 1]
        # POS cache aligned with filtered tokens
        self.pos_cache = [p for s, p in zip(surfaces, pos_list) if s in self.tokens]

        self.word_freq = Counter(self.tokens)
        self.edit_area.delete(1.0, tk.END)
        self.edit_area.insert(1.0, " ".join(self.tokens))
        self.refresh_word_list()

        # タブ切り替え
        self.notebook.select(1)
        messagebox.showinfo("完了", f"{len(self.tokens)}個の単語を抽出しました。")

    def refresh_word_list(self):
        text = self.edit_area.get(1.0, tk.END).strip()
        self.tokens = text.split()
        self.word_freq = Counter(self.tokens)
        self.pos_cache = [self.get_pos(t) for t in self.tokens]  # get_pos below uses cached Ochasen tagging

        # 【改善】編集内容を行単位のトークン列として保持し、共起ネットワークに反映
        self.original_lines = [" ".join(line.split()) for line in text.split('\n') if line.split()]

        # リスト更新
        self.word_listbox.delete(0, tk.END)
        for word, count in self.word_freq.most_common():
            self.word_listbox.insert(tk.END, f"{word} ({count}回)")

        # ストップワード表示も更新
        self.refresh_stopword_list()

    def refresh_stopword_list(self):
        if not hasattr(self, "stopword_listbox"):
            return
        self.stopword_listbox.delete(0, tk.END)
        for w in sorted(self.stop_words):
            self.stopword_listbox.insert(tk.END, w)

    def add_stop_word(self):
        word = self.stopword_entry.get().strip()
        if not word:
            return
        self.stop_words.add(word)
        self.stopword_entry.delete(0, tk.END)
        self.refresh_stopword_list()
        self.apply_stop_words()

    def remove_selected_stop_word(self):
        selection = self.stopword_listbox.curselection()
        if not selection:
            return
        word = self.stopword_listbox.get(selection[0])
        self.stop_words.discard(word)
        self.refresh_stopword_list()
        self.apply_stop_words()

    def apply_stop_words(self):
        # Listbox をソースとして self.stop_words を同期
        if hasattr(self, "stopword_listbox"):
            self.stop_words = set(self.stopword_listbox.get(0, tk.END))

        text = self.edit_area.get(1.0, tk.END).strip()
        if not text:
            return
        words = text.split()
        filtered = [w for w in words if w not in self.stop_words]
        self.edit_area.delete(1.0, tk.END)
        self.edit_area.insert(1.0, " ".join(filtered))
        self.refresh_word_list()

    def filter_word_list(self, *args):
        search_term = self.search_var.get()
        self.word_listbox.delete(0, tk.END)

        for word, count in self.word_freq.most_common():
            if search_term.lower() in word.lower():
                self.word_listbox.insert(tk.END, f"{word} ({count}回)")

    @lru_cache(maxsize=4096)
    def get_pos(self, word: str) -> str:
        if not word or not self.mecab:
            return ""
        tagger = MeCab.Tagger(f"{ipadic.MECAB_ARGS} -Ochasen")
        parsed = tagger.parse(word)
        if not parsed:
            return ""
        first = parsed.splitlines()[0]
        parts = first.split("\t")
        if len(parts) < 4:
            return ""
        pos_field = parts[3]
        return pos_field.split("-")[0] if pos_field else ""


    def find_font_path(self) -> Optional[str]:
        return "C:/Windows/Fonts/meiryo.ttc"


    def delete_by_pos(self):
        if not self.tokens:
            return

        pos_window = tk.Toplevel(self.root)
        pos_window.title("品詞で削除")
        pos_window.geometry("360x220")

        ttk.Label(pos_window, text="削除する品詞を選択してください").pack(pady=8)

        pos_options = [
            "名詞", "動詞", "形容詞", "副詞", "形容動詞", "助詞", "助動詞", "記号"
        ]
        # 現在存在する品詞を表示
        current_pos_counts = Counter(self.pos_cache)
        if current_pos_counts:
            info_lines = [f"{p}: {current_pos_counts[p]}件" for p in pos_options if p in current_pos_counts]
            extra = [f"{p}: {n}件" for p, n in current_pos_counts.items() if p not in pos_options]
            info_lines.extend(extra)
            ttk.Label(pos_window, text="現在の品詞内訳").pack(pady=(4, 2))
            info_text = "\n".join(info_lines) if info_lines else "なし"
            ttk.Label(pos_window, text=info_text, justify=tk.LEFT).pack()

        selected_pos = tk.StringVar(value=pos_options[0])
        ttk.Combobox(pos_window, values=pos_options, textvariable=selected_pos, state="readonly").pack(pady=8)

        def perform_delete():
            target = selected_pos.get()
            filtered_tokens = [t for t, p in zip(self.tokens, self.pos_cache) if p != target]
            self.edit_area.delete(1.0, tk.END)
            self.edit_area.insert(1.0, " ".join(filtered_tokens))
            self.refresh_word_list()
            pos_window.destroy()

        ttk.Button(pos_window, text="削除", command=perform_delete).pack(pady=10)

    def delete_selected_word(self):
        selection = self.word_listbox.curselection()
        if not selection:
            return

        item = self.word_listbox.get(selection[0])
        word = item.split(' (')[0]

        # 編集エリアから削除
        text = self.edit_area.get(1.0, tk.END)
        words = text.split()
        words = [w for w in words if w != word]

        self.edit_area.delete(1.0, tk.END)
        self.edit_area.insert(1.0, ' '.join(words))

        self.refresh_word_list()

    def replace_word(self):
        from_word = self.replace_from.get()
        to_word = self.replace_to.get()

        if not from_word:
            return

        text = self.edit_area.get(1.0, tk.END)
        words = text.split()
        words = [to_word if w == from_word else w for w in words]

        self.edit_area.delete(1.0, tk.END)
        self.edit_area.insert(1.0, ' '.join(words))

        self.refresh_word_list()
        self.replace_from.delete(0, tk.END)
        self.replace_to.delete(0, tk.END)

    def visualize(self):
        # 編集された単語を取得
        text = self.edit_area.get(1.0, tk.END).strip()
        if not text:
            messagebox.showwarning("警告", "単語データがありません。")
            return

        tokens = text.split()
        word_freq = Counter(tokens)

        # 最小出現回数でフィルタリング
        min_freq = self.min_freq_var.get()
        filtered_freq = {k: v for k, v in word_freq.items() if v >= min_freq}

        if not filtered_freq:
            messagebox.showwarning("警告", f"最小出現回数{min_freq}回以上の単語がありません。")
            return

        try:
            # WordCloud生成
            self.generate_wordcloud(filtered_freq)

            # 共起ネットワーク生成
            self.generate_network(tokens, filtered_freq)

            # 頻度グラフ生成
            self.generate_frequency_chart(filtered_freq)
        except Exception as e:
            messagebox.showerror("エラー", f"可視化の生成中に問題が発生しました: {e}")
            return

        # タブ切り替え
        self.notebook.select(2)
        messagebox.showinfo("完了", "可視化が完了しました。")

    def generate_wordcloud(self, word_freq):
        # 既存のウィジェットをクリア
        for widget in self.wordcloud_frame.winfo_children():
            widget.destroy()

        # タブで指定したサイズを使用する（デフォルト値は Spinbox にて設定）
        width = getattr(self, "wc_width_var", tk.IntVar(value=1000)).get()
        height = getattr(self, "wc_height_var", tk.IntVar(value=600)).get()

        # WordCloud生成
        wc = WordCloud(
            width=width,
            height=height,
            background_color='white',
            font_path=self.font_path,
            relative_scaling=0.5,
            min_font_size=10,
            max_font_size=100,
            colormap='tab10'
        ).generate_from_frequencies(word_freq)

        # 描画（Figure サイズは表示用に固定、画像保存は save_figure で行う）
        fig, ax = plt.subplots(figsize=(12, 7))
        ax.imshow(wc, interpolation='bilinear')
        ax.axis('off')
        ax.set_title('WordCloud', fontsize=16, pad=20)

        canvas = FigureCanvasTkAgg(fig, self.wordcloud_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # 保存ボタン
        ttk.Button(self.wordcloud_frame, text="画像として保存",
                   command=lambda: self.save_figure(fig, "wordcloud")).pack(pady=5)

        if not self.font_path:
            ttk.Label(self.wordcloud_frame, text="※日本語フォントが見つからないため、文字化けする可能性があります。", foreground="red").pack(pady=5)

    def generate_network(self, tokens, word_freq):
        # 既存のウィジェットをクリア
        for widget in self.network_frame.winfo_children():
            widget.destroy()

        window_size = self.window_var.get()
        edge_count = getattr(self, "net_edge_count_var", tk.IntVar(value=50)).get()
        self_loop_mode = getattr(self, "self_loop_var", tk.StringVar(value="remove")).get()
        window_mode = getattr(self, "window_mode_var", tk.StringVar(value="sliding")).get()

        # 共起ペア抽出
        cooc_pairs = []
        
        if window_mode == "sliding":
            # スライディングウィンドウ形式（従来通り）
            for i in range(len(tokens)):
                if tokens[i] not in word_freq:
                    continue
                for j in range(i + 1, min(i + window_size, len(tokens))):
                    if tokens[j] not in word_freq:
                        continue
                    pair = tuple(sorted([tokens[i], tokens[j]]))
                    cooc_pairs.append(pair)
        else:
            # 【改善】行ごと形式：保持した行情報を使用
            for line in self.original_lines:
                if not line.strip():  # 空行スキップ
                    continue
                # 行を分割しトークン化（既に分かち書きされている場合）
                line_tokens = line.split()
                for i in range(len(line_tokens)):
                    if line_tokens[i] not in word_freq:
                        continue
                    for j in range(i + 1, len(line_tokens)):
                        if line_tokens[j] not in word_freq:
                            continue
                        pair = tuple(sorted([line_tokens[i], line_tokens[j]]))
                        cooc_pairs.append(pair)

        cooc_count = Counter(cooc_pairs)

        if not cooc_count:
            ttk.Label(self.network_frame, text="共起データがありません").pack(pady=20)
            return

        # ネットワーク構築（指定した組数を使用）
        G = nx.Graph()
        for (word1, word2), count in cooc_count.most_common(edge_count):
            # 【新機能】自己ループ（同じ単語同士）の制御
            if word1 == word2 and self_loop_mode == "remove":
                continue
            G.add_edge(word1, word2, weight=count)

        if len(G.nodes()) == 0:
            ttk.Label(self.network_frame, text="表示できるネットワークがありません").pack(pady=20)
            return

        # 【改善】最大連結成分のみを抽出
        if not nx.is_connected(G):
            largest_cc = max(nx.connected_components(G), key=len)
            G = G.subgraph(largest_cc).copy()

        # 【改善】最小重みしきい値でエッジをフィルタリング（孤立防止）
        if len(G.edges()) > 0:
            max_weight_val = max([d['weight'] for u, v, d in G.edges(data=True)])
            min_weight = max(1, max_weight_val // 5)
            G = nx.Graph([(u, v, d) for u, v, d in G.edges(data=True) if d['weight'] >= min_weight])

        if len(G.nodes()) < 2:
            ttk.Label(self.network_frame, text="十分な共起ネットワークが見つかりません。\nウィンドウサイズを大きくするか、共起組数を増やしてください。").pack(pady=20)
            return

        # 描画サイズを取得
        w = getattr(self, "net_width_var", None)
        h = getattr(self, "net_height_var", None)
        fig_w = w.get() / 100 if w else 12
        fig_h = h.get() / 100 if h else 8
        fig, ax = plt.subplots(figsize=(fig_w, fig_h), facecolor='white')
        ax.set_facecolor('white')

        # コミュニティ検出でノードを色分け
        communities = list(nx.community.greedy_modularity_communities(G))
        comm_map = {}
        for idx, nodes in enumerate(communities):
            for n in nodes:
                comm_map[n] = idx

        # 【改善】レイアウト計算：Kamada-Kawaiアルゴリズムで視認性向上
        try:
            pos = nx.kamada_kawai_layout(G, scale=2)
        except:
            # Kamada-Kawaiが失敗した場合はspring_layoutで代替
            layout_k = 2 / (len(G.nodes()) ** 0.5)
            pos = nx.spring_layout(
                G,
                k=layout_k,
                iterations=500,
                seed=42,
                scale=2,
                weight="weight"
            )

        # ノードサイズを単語の出現頻度に基づいて計算
        node_sizes = [max(300, word_freq.get(node, 1) * 150) for node in G.nodes()]

        # エッジの重みを正規化して透明度と幅に反映
        edges = G.edges()
        weights = [G[u][v]['weight'] for u, v in edges]
        max_weight = max(weights) if weights else 1
        normalized_weights = [w / max_weight for w in weights]
        edge_widths = [1.5 + normalized_w * 6 for normalized_w in normalized_weights]
        edge_alphas = [0.3 + normalized_w * 0.5 for normalized_w in normalized_weights]

        # ノードの色をコミュニティに基づいて設定
        cmap_name = getattr(self, "network_cmap_var", tk.StringVar(value="Pastel1")).get()
        cmap = cm.get_cmap(cmap_name)
        num_colors = getattr(cmap, "N", 20)
        colors = [
            cmap((comm_map.get(n, 0) % num_colors) / max(num_colors - 1, 1))
            for n in G.nodes()
        ]

        # ノード描画
        nx.draw_networkx_nodes(
            G, pos,
            node_size=node_sizes,
            node_color=colors,
            alpha=0.85,
            ax=ax,
            linewidths=2,
            edgecolors='#222'
        )

        # エッジ描画（重みに応じた透明度）
        for (u, v), width, alpha in zip(edges, edge_widths, edge_alphas):
            nx.draw_networkx_edges(
                G, pos,
                edgelist=[(u, v)],
                width=width,
                alpha=alpha,
                edge_color='#666',
                ax=ax
            )

        # ラベル描画
        font_family = self.font_prop.get_name() if getattr(self, "font_prop", None) else "sans-serif"
        nx.draw_networkx_labels(
            G, pos,
            font_size=9,
            font_family=font_family,
            font_weight='bold',
            ax=ax
        )

        ax.set_title(f"共起ネットワーク（{len(G.nodes())} ノード、{len(G.edges())} エッジ）\n{window_mode}形式、{self_loop_mode}", fontsize=16, pad=20, weight='bold')
        ax.axis("off")
        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, self.network_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # 保存ボタン
        ttk.Button(self.network_frame, text="画像として保存",
                   command=lambda: self.save_figure(fig, "network")).pack(pady=5)
        ttk.Button(self.network_frame, text="SVGで保存",
                   command=lambda: self.save_figure(fig, "network", fmt="svg")).pack(pady=5)

    def generate_frequency_chart(self, word_freq):
        # 既存のウィジェットをクリア
        for widget in self.freq_frame.winfo_children():
            widget.destroy()

        # 上位30単語
        top_words = dict(sorted(word_freq.items(), key=lambda x: x[1], reverse=True)[:30])

        # 描画
        fig, ax = plt.subplots(figsize=(12, 8))
        words = list(top_words.keys())
        counts = list(top_words.values())

        ax.barh(words, counts, color='steelblue')
        ax.set_xlabel('出現回数', fontsize=12)
        ax.set_title(f'単語出現頻度（全{len(word_freq)}単語中の上位30単語）', fontsize=16, pad=20)
        ax.invert_yaxis()
        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, self.freq_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # 保存ボタン
        ttk.Button(self.freq_frame, text="画像として保存",
                   command=lambda: self.save_figure(fig, "frequency")).pack(pady=5)

    def save_figure(self, fig, name, fmt=None):
        filepath = filedialog.asksaveasfilename(
            defaultextension=".png",
            initialfile=f"{name}.png",
            filetypes=[("PNG", "*.png"), ("PDF", "*.pdf"), ("SVG", "*.svg")]
        )
        if filepath:
            save_kwargs = {"dpi": 300, "bbox_inches": "tight"}
            if fmt:
                save_kwargs["format"] = fmt
            fig.savefig(filepath, **save_kwargs)
            messagebox.showinfo("完了", f"保存しました: {filepath}")

    # ---------- 追加: 編集タブから呼び出すラッパー関数 ----------
    def on_generate_wordcloud(self):
        # 編集エリアから単語・頻度を取得し、最小出現回数でフィルタ
        text = self.edit_area.get(1.0, tk.END).strip()
        if not text:
            messagebox.showwarning("警告", "単語データがありません。")
            return
        tokens = text.split()
        word_freq = Counter(tokens)
        min_freq = self.min_freq_var.get()
        filtered_freq = {k: v for k, v in word_freq.items() if v >= min_freq}
        if not filtered_freq:
            messagebox.showwarning("警告", f"最小出現回数{min_freq}回以上の単語がありません。")
            return
        try:
            self.generate_wordcloud(filtered_freq)
            self.notebook.select(2)  # 可視化タブへ
        except Exception as e:
            messagebox.showerror("エラー", f"WordCloud の生成中に問題が発生しました: {e}")

    def on_generate_network(self):
        text = self.edit_area.get(1.0, tk.END).strip()
        if not text:
            messagebox.showwarning("警告", "単語データがありません。")
            return
        tokens = text.split()
        word_freq = Counter(tokens)
        min_freq = self.min_freq_var.get()
        filtered_freq = {k: v for k, v in word_freq.items() if v >= min_freq}
        if not filtered_freq:
            messagebox.showwarning("警告", f"最小出現回数{min_freq}回以上の単語がありません。")
            return
        try:
            self.generate_network(tokens, filtered_freq)
            self.notebook.select(2)
        except Exception as e:
            messagebox.showerror("エラー", f"共起ネットワークの生成中に問題が発生しました: {e}")

    def on_generate_frequency_chart(self):
        text = self.edit_area.get(1.0, tk.END).strip()
        if not text:
            messagebox.showwarning("警告", "単語データがありません。")
            return
        tokens = text.split()
        word_freq = Counter(tokens)
        min_freq = self.min_freq_var.get()
        filtered_freq = {k: v for k, v in word_freq.items() if v >= min_freq}
        if not filtered_freq:
            messagebox.showwarning("警告", f"最小出現回数{min_freq}回以上の単語がありません。")
            return
        try:
            self.generate_frequency_chart(filtered_freq)
            self.notebook.select(2)
        except Exception as e:
            messagebox.showerror("エラー", f"頻度グラフの生成中に問題が発生しました: {e}")

def main():
    root = tk.Tk()
    app = JapaneseTextAnalyzer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
