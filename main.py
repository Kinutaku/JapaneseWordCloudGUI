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
from PIL import Image
import numpy as np




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

        # --- 追加: 分かち書き（ストップワード除去前）行情報と連語ルール ---
        self.pre_tokens_lines = []          # 各行ごとの MeCab 分かち書き（ストップワード除去前）
        self.merge_rules = []               # ルールリスト: {"len":n, "seq":tuple(...), "merged": "結合語"}

        # ストップワード
        self.stop_words = set([
            '（','）','(',')','［','］','[',']','{','}','【','】','※','→','⇒','…','‥','…','—','〜','%','!','?','！？','?!',
            'へと','よりも','つつ','ながらも','だろ','だろう','でしょう','です','でした','ますが','ません','ませんでした','んで','のでしょう','のでした',
            'ところ','ところが','ところで','ために','ための','ためには','わけ','わけで','わけでは','はず','はずが','はずだ','ものの','ものと','ことが','ことに','ことから','それぞれ','それぞれの','ように','ような','ようで',
            'こんな','そんな','あんな','どの','どれ','どう','どういった','ここ','そこ','あそこ','どこ','こちら','そちら','あちら',
            'まず','次に','そして','一方','ただ','だが','その結果','結果として','つまり','要するに',
            '的','的な','的に','等','等の','等について','化','性',
            '0','1','2','3','4','5','6','7','8','9',
            '０','１','２','３','４','５','６','７','８','９',
            '年','月','日','時','分','％',
            'の', 'に', 'は', 'を', 'た', 'が', 'で', 'て', 'と', 'し', 
            'れ', 'さ', 'ある', 'いる', 'も', 'する', 'から', 'な', 'こと', 
            'として', 'い', 'や', 'れる', 'など', 'なっ', 'ない', 'この', 'ため', 
            'その', 'あっ', 'よう', 'また', 'もの', 'という', 'あり', 'まで', 'られ', 
            'なる', 'へ', 'か', 'だ', 'これ', 'によって', 'により', 'おり', 'より', 
            'による', 'ず', 'なり', 'られる', 'において', 'ば', 'なかっ', 'なく', 
            'しかし', 'について', 'せ', 'だっ', 'その後', 'できる', 'それ', 
            'う', 'ので', 'なお', 'のみ', 'でき', 'き', 'つ', 'における', 
            'および', 'いう', 'さらに', 'でも', 'ら', 'たり', 'その他', 
            'に関する', 'たち', 'ます', 'ん', 'なら', 'に対して', '特に', 
            'せる', 'あるいは', 'まし', 'ながら', 'ただし', 'かつて', 
            'ください', 'なし', 'これら', 'それら',"、","。","・",
            "「","」","『","』","〈","〉","《","》","．","，","：","；","！","？"

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

        # --- 追加: 連語結合タブ ---
        self.setup_merge_tab()

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

    def setup_merge_tab(self):
        """新規タブ: 連語（2/3/4語）を結合して1語として扱うルールを管理"""
        merge_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(merge_frame, text="2. 連語結合")

        # 上段: 元の分かち書き（ストップワード除去前）を表示するエリア
        ttk.Label(merge_frame, text="分かち書き（ストップワード除去前）:").pack(anchor=tk.W)
        self.pre_token_area = scrolledtext.ScrolledText(merge_frame, width=100, height=12, wrap=tk.WORD)
        self.pre_token_area.pack(fill=tk.BOTH, expand=False, pady=4)

        preview_btn_frame = ttk.Frame(merge_frame); preview_btn_frame.pack(fill=tk.X)
        ttk.Button(preview_btn_frame, text="元テキストを分かち書き表示", command=self.show_pre_tokenized).pack(side=tk.LEFT, padx=4)
        ttk.Button(preview_btn_frame, text="分かち書きを更新(再解析)", command=self.update_pre_tokens).pack(side=tk.LEFT, padx=4)

        # 中段: ルール作成・一覧
        rule_frame = ttk.LabelFrame(merge_frame, text="結合ルール（2〜4語）", padding=6)
        rule_frame.pack(fill=tk.BOTH, expand=False, pady=6)

        control_row = ttk.Frame(rule_frame); control_row.pack(fill=tk.X, pady=4)
        ttk.Label(control_row, text="語数:").pack(side=tk.LEFT, padx=4)
        self.merge_len_var = tk.IntVar(value=2)
        ttk.Combobox(control_row, values=[2,3,4], textvariable=self.merge_len_var, width=4, state="readonly").pack(side=tk.LEFT)

        ttk.Label(control_row, text="結合する語（スペース区切り）:").pack(side=tk.LEFT, padx=6)
        self.merge_seq_entry = ttk.Entry(control_row, width=40); self.merge_seq_entry.pack(side=tk.LEFT, padx=4)

        ttk.Label(control_row, text="結合後の語:").pack(side=tk.LEFT, padx=6)
        self.merge_to_entry = ttk.Entry(control_row, width=20); self.merge_to_entry.pack(side=tk.LEFT, padx=4)

        ttk.Button(control_row, text="追加", command=self.add_merge_rule).pack(side=tk.LEFT, padx=4)
        ttk.Button(control_row, text="削除", command=self.remove_selected_merge_rule).pack(side=tk.LEFT, padx=4)

        # ルール一覧
        self.merge_rule_listbox = tk.Listbox(rule_frame, height=6)
        self.merge_rule_listbox.pack(fill=tk.BOTH, expand=True, pady=4)

        # 下段: プレビュー・適用
        action_frame = ttk.Frame(merge_frame); action_frame.pack(fill=tk.X, pady=6)
        ttk.Button(action_frame, text="プレビュー（結合後の分かち書き）", command=self.apply_merge_rules_preview).pack(side=tk.LEFT, padx=6)
        ttk.Button(action_frame, text="適用して編集領域を更新（ストップワード除去後）", command=self.apply_merge_rules_and_update_edit_area).pack(side=tk.LEFT, padx=6)

        # プレビュー表示領域
        ttk.Label(merge_frame, text="プレビュー:").pack(anchor=tk.W, pady=(8,0))
        self.merge_preview_area = scrolledtext.ScrolledText(merge_frame, width=100, height=10, wrap=tk.WORD)
        self.merge_preview_area.pack(fill=tk.BOTH, expand=True, pady=4)

        # 初期化表示
        self.update_pre_tokens()

    def setup_edit_tab(self):
        edit_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(edit_frame, text="3. 単語編集")

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

        self.edit_area = scrolledtext.ScrolledText(right_frame, width=60, height=15, wrap=tk.WORD)
        self.edit_area.pack(fill=tk.BOTH, expand=True, pady=5)

        # パラメータと実行ボタン（サブタブで機能ごとに分割）
        param_notebook = ttk.Notebook(right_frame)
        param_notebook.pack(fill=tk.BOTH, expand=True, pady=5)

        # ===== タブ1: WordCloud生成 =====
        wc_tab = ttk.Frame(param_notebook)
        param_notebook.add(wc_tab, text="📊 WordCloud")
        
        wc_params = ttk.Frame(wc_tab, padding=10)
        wc_params.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(wc_params, text="幅:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        self.wc_width_var = tk.IntVar(value=1000)
        ttk.Spinbox(wc_params, from_=100, to=5000, textvariable=self.wc_width_var, width=7).grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(wc_params, text="高さ:").grid(row=0, column=2, padx=3, pady=2, sticky=tk.W)
        self.wc_height_var = tk.IntVar(value=600)
        ttk.Spinbox(wc_params, from_=100, to=5000, textvariable=self.wc_height_var, width=7).grid(row=0, column=3, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(wc_params, text="形状:").grid(row=1, column=0, padx=3, pady=2, sticky=tk.W)
        self.wc_shape_var = tk.StringVar(value="rectangle")
        shape_frame = ttk.Frame(wc_params)
        shape_frame.grid(row=1, column=1, columnspan=3, padx=3, pady=2, sticky=tk.W)
        ttk.Radiobutton(shape_frame, text="四角形", variable=self.wc_shape_var, value="rectangle").pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(shape_frame, text="楕円形", variable=self.wc_shape_var, value="ellipse").pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(shape_frame, text="カスタム画像", variable=self.wc_shape_var, value="custom").pack(side=tk.LEFT, padx=2)
        
        ttk.Label(wc_params, text="カスタム画像パス:").grid(row=2, column=0, padx=3, pady=2, sticky=tk.W)
        self.wc_custom_image_var = tk.StringVar(value="")
        custom_img_frame = ttk.Frame(wc_params)
        custom_img_frame.grid(row=2, column=1, columnspan=3, padx=3, pady=2, sticky=tk.EW)
        ttk.Entry(custom_img_frame, textvariable=self.wc_custom_image_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(custom_img_frame, text="参照...", command=self.select_wordcloud_image).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(wc_params, text="🎨 WordCloud生成", command=self.on_generate_wordcloud).grid(row=3, column=0, columnspan=4, padx=3, pady=10, sticky=tk.EW)

        # ===== タブ2: 共起ネットワーク生成 =====
        net_tab = ttk.Frame(param_notebook)
        param_notebook.add(net_tab, text="🔗 Network")
        
        net_params = ttk.Frame(net_tab, padding=10)
        net_params.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(net_params, text="ウィンドウサイズ:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        self.window_var = tk.IntVar(value=5)
        ttk.Spinbox(net_params, from_=2, to=20, textvariable=self.window_var, width=7).grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(net_params, text="ウィンドウ形式:").grid(row=0, column=2, padx=3, pady=2, sticky=tk.W)
        self.window_mode_var = tk.StringVar(value="sliding")
        mode_frame = ttk.Frame(net_params)
        mode_frame.grid(row=0, column=3, padx=3, pady=2, sticky=tk.W)
        ttk.Radiobutton(mode_frame, text="スライディング", variable=self.window_mode_var, value="sliding").pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(mode_frame, text="行ごと", variable=self.window_mode_var, value="line").pack(side=tk.LEFT, padx=2)
        
        ttk.Label(net_params, text="ネットワーク幅:").grid(row=1, column=0, padx=3, pady=2, sticky=tk.W)
        self.net_width_var = tk.IntVar(value=1200)
        ttk.Spinbox(net_params, from_=200, to=5000, textvariable=self.net_width_var, width=7).grid(row=1, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(net_params, text="ネットワーク高さ:").grid(row=1, column=2, padx=3, pady=2, sticky=tk.W)
        self.net_height_var = tk.IntVar(value=800)
        ttk.Spinbox(net_params, from_=200, to=5000, textvariable=self.net_height_var, width=7).grid(row=1, column=3, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(net_params, text="表示組数:").grid(row=2, column=0, padx=3, pady=2, sticky=tk.W)
        self.net_edge_count_var = tk.IntVar(value=50)
        ttk.Spinbox(net_params, from_=10, to=500, textvariable=self.net_edge_count_var, width=7).grid(row=2, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(net_params, text="自己ループ:").grid(row=2, column=2, padx=3, pady=2, sticky=tk.W)
        self.self_loop_var = tk.StringVar(value="remove")
        loop_frame = ttk.Frame(net_params)
        loop_frame.grid(row=2, column=3, padx=3, pady=2, sticky=tk.W)
        ttk.Radiobutton(loop_frame, text="削除", variable=self.self_loop_var, value="remove").pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(loop_frame, text="描画", variable=self.self_loop_var, value="keep").pack(side=tk.LEFT, padx=2)
        
        ttk.Label(net_params, text="ノード色:").grid(row=3, column=0, padx=3, pady=2, sticky=tk.W)
        self.network_cmap_var = tk.StringVar(value="Pastel1")
        ttk.Combobox(net_params, values=["Pastel1", "Pastel2", "Set3", "Accent", "tab20"], 
                     textvariable=self.network_cmap_var, state="readonly", width=15).grid(row=3, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(net_params, text="最小共起回数:").grid(row=3, column=2, padx=3, pady=2, sticky=tk.W)
        self.min_cooc_var = tk.IntVar(value=1)
        ttk.Spinbox(net_params, from_=1, to=100, textvariable=self.min_cooc_var, width=7).grid(row=3, column=3, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(net_params, text="詳細オプション:").grid(row=4, column=0, columnspan=4, padx=3, pady=4, sticky=tk.W)
        
        self.collapse_consecutive_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(net_params, text="連続同一単語を1つとして扱う", 
                       variable=self.collapse_consecutive_var).grid(row=5, column=0, columnspan=4, padx=3, pady=2, sticky=tk.W)
        
        self.dedup_pairs_per_line_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(net_params, text="行ごとペア重複カウント制御（同じ行内の同じペアは1回のみ）", 
                       variable=self.dedup_pairs_per_line_var).grid(row=6, column=0, columnspan=4, padx=3, pady=2, sticky=tk.W)
        
        ttk.Button(net_params, text="🔗 ネットワーク生成", command=self.on_generate_network).grid(row=7, column=0, columnspan=4, padx=3, pady=10, sticky=tk.EW)

        # ===== タブ3: 頻度グラフ生成 =====
        freq_tab = ttk.Frame(param_notebook)
        param_notebook.add(freq_tab, text="📈 Frequency")
        
        freq_params = ttk.Frame(freq_tab, padding=10)
        freq_params.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(freq_params, text="最小出現回数:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        self.min_freq_var = tk.IntVar(value=2)
        ttk.Spinbox(freq_params, from_=1, to=20, textvariable=self.min_freq_var, width=7).grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Button(freq_params, text="📊 グラフ生成", command=self.on_generate_frequency_chart).grid(row=1, column=0, columnspan=2, padx=3, pady=10, sticky=tk.EW)

        # ===== タブ4: 共起頻度表表示 =====
        cooc_tab = ttk.Frame(param_notebook)
        param_notebook.add(cooc_tab, text="📋 CoocTable")
        
        cooc_params = ttk.Frame(cooc_tab, padding=10)
        cooc_params.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(cooc_params, text="最小共起回数:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        ttk.Spinbox(cooc_params, from_=1, to=100, textvariable=self.min_cooc_var, width=7).grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Button(cooc_params, text="📋 表を表示", command=self.show_cooccurrence_table).grid(row=1, column=0, columnspan=2, padx=3, pady=10, sticky=tk.EW)

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

        # --- 追加: 共起頻度表タブ（第三タブ内） ---
        self.cooc_frame = ttk.Frame(self.vis_notebook)
        self.vis_notebook.add(self.cooc_frame, text="共起頻度表")

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
                # 行情報と元テキストを保持（生テキスト行）
                self.original_text = combined_text
                self.original_lines = text_lines

                # --- 追加: CSV取り込み直後に MeCab で再解析して pre_tokens_lines を更新し、
                #     original_lines を一貫して「ストップワード除去済みのトークン列（文字列）」に整備する ---
                try:
                    self.update_pre_tokens()  # pre_tokens_lines を生成
                    # original_lines を token-list -> " " 結合 の形式で更新（ストップワード除去）
                    normalized_lines = []
                    for surfaces in self.pre_tokens_lines:
                        if not surfaces:
                            continue
                        line_tokens = [s for s in surfaces if s not in self.stop_words and len(s) > 1]
                        if line_tokens:
                            normalized_lines.append(" ".join(line_tokens))
                    # 上書きしておく（以降の行モードはこの整備済み original_lines を利用する）
                    self.original_lines = normalized_lines
                except Exception:
                    # 失敗しても致命的ではないので続行
                    pass

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

        # ----- 追加: 分かち書き（ストップワード除去前）を行単位で保持 -----
        self.pre_tokens_lines = []
        for raw_line in text.split('\n'):
            surfaces, _ = self.parse_with_pos(raw_line)
            # そのままの分かち書きを保持（ストップワード除去前）
            if surfaces:
                self.pre_tokens_lines.append(surfaces)
            else:
                self.pre_tokens_lines.append([])

        # 【改善】行情報をトークン化済みで保持（行ごとの共起計算用）
        self.original_lines = []
        for raw_line, surfaces in zip(text.split('\n'), self.pre_tokens_lines):
            # ここでストップワードを削除して original_lines を作成（共起計算用）
            line_tokens = [s for s in surfaces if s not in self.stop_words and len(s) > 1]
            if line_tokens:
                self.original_lines.append(" ".join(line_tokens))

        # 全文の解析（pre_tokens と pos を取得）
        surfaces, pos_list = self.parse_with_pos(text)
        if not surfaces:
            messagebox.showerror("エラー", "MeCabの解析結果を取得できませんでした。")
            return

        # フィルタ後トークン（編集領域に入れるのもの）: ストップワードを除去
        self.tokens = [s for s in surfaces if s not in self.stop_words and len(s) > 1]
        # POS cache aligned with filtered tokens
        self.pos_cache = [p for s, p in zip(surfaces, pos_list) if s in self.tokens]

        self.word_freq = Counter(self.tokens)
        self.edit_area.delete(1.0, tk.END)
        self.edit_area.insert(1.0, " ".join(self.tokens))
        self.refresh_word_list()

        # タブ切り替え
        # 新設した「連語結合」タブの影響で編集タブのインデックスが変わっている可能性があるが
        # notebook.select はオブジェクト参照で使えるため既存の動作を維持
        self.notebook.select(self.notebook.index("2. 単語編集") if "2. 単語編集" in [self.notebook.tab(i, option="text") for i in range(self.notebook.index("end"))] else 1)
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
        """
        選択した品詞だけを「保持」して、それ以外を削除するUIに変更。
        複数品詞を選択可能（Ctrl/Shiftで複数選択）で、選択された品詞のみ残します。
        """
        if not self.tokens:
            return

        pos_window = tk.Toplevel(self.root)
        pos_window.title("保持する品詞を選択（複数選択可）")
        pos_window.geometry("420x360")

        ttk.Label(pos_window, text="保持したい品詞を複数選択してください").pack(pady=8)

        # 現在の品詞分布を取得
        current_pos_counts = Counter(self.pos_cache)
        if not current_pos_counts:
            ttk.Label(pos_window, text="品詞情報がありません。").pack(pady=6)
            return

        # 品詞一覧を Listbox（複数選択）で表示
        list_frame = ttk.Frame(pos_window)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)

        scroll = ttk.Scrollbar(list_frame)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.pos_listbox = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, yscrollcommand=scroll.set, height=12)
        self.pos_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.config(command=self.pos_listbox.yview)

        # 表示用に "品詞 (件数)" を入れる（後で分割して品詞部分だけを取り出す）
        for pos, cnt in sorted(current_pos_counts.items(), key=lambda x: (-x[1], x[0])):
            self.pos_listbox.insert(tk.END, f"{pos} ({cnt}件)")

        # ヘルプ行
        ttk.Label(pos_window, text="※選択した品詞のみが残ります。選択なしはキャンセル。", foreground="gray").pack(pady=(4,0))

        def perform_keep():
            selection = self.pos_listbox.curselection()
            if not selection:
                messagebox.showwarning("警告", "最低1つの品詞を選択してください。")
                return
            # 選択アイテムから品詞文字列を抽出（"品詞 (件数)" -> 品詞）
            selected_pos = set()
            for i in selection:
                item = self.pos_listbox.get(i)
                pos_str = item.split(' (')[0]
                selected_pos.add(pos_str)

            # self.tokens と self.pos_cache を同時走査して、選択品詞のみ保持
            filtered_tokens = [t for t, p in zip(self.tokens, self.pos_cache) if p in selected_pos]

            # 編集エリアへ反映
            self.edit_area.delete(1.0, tk.END)
            self.edit_area.insert(1.0, " ".join(filtered_tokens))

            # refresh 状態（word_freq, pos_cache などを更新）
            self.refresh_word_list()

            pos_window.destroy()

        btn_frame = ttk.Frame(pos_window)
        btn_frame.pack(pady=8)
        ttk.Button(btn_frame, text="選択品詞のみ保持", command=perform_keep).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_frame, text="キャンセル", command=pos_window.destroy).pack(side=tk.LEFT, padx=6)
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

    def select_wordcloud_image(self):
        """WordCloud用のカスタム画像を選択"""
        filepath = filedialog.askopenfilename(
            filetypes=[("PNG画像", "*.png"), ("JPEG画像", "*.jpg"), ("すべてのファイル", "*.*")]
        )
        if filepath:
            self.wc_custom_image_var.set(filepath)

    def generate_wordcloud(self, word_freq):
        # 既存のウィジェットをクリア
        for widget in self.wordcloud_frame.winfo_children():
            widget.destroy()

        # タブで指定したサイズを使用する（デフォルト値は Spinbox にて設定）
        width = getattr(self, "wc_width_var", tk.IntVar(value=1000)).get()
        height = getattr(self, "wc_height_var", tk.IntVar(value=600)).get()

        # 【新機能】形状に応じたマスク生成
        shape = getattr(self, "wc_shape_var", tk.StringVar(value="rectangle")).get()
        mask = None
        
        if shape == "ellipse":
            # 楕円形マスク生成（デフォルト楕円画像を使用）
            ellipse_path = Path(__file__).parent / "frame_image" / "楕円.png"
            if ellipse_path.exists():
                try:
                    img = Image.open(ellipse_path)
                    img = img.resize((width, height))
                    mask = np.array(img.convert('L'))
                except Exception as e:
                    messagebox.showwarning("警告", f"楕円画像の読み込みに失敗しました: {e}\n四角形で生成します。")
                    mask = None
            else:
                messagebox.showwarning("警告", f"楕円画像が見つかりません: {ellipse_path}\n四角形で生成します。")
                mask = None
        elif shape == "custom":
            # カスタム画像マスク
            img_path = getattr(self, "wc_custom_image_var", tk.StringVar(value="")).get()
            if img_path and Path(img_path).exists():
                try:
                    img = Image.open(img_path)
                    img = img.resize((width, height))
                    mask = np.array(img.convert('L'))
                except Exception as e:
                    messagebox.showwarning("警告", f"カスタム画像の読み込みに失敗しました: {e}\n四角形で生成します。")
                    mask = None

        # WordCloud生成
        wc_kwargs = {
            "width": width,
            "height": height,
            "background_color": 'white',
            "font_path": self.font_path,
            "relative_scaling": 0.5,
            "min_font_size": 10,
            "max_font_size": 100,
            "colormap": 'tab10'
        }
        if mask is not None:
            wc_kwargs["mask"] = mask
            wc_kwargs["contour_width"] = 0  # 【改善】輪郭線を非表示
        
        wc = WordCloud(**wc_kwargs).generate_from_frequencies(word_freq)

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
        
        # --- collapse 対応: トークン列を必要に応じて変換 ---
        def _maybe_collapse(seq):
            if getattr(self, "collapse_consecutive_var", tk.BooleanVar(value=False)).get():
                return self._collapse_consecutive(seq)
            return seq

        if window_mode == "sliding":
            tokens_used = _maybe_collapse(tokens)
            # スライディングウィンドウ形式（従来通り）で tokens_used を使う
            for i in range(len(tokens_used)):
                if tokens_used[i] not in word_freq:
                    continue
                for j in range(i + 1, min(i + window_size, len(tokens_used))):
                    if tokens_used[j] not in word_freq:
                        continue
                    pair = tuple(sorted([tokens_used[i], tokens_used[j]]))
                    cooc_pairs.append(pair)
        else:
            # 行ごと形式：保持した行情報を使用（行内で折りたたむ）
            # pre_tokens_lines を優先的に使い、行ごとに独立して共起ペアを抽出
            dedup_mode = getattr(self, "dedup_pairs_per_line_var", tk.BooleanVar(value=False)).get()
            
            if getattr(self, "pre_tokens_lines", None) and len(self.pre_tokens_lines) > 0:
                for surfaces in self.pre_tokens_lines:
                    if not surfaces:
                        continue
                    # ストップワード除去・長さ条件を統一して適用
                    line_tokens = [s for s in surfaces if s in word_freq]
                    if getattr(self, "collapse_consecutive_var", tk.BooleanVar(value=False)).get():
                        line_tokens = self._collapse_consecutive(line_tokens)
                    # この行内でのペア抽出（行間にまたがらない）
                    seen_pairs_in_line = set() if dedup_mode else None
                    for i in range(len(line_tokens)):
                        for j in range(i + 1, len(line_tokens)):
                            pair = tuple(sorted([line_tokens[i], line_tokens[j]]))
                            if dedup_mode:
                                if pair not in seen_pairs_in_line:
                                    cooc_pairs.append(pair)
                                    seen_pairs_in_line.add(pair)
                            else:
                                cooc_pairs.append(pair)
            else:
                # フォールバック：original_lines から行ごとに抽出
                for line in self.original_lines:
                    if not line.strip():
                        continue
                    line_tokens = line.split()
                    if getattr(self, "collapse_consecutive_var", tk.BooleanVar(value=False)).get():
                        line_tokens = self._collapse_consecutive(line_tokens)
                    # この行内でのペア抽出（行間にまたがらない）
                    seen_pairs_in_line = set() if dedup_mode else None
                    for i in range(len(line_tokens)):
                        if line_tokens[i] not in word_freq:
                            continue
                        for j in range(i + 1, len(line_tokens)):
                            if line_tokens[j] not in word_freq:
                                continue
                            pair = tuple(sorted([line_tokens[i], line_tokens[j]]))
                            if dedup_mode:
                                if pair not in seen_pairs_in_line:
                                    cooc_pairs.append(pair)
                                    seen_pairs_in_line.add(pair)
                            else:
                                cooc_pairs.append(pair)

        cooc_count = Counter(cooc_pairs)

        # --- 最小共起回数でペアを事前にフィルタ ---
        min_cooc = getattr(self, "min_cooc_var", tk.IntVar(value=1)).get()
        cooc_count = Counter({p: c for p, c in cooc_count.items() if c >= min_cooc})

        if not cooc_count:
            ttk.Label(self.network_frame, text=f"共起データがありません（min共起={min_cooc}）").pack(pady=20)
            return

        # ネットワーク構築（指定した組数を使用）
        G = nx.Graph()
        for (word1, word2), count in cooc_count.most_common(edge_count):
            # 自己ループの制御
            if word1 == word2 and self_loop_mode == "remove":
                continue
            if count < min_cooc:
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

        # 【改善】ノードサイズ倍率を適用
        node_size_scale = getattr(self, "node_size_scale_var", tk.DoubleVar(value=1.0)).get()
        scaled_node_sizes = [size * node_size_scale for size in node_sizes]

        # ノード描画
        nx.draw_networkx_nodes(
            G, pos,
            node_size=scaled_node_sizes,
            node_color=colors,
            alpha=0.85,
            ax=ax,
            linewidths=2,
            edgecolors='#222'
        )

        # 【改善】エッジ描画（全エッジを共起強度で視別化）
        for (u, v) in edges:
            weight = G[u][v]['weight']
            norm_weight = weight / max_weight
            
            # 強度に応じた色（グラデーション）
            edge_color = plt.cm.Reds(norm_weight)
            
            # 強度が低い場合は破線
            linestyle = 'solid' if norm_weight > 0.3 else 'dashed'
            edge_width = 1.5 + norm_weight * 6
            
            nx.draw_networkx_edges(
                G, pos,
                edgelist=[(u, v)],
                width=edge_width,
                edge_color=[edge_color],
                style=linestyle,
                ax=ax,
                alpha=0.7
            )

        # ラベル描画
        font_family = self.font_prop.get_name() if getattr(self, "font_prop", None) else "sans-serif"
        font_size_scale = getattr(self, "font_size_scale_var", tk.DoubleVar(value=1.0)).get()
        scaled_font_size = 9 * font_size_scale
        
        nx.draw_networkx_labels(
            G, pos,
            font_size=scaled_font_size,
            font_family=font_family,
            font_weight='bold',
            ax=ax
        )

        ax.set_title(f"共起ネットワーク（{len(G.nodes())} ノード、{len(G.edges())} エッジ）\n{window_mode}形式、{self_loop_mode}", fontsize=16, pad=20, weight='bold')
        ax.axis("off")
        
        # 【改善】凡例を実際のデータ範囲で生成（間隔調整・統一）
        from matplotlib.lines import Line2D
        from matplotlib.patches import Patch
        
        # ノード頻度の範囲を取得
        node_freqs = [word_freq.get(node, 1) for node in G.nodes()]
        min_freq = min(node_freqs) if node_freqs else 1
        max_freq = max(node_freqs) if node_freqs else 1
        mid_freq = (min_freq + max_freq) // 2
        
        # ノードサイズの凡例
        min_node_size = max(300, min_freq * 150)
        mid_node_size = max(300, mid_freq * 150)
        max_node_size = max(300, max_freq * 150)
        
        legend_elements = [
            Line2D([0], [0], marker='o', color='w', markerfacecolor='gray', markersize=np.sqrt(min_node_size/np.pi), 
                   label=f'ノード: 出現{min_freq}回 (最小)'),
            Line2D([0], [0], marker='o', color='w', markerfacecolor='gray', markersize=np.sqrt(mid_node_size/np.pi), 
                   label=f'ノード: 出現{mid_freq}回 (中央)'),
            Line2D([0], [0], marker='o', color='w', markerfacecolor='gray', markersize=np.sqrt(max_node_size/np.pi), 
                   label=f'ノード: 出現{max_freq}回 (最大)'),
        ]
        
        # 【改善】エッジ（共起関係）の凡例 - 実際の描画ロジックと統一
        if weights:
            min_weight_val = min(weights)
            max_weight_val = max(weights)
            weight_range = max_weight_val - min_weight_val
            
            if weight_range == 0:
                weight_range = 1
            
            # 凡例用の4段階サンプル
            sample_weights = [
                min_weight_val,
                min_weight_val + weight_range // 3,
                min_weight_val + weight_range * 2 // 3,
                max_weight_val
            ]
            
            # 空行を追加して見やすくする
            legend_elements.append(Patch(facecolor='none', edgecolor='none', label=''))
            legend_elements.append(Patch(facecolor='none', edgecolor='none', label='エッジ（共起関係）:'))
            
            for w in sample_weights:
                norm_w = (w - min_weight_val) / max(weight_range, 1)
                
                # 実際の描画ロジックと同じ計算
                edge_color_tuple = plt.cm.Reds(norm_w)
                linestyle = 'solid' if norm_w > 0.3 else 'dashed'
                
                legend_elements.append(
                    Line2D([0], [0], color=edge_color_tuple, linewidth=3, linestyle=linestyle,
                           label=f'共起{int(w)}回 ({norm_w:.0%})')
                )
        
        # 凡例を配置（自動調整）
        legend = ax.legend(handles=legend_elements, loc='upper left', fontsize=7.5, title='凡例', 
                  title_fontsize=8, framealpha=0.95, labelspacing=1.2, handlelength=2.5)
        
        # 凡例のサイズを計算して、プロット領域を調整
        fig.canvas.draw()
        legend_bbox = legend.get_window_extent(renderer=fig.canvas.get_renderer())
        legend_width_inches = legend_bbox.width / fig.dpi
        legend_height_inches = legend_bbox.height / fig.dpi
        
        # 凡例がプロット内に収まるようにサブプロットを調整
        # 左マージンを増やす（凡例の幅に応じて）
        left_margin = min(0.3, 0.1 + legend_width_inches / fig_w)
        fig.subplots_adjust(left=left_margin, top=0.95, bottom=0.05, right=0.98)

        plt.tight_layout(rect=[left_margin, 0.05, 0.98, 0.95])

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
        legend_width_inches = legend_bbox.width / fig.dpi
        legend_height_inches = legend_bbox.height / fig.dpi
        
        # 凡例がプロット内に収まるようにサブプロットを調整
        # 左マージンを増やす（凡例の幅に応じて）
        left_margin = min(0.3, 0.1 + legend_width_inches / fig_w)
        fig.subplots_adjust(left=left_margin, top=0.95, bottom=0.05, right=0.98)

        plt.tight_layout(rect=[left_margin, 0.05, 0.98, 0.95])

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

    def show_cooccurrence_table(self):
        """共起ペアの頻度を可視化タブ内で表示（CSV出力可能）"""
        # clear previous contents
        for w in self.cooc_frame.winfo_children():
            w.destroy()

        text = self.edit_area.get(1.0, tk.END).strip()
        if not text:
            ttk.Label(self.cooc_frame, text="単語データがありません。").pack(pady=10)
            return

        tokens = text.split()
        if len(tokens) < 2:
            ttk.Label(self.cooc_frame, text="共起ペアを計算するには単語が2つ以上必要です。").pack(pady=10)
            return

        word_freq = Counter(tokens)
        window_size = self.window_var.get()
        window_mode = getattr(self, "window_mode_var", tk.StringVar(value="sliding")).get()
        collapse = getattr(self, "collapse_consecutive_var", tk.BooleanVar(value=False)).get()

        # ペア抽出（collapse を反映）
        cooc_pairs = []
        def maybe_collapse(seq):
            return self._collapse_consecutive(seq) if collapse else seq

        if window_mode == "sliding":
            tokens_used = maybe_collapse(tokens)
            for i in range(len(tokens_used)):
                for j in range(i + 1, min(i + window_size, len(tokens_used))):
                    pair = tuple(sorted([tokens_used[i], tokens_used[j]]))
                    cooc_pairs.append(pair)
        else:
            # 行ごと形式：pre_tokens_lines を優先的に使い、行ごとに独立して抽出
            dedup_mode = getattr(self, "dedup_pairs_per_line_var", tk.BooleanVar(value=False)).get()
            
            if getattr(self, "pre_tokens_lines", None) and len(self.pre_tokens_lines) > 0:
                for surfaces in self.pre_tokens_lines:
                    if not surfaces:
                        continue
                    # ストップワード除去・長さ条件を統一して適用
                    line_tokens = [s for s in surfaces if s in word_freq]
                    if collapse:
                        line_tokens = self._collapse_consecutive(line_tokens)
                    # この行内でのペア抽出（行間にまたがらない）
                    seen_pairs_in_line = set() if dedup_mode else None
                    for i in range(len(line_tokens)):
                        for j in range(i + 1, len(line_tokens)):
                            pair = tuple(sorted([line_tokens[i], line_tokens[j]]))
                            if dedup_mode:
                                if pair not in seen_pairs_in_line:
                                    cooc_pairs.append(pair)
                                    seen_pairs_in_line.add(pair)
                            else:
                                cooc_pairs.append(pair)
            else:
                # フォールバック：original_lines から行ごとに抽出
                for line in self.original_lines:
                    if not line.strip():
                        continue
                    line_tokens = line.split()
                    if collapse:
                        line_tokens = self._collapse_consecutive(line_tokens)
                    # この行内でのペア抽出（行間にまたがらない）
                    seen_pairs_in_line = set() if dedup_mode else None
                    for i in range(len(line_tokens)):
                        for j in range(i + 1, len(line_tokens)):
                            pair = tuple(sorted([line_tokens[i], line_tokens[j]]))
                            if dedup_mode:
                                if pair not in seen_pairs_in_line:
                                    cooc_pairs.append(pair)
                                    seen_pairs_in_line.add(pair)
                            else:
                                cooc_pairs.append(pair)

        if not cooc_pairs:
            ttk.Label(self.cooc_frame, text="共起ペアが見つかりません。").pack(pady=10)
            return

        cooc_count = Counter(cooc_pairs)

        # 最小共起回数フィルタ
        min_cooc = getattr(self, "min_cooc_var", tk.IntVar(value=1)).get()
        items = [(p[0], p[1], c) for p, c in cooc_count.items() if c >= min_cooc]
        if not items:
            ttk.Label(self.cooc_frame, text=f"min共起={min_cooc} を満たすペアがありません。").pack(pady=10)
            return

        # ヘッダー
        ttk.Label(self.cooc_frame, text=f"共起ペア一覧（min共起={min_cooc}、全{len(items)}件）", font=("Meiryo", 12, "bold")).pack(pady=8)

        # Treeview 表示
        tree_frame = ttk.Frame(self.cooc_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        scrollbar = ttk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        columns = ("単語1", "単語2", "共起回数")
        tree = ttk.Treeview(tree_frame, columns=columns, height=20, yscrollcommand=scrollbar.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=tree.yview)

        # ヘッダー
        tree.heading('#0', text='')
        tree.column('#0', width=0, stretch=tk.NO)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150 if col != "共起回数" else 90, anchor=(tk.CENTER if col=="共起回数" else tk.W))

        # データ挿入（頻度順）
        for word1, word2, count in sorted(items, key=lambda x: x[2], reverse=True):
            tree.insert('', tk.END, values=(word1, word2, count))

        # CSV保存
        def export_csv_from_tab():
            filepath = filedialog.asksaveasfilename(
                defaultextension=".csv",
                initialfile="cooccurrence.csv",
                filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
            )
            if not filepath:
                return
           

            try:
                with open(filepath, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(columns)
                    for word1, word2, count in sorted(items, key=lambda x: x[2], reverse=True):
                        writer.writerow([word1, word2, count])
                messagebox.showinfo("完了", f"保存しました: {filepath}")
            except Exception as e:
                messagebox.showerror("エラー", f"保存に失敗しました: {e}")

        btn_frame = ttk.Frame(self.cooc_frame)
        btn_frame.pack(pady=6)
        ttk.Button(btn_frame, text="CSV出力", command=export_csv_from_tab).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_frame, text="共起一覧を更新", command=self.show_cooccurrence_table).pack(side=tk.LEFT, padx=6)

    # --- 追加ユーティリティ ---
    def _collapse_consecutive(self, seq):
        """連続して同じ要素が続く場合、それらを1つにまとめて返す."""
        if not seq:
            return []
        out = [seq[0]]
        for s in seq[1:]:
            if s == out[-1]:
                continue
            out.append(s)
        return out
    
    # --- ここから追加メソッド（setup_merge_tab の直後に配置） ---
    def update_pre_tokens(self):
        """original_text を MeCab で再解析して pre_tokens_lines を更新する（ストップワード除去前）"""
        self.pre_tokens_lines = []
        text = getattr(self, "original_text", "") or self.text_area.get(1.0, tk.END).strip()
        if not text:
            if hasattr(self, "pre_token_area"):
                self.pre_token_area.delete(1.0, tk.END)
            return

        lines = text.split('\n')
        for raw_line in lines:
            surfaces, _ = self.parse_with_pos(raw_line)
            self.pre_tokens_lines.append(surfaces if surfaces else [])

        # 表示を更新
        self.show_pre_tokenized()

    def show_pre_tokenized(self):
        """pre_tokens_lines をテキスト領域に表示（行ごとにスペースで区切る）"""
        if not hasattr(self, "pre_tokens_lines") or not self.pre_tokens_lines:
            self.update_pre_tokens()
        if not hasattr(self, "pre_token_area"):
            return
        self.pre_token_area.delete(1.0, tk.END)
        for line_tokens in self.pre_tokens_lines:
            if line_tokens:
                self.pre_token_area.insert(tk.END, " ".join(line_tokens) + "\n")
            else:
                self.pre_token_area.insert(tk.END, "\n")

    def add_merge_rule(self):
        """ルールを追加（語数チェック・重複チェックあり）"""
        raw = self.merge_seq_entry.get().strip()
        if not raw:
            messagebox.showwarning("警告", "結合する語を入力してください（スペース区切り）。")
            return
        seq = tuple(raw.split())
        n = self.merge_len_var.get()
        if len(seq) != n:
            messagebox.showwarning("警告", f"指定した語数が一致しません（期待: {n}語）。")
            return
        merged = self.merge_to_entry.get().strip() or "".join(seq)
        rule = {"len": n, "seq": seq, "merged": merged}
        if any(r["seq"] == seq for r in self.merge_rules):
            messagebox.showwarning("警告", "同じ語列のルールが既に存在します。")
            return
        self.merge_rules.append(rule)
        self.merge_rule_listbox.insert(tk.END, f'{n}語: {" ".join(seq)} → {merged}')
        # 入力クリア
        self.merge_seq_entry.delete(0, tk.END)
        self.merge_to_entry.delete(0, tk.END)

    def remove_selected_merge_rule(self):
        idx = self.merge_rule_listbox.curselection()
        if not idx:
            return
        i = idx[0]
        self.merge_rule_listbox.delete(i)
        del self.merge_rules[i]

    def apply_rules_to_tokens(self, tokens_line):
        """与えられたトークン行に対して merge_rules を適用して新しいトークン行を返す（長いルール優先）"""
        if not tokens_line:
            return []
        rules_sorted = sorted(self.merge_rules, key=lambda r: r["len"], reverse=True)
        out = []
        i = 0
        L = len(tokens_line)
        while i < L:
            matched = False
            for r in rules_sorted:
                n = r["len"]
                if i + n <= L and tuple(tokens_line[i:i+n]) == r["seq"]:
                    out.append(r["merged"])
                    i += n
                    matched = True
                    break
            if not matched:
                out.append(tokens_line[i])
                i += 1
        return out

    def apply_merge_rules_preview(self):
        """pre_tokens_lines に対してルールを適用した結果をプレビュー表示"""
        if not hasattr(self, "pre_tokens_lines") or not self.pre_tokens_lines:
            self.update_pre_tokens()
        preview_lines = []
        for tokens_line in self.pre_tokens_lines:
            new_line = self.apply_rules_to_tokens(tokens_line) if self.merge_rules else tokens_line
            preview_lines.append(" ".join(new_line))
        if hasattr(self, "merge_preview_area"):
            self.merge_preview_area.delete(1.0, tk.END)
            self.merge_preview_area.insert(tk.END, "\n".join(preview_lines))

    def apply_merge_rules_and_update_edit_area(self):
        """ルールを適用 -> ストップワード除去 -> edit_area に反映 -> 単語リストを更新"""
        if not hasattr(self, "pre_tokens_lines") or not self.pre_tokens_lines:
            self.update_pre_tokens()

        # --- 変更: 編集領域を更新する前に Listbox と同期して最新の stop_words を反映 ---
        if hasattr(self, "stopword_listbox"):
            try:
                self.stop_words = set(self.stopword_listbox.get(0, tk.END))
            except Exception:
                # 万一の取得エラーは既存の self.stop_words を維持
                pass

        merged_tokens_all = []
        for tokens_line in self.pre_tokens_lines:
            new_line = self.apply_rules_to_tokens(tokens_line) if self.merge_rules else tokens_line
            # ストップワード削除（結合は既に行われている）
            filtered = [t for t in new_line if t not in self.stop_words and len(t) > 0]
            merged_tokens_all.extend(filtered)
        # 編集エリアへ反映
        self.edit_area.delete(1.0, tk.END)
        self.edit_area.insert(tk.END, " ".join(merged_tokens_all))
        self.refresh_word_list()
        messagebox.showinfo("完了", "結合ルールを適用し、編集領域を更新しました。")
    # --- 追加メソッドここまで ---

def main():
    root = tk.Tk()
    app = JapaneseTextAnalyzer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
