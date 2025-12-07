"""
日本語テキスト分析ツール
WordCloudと共起ネットワークを生成するGUIアプリケーション

必要なライブラリ:
pip install tkinter pillow wordcloud sudachipy sudachidict_core networkx matplotlib japanize-matplotlib

※Sudachiの辞書のインストールが必要です
python -m pip install sudachipy sudachidict_core            
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import sudachipy  # SudachiPy (Apache-2.0); sudachi-dictionary-full includes IPA data under BSD notice that must accompany redistribution
import re
from functools import lru_cache
from pathlib import Path
from typing import Optional
from collections import Counter
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import itertools
import csv
import io
from PIL import Image
import numpy as np

from services.files import FileService
from services.tokenization import TokenizationService
from services.visualization import VisualizationService




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

        # Sudachi 形態素解析
        try:
            config = sudachipy.Config()
            dictionary = sudachipy.Dictionary(config)
            self.sudachi = dictionary.create()
        except Exception:
            messagebox.showerror("警告", "Sudachiが見つかりません")
            self.sudachi = None

        self.token_service = TokenizationService(self.sudachi) if self.sudachi else None
        self.file_service = FileService()
        self.visual_service = VisualizationService()


        # データ保持
        self.original_text = ""
        self.tokens = []
        self.word_freq = Counter()
        self.pos_cache = []
        self.original_lines = []  # 【新機能】行情報を保持

        # --- 追加: 分かち書き（ストップワード除去前）行情報と連語ルール ---
        self.pre_tokens_lines = []          # 各行ごとの Sudachi 分かち書き（ストップワード除去前）
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
       
        ttk.Label(wc_params, text="最小出現回数:").grid(row=3, column=0, padx=3, pady=2, sticky=tk.W)
        self.min_freq_var = tk.IntVar(value=2)
        ttk.Spinbox(wc_params, from_=1, to=20, textvariable=self.min_freq_var, width=7).grid(row=3, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(wc_params, text="詳細オプション:").grid(row=4, column=0, columnspan=4, padx=3, pady=4, sticky=tk.W)
        
        self.dedup_word_per_line_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(wc_params, text="行ごと単語重複カウント制御（同じ行内の同じ単語は1回のみ）", 
                       variable=self.dedup_word_per_line_var).grid(row=5, column=0, columnspan=4, padx=3, pady=2, sticky=tk.W)
        
            
        ttk.Button(wc_params, text="🎨 WordCloud生成", command=self.on_generate_wordcloud).grid(row=6, column=0, columnspan=4, padx=3, pady=10, sticky=tk.EW)

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
        
        ttk.Label(net_params, text="エッジ色:").grid(row=3, column=2, padx=3, pady=2, sticky=tk.W)
        self.edge_cmap_var = tk.StringVar(value="Blues")
        ttk.Combobox(net_params, values=["Blues", "Reds", "Greens", "Purples", "Oranges", "Greys"], 
                     textvariable=self.edge_cmap_var, state="readonly", width=15).grid(row=3, column=3, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(net_params, text="最小共起回数:").grid(row=4, column=0, padx=3, pady=2, sticky=tk.W)
        self.min_cooc_var = tk.IntVar(value=1)
        ttk.Spinbox(net_params, from_=1, to=100, textvariable=self.min_cooc_var, width=7).grid(row=4, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(net_params, text="詳細オプション:").grid(row=5, column=0, columnspan=4, padx=3, pady=4, sticky=tk.W)
        
        self.collapse_consecutive_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(net_params, text="連続同一単語を1つとして扱う", 
                       variable=self.collapse_consecutive_var).grid(row=6, column=0, columnspan=4, padx=3, pady=2, sticky=tk.W)
        
        self.dedup_pairs_per_line_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(net_params, text="行ごとペア重複カウント制御（同じ行内の同じペアは1回のみ）", 
                       variable=self.dedup_pairs_per_line_var).grid(row=7, column=0, columnspan=4, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(net_params, text="表示調整:").grid(row=8, column=0, columnspan=4, padx=3, pady=4, sticky=tk.W)
        
        ttk.Label(net_params, text="ノードサイズ倍率:").grid(row=9, column=0, padx=3, pady=2, sticky=tk.W)
        self.node_size_scale_var = tk.DoubleVar(value=1.0)
        ttk.Spinbox(net_params, from_=0.5, to=3.0, increment=0.1, textvariable=self.node_size_scale_var, width=7).grid(row=9, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(net_params, text="フォントサイズ倍率:").grid(row=9, column=2, padx=3, pady=2, sticky=tk.W)
        self.font_size_scale_var = tk.DoubleVar(value=1.0)
        ttk.Spinbox(net_params, from_=0.5, to=3.0, increment=0.1, textvariable=self.font_size_scale_var, width=7).grid(row=9, column=3, padx=3, pady=2, sticky=tk.W)
        
        self.show_legend_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(net_params, text="凡例を表示", 
                       variable=self.show_legend_var).grid(row=10, column=0, columnspan=4, padx=3, pady=2, sticky=tk.W)
        
        ttk.Button(net_params, text="🔗 ネットワーク生成", command=self.on_generate_network).grid(row=11, column=0, columnspan=4, padx=3, pady=10, sticky=tk.EW)

        # ===== タブ3: 頻度グラフ生成 =====
        freq_tab = ttk.Frame(param_notebook)
        param_notebook.add(freq_tab, text="📈 Frequency")
        
        freq_params = ttk.Frame(freq_tab, padding=10)
        freq_params.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(freq_params, text="最小出現回数:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        ttk.Spinbox(freq_params, from_=1, to=20, textvariable=self.min_freq_var, width=7).grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(freq_params, text="詳細オプション:").grid(row=1, column=0, columnspan=2, padx=3, pady=4, sticky=tk.W)
        
        self.dedup_word_per_line_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(freq_params, text="行ごと単語重複カウント制御（同じ行内の同じ単語は1回のみ）", 
                       variable=self.dedup_word_per_line_var).grid(row=2, column=0, columnspan=2, padx=3, pady=2, sticky=tk.W)
        
        ttk.Button(freq_params, text="📊 グラフ生成", command=self.on_generate_frequency_chart).grid(row=3, column=0, columnspan=2, padx=3, pady=10, sticky=tk.EW)

        # ===== タブ4: 共起頻度表表示 =====
        cooc_tab = ttk.Frame(param_notebook)
        param_notebook.add(cooc_tab, text="📋 CoocTable")
        
        cooc_params = ttk.Frame(cooc_tab, padding=10)
        cooc_params.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(cooc_params, text="最小共起回数:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        ttk.Spinbox(cooc_params, from_=1, to=100, textvariable=self.min_cooc_var, width=7).grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(cooc_params, text="詳細オプション:").grid(row=1, column=0, columnspan=2, padx=3, pady=4, sticky=tk.W)
        
        ttk.Checkbutton(cooc_params, text="行ごとペア重複カウント制御（同じ行内の同じペアは1回のみ）", 
                       variable=self.dedup_pairs_per_line_var).grid(row=2, column=0, columnspan=2, padx=3, pady=2, sticky=tk.W)
        
        ttk.Button(cooc_params, text="📋 表を表示", command=self.show_cooccurrence_table).grid(row=3, column=0, columnspan=2, padx=3, pady=10, sticky=tk.EW)

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
            detection = self.file_service.detect_csv_content(filepath)
            rows = detection.rows
            if not rows:
                messagebox.showwarning("警告", "CSVファイルが空です。")
                return

            # 列選択ダイアログ
            col_window = tk.Toplevel(self.root)
            col_window.title("CSV列選択")
            col_window.geometry("420x360")

            ttk.Label(col_window, text=f"検出エンコーディング: {detection.used_encoding}   推定区切り文字: '{detection.delimiter}'", wraplength=380).pack(pady=6, padx=10)
            ttk.Label(col_window, text="結合する列を選択してください（複数選択可）:", wraplength=380).pack(pady=6, padx=10)

            # ヘッダー行の判定（ユーザに確認）
            header_prompt = "最初の行はヘッダーですか？"
            if detection.has_header_guess:
                header_prompt += " (推定: ヘッダーあり)"

            has_header = messagebox.askyesno("CSVヘッダ", header_prompt)
            header_row = rows[0] if has_header else [f"列{i+1}" for i in range(len(rows[0]))]

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

                combined_text = self.file_service.combine_columns(rows, selected_indices, has_header).strip()
                if not combined_text:
                    messagebox.showwarning("警告", "選択列の結合結果が空でした。別の列を選択してください。")
                    return

                # テキストエリアに確実に挿入
                self.text_area.delete(1.0, tk.END)
                self.text_area.insert(1.0, combined_text)
                # 行情報と元テキストを保持（生テキスト行）
                self.original_text = combined_text
                self.original_lines = combined_text.split("\n")

                # --- 追加: CSV取り込み直後に Sudachi で再解析して pre_tokens_lines を更新し、
                #     original_lines を一貫して「ストップワード除去済みのトークン列（文字列）」に整備する ---
                try:
                    self.update_pre_tokens()
                    if self.pre_tokens_lines:
                        _, filtered = TokenizationService.merge_lines(self.pre_tokens_lines, self.merge_rules, self.stop_words)
                        if filtered:
                            self.original_lines = [" ".join(self._collapse_consecutive(filtered))]
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

    def tokenize_text(self):
        if not self.token_service:
            messagebox.showerror("エラー", "Sudachiが初期化されていません。")
            return

        text = self.text_area.get(1.0, tk.END).strip()
        if not text:
            messagebox.showwarning("警告", "テキストを入力してください。")
            return

        self.original_text = text

        result = self.token_service.tokenize_text(text, self.stop_words)
        if not result.surfaces:
            messagebox.showerror("エラー", "Sudachiの解析結果を取得できませんでした。")
            return

        self.tokens = result.tokens
        self.pos_cache = result.pos_cache
        self.word_freq = result.word_freq
        self.pre_tokens_lines = result.pre_tokens_lines
        self.original_lines = result.original_lines

        self.edit_area.delete(1.0, tk.END)
        self.edit_area.insert(1.0, " ".join(self.tokens))
        self.refresh_word_list()

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
        if not word or not self.sudachi:
            return ""
        tokens = self.sudachi.tokenize(word)
        if not tokens:
            return ""
        pos_field = tokens[0].part_of_speech()[0]
        return pos_field if pos_field else ""


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
        for widget in self.wordcloud_frame.winfo_children():
            widget.destroy()

        width = getattr(self, "wc_width_var", tk.IntVar(value=1000)).get()
        height = getattr(self, "wc_height_var", tk.IntVar(value=600)).get()
        shape = getattr(self, "wc_shape_var", tk.StringVar(value="rectangle")).get()
        custom_image = getattr(self, "wc_custom_image_var", tk.StringVar(value="")).get()

        fig = self.visual_service.build_wordcloud_figure(
            word_freq,
            width=width,
            height=height,
            shape=shape,
            font_path=self.font_path,
            custom_image_path=custom_image,
        )

        canvas = FigureCanvasTkAgg(fig, self.wordcloud_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        ttk.Button(self.wordcloud_frame, text="画像として保存",
                   command=lambda: self.save_figure(fig, "wordcloud")).pack(pady=5)

        if not self.font_path:
            ttk.Label(self.wordcloud_frame, text="※日本語フォントが見つからないため、文字化けする可能性があります。", foreground="red").pack(pady=5)

    def generate_network(self, tokens, word_freq):
        for widget in self.network_frame.winfo_children():
            widget.destroy()

        window_size = self.window_var.get()
        edge_count = getattr(self, "net_edge_count_var", tk.IntVar(value=50)).get()
        self_loop_mode = getattr(self, "self_loop_var", tk.StringVar(value="remove")).get()
        window_mode = getattr(self, "window_mode_var", tk.StringVar(value="sliding")).get()
        collapse_consecutive = getattr(self, "collapse_consecutive_var", tk.BooleanVar(value=False)).get()
        dedup_pairs_per_line = getattr(self, "dedup_pairs_per_line_var", tk.BooleanVar(value=False)).get()
        min_cooc = getattr(self, "min_cooc_var", tk.IntVar(value=1)).get()
        net_width = getattr(self, "net_width_var", tk.IntVar(value=1200)).get()
        net_height = getattr(self, "net_height_var", tk.IntVar(value=800)).get()
        cmap_name = getattr(self, "network_cmap_var", tk.StringVar(value="Pastel1")).get()
        edge_cmap_name = getattr(self, "edge_cmap_var", tk.StringVar(value="Blues")).get()
        node_size_scale = getattr(self, "node_size_scale_var", tk.DoubleVar(value=1.0)).get()
        font_size_scale = getattr(self, "font_size_scale_var", tk.DoubleVar(value=1.0)).get()
        show_legend = getattr(self, "show_legend_var", tk.BooleanVar(value=True)).get()

        fig = self.visual_service.build_network_figure(
            tokens,
            word_freq,
            self.pre_tokens_lines,
            self.original_lines,
            window_mode=window_mode,
            window_size=window_size,
            collapse_consecutive=collapse_consecutive,
            dedup_pairs_per_line=dedup_pairs_per_line,
            self_loop_mode=self_loop_mode,
            edge_count=edge_count,
            min_cooc=min_cooc,
            net_width=net_width,
            net_height=net_height,
            cmap_name=cmap_name,
            edge_cmap_name=edge_cmap_name,
            node_size_scale=node_size_scale,
            font_size_scale=font_size_scale,
            show_legend=show_legend,
        )

        if not fig:
            ttk.Label(self.network_frame, text="表示できるネットワークがありません").pack(pady=20)
            return

        canvas = FigureCanvasTkAgg(fig, self.network_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        ttk.Button(self.network_frame, text="画像として保存",
                   command=lambda: self.save_figure(fig, "network")).pack(pady=5)
        ttk.Button(self.network_frame, text="SVGで保存",
                   command=lambda: self.save_figure(fig, "network", fmt="svg")).pack(pady=5)
    def generate_frequency_chart(self, word_freq):
        for widget in self.freq_frame.winfo_children():
            widget.destroy()

        fig = self.visual_service.build_frequency_figure(word_freq)

        canvas = FigureCanvasTkAgg(fig, self.freq_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # 保存ボタン群
        btn_frame = ttk.Frame(self.freq_frame)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text="画像として保存",
                   command=lambda: self.save_figure(fig, "frequency")).pack(side=tk.LEFT, padx=5)
        
        # CSV 出力ボタン
        def export_frequency_csv():
            filepath = filedialog.asksaveasfilename(
                defaultextension=".csv",
                initialfile="frequency.csv",
                filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
            )
            if not filepath:
                return
            try:
                with open(filepath, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(['単語', '出現回数'])
                    for word, count in sorted(word_freq.items(), key=lambda x: x[1], reverse=True):
                        writer.writerow([word, count])
                messagebox.showinfo("完了", f"保存しました: {filepath}")
            except Exception as e:
                messagebox.showerror("エラー", f"保存に失敗しました: {e}")
        
        ttk.Button(btn_frame, text="CSV出力", command=export_frequency_csv).pack(side=tk.LEFT, padx=5)

    def on_generate_wordcloud(self):
        # 編集エリアから単語・頻度を取得し、最小出現回数でフィルタ
        text = self.edit_area.get(1.0, tk.END).strip()
        if not text:
            messagebox.showwarning("警告", "単語データがありません。")
            return
        
        # 行ごと重複カウント制御オプションを確認
        dedup_word_mode = getattr(self, "dedup_word_per_line_var", tk.BooleanVar(value=False)).get()
        
        if dedup_word_mode and self.original_lines:
            # 共起ネットワークと同じロジック：pre_tokens_lines を優先的に使用
            unique_tokens = []
            if getattr(self, "pre_tokens_lines", None) and len(self.pre_tokens_lines) > 0:
                # pre_tokens_lines がある場合（分かち書き後）
                for surfaces in self.pre_tokens_lines:
                    if not surfaces:
                        continue
                    # ストップワード除去・長さ条件を適用
                    line_tokens = [s for s in surfaces if s not in self.stop_words and len(s) > 1]
                    # 行内で重複排除
                    seen = set()
                    for t in line_tokens:
                        if t not in seen:
                            unique_tokens.append(t)
                            seen.add(t)
            else:
                # フォールバック：original_lines から
                for line in self.original_lines:
                    if not line.strip():
                        continue
                    line_tokens = line.split()
                    # 行内で重複排除
                    seen = set()
                    for t in line_tokens:
                        if t not in seen:
                            unique_tokens.append(t)
                            seen.add(t)
            word_freq = Counter(unique_tokens)
        else:
            # 行ごとカウント無効：単純に全トークンをカウント
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
        
        # 行ごと重複カウント制御オプションを確認
        dedup_word_mode = getattr(self, "dedup_word_per_line_var", tk.BooleanVar(value=False)).get()
        
        if dedup_word_mode and self.original_lines:
            # 共起ネットワークと同じロジック：pre_tokens_lines を優先的に使用
            unique_tokens = []
            if getattr(self, "pre_tokens_lines", None) and len(self.pre_tokens_lines) > 0:
                # pre_tokens_lines がある場合（分かち書き後）
                for surfaces in self.pre_tokens_lines:
                    if not surfaces:
                        continue
                    # ストップワード除去・長さ条件を適用
                    line_tokens = [s for s in surfaces if s not in self.stop_words and len(s) > 1]
                    # 行内で重複排除
                    seen = set()
                    for t in line_tokens:
                        if t not in seen:
                            unique_tokens.append(t)
                            seen.add(t)
            else:
                # フォールバック：original_lines から
                for line in self.original_lines:
                    if not line.strip():
                        continue
                    line_tokens = line.split()
                    # 行内で重複排除
                    seen = set()
                    for t in line_tokens:
                        if t not in seen:
                            unique_tokens.append(t)
                            seen.add(t)
            word_freq = Counter(unique_tokens)
        else:
            # 行ごとカウント無効：単純に全トークンをカウント
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

    def save_figure(self, fig, prefix: str, fmt: str = "png"):
        """matplotlib Figure をファイルに保存する共通処理。"""
        default_ext = f".{fmt}"
        initial_name = f"{prefix}.{fmt}"
        filetypes = [
            (f"{fmt.upper()}ファイル", f"*.{fmt}"),
            ("PNGファイル", "*.png"),
            ("SVGファイル", "*.svg"),
            ("すべてのファイル", "*.*"),
        ]
        filepath = filedialog.asksaveasfilename(
            defaultextension=default_ext,
            initialfile=initial_name,
            filetypes=filetypes,
        )
        if not filepath:
            return
        try:
            ext = Path(filepath).suffix.lower().lstrip(".") or fmt
            fig.savefig(filepath, format=ext, bbox_inches="tight")
            messagebox.showinfo("完了", f"保存しました: {filepath}")
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました: {e}")

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

        # 保存ボタン群
        btn_frame = ttk.Frame(self.freq_frame)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text="画像として保存",
                   command=lambda: self.save_figure(fig, "frequency")).pack(side=tk.LEFT, padx=5)
        
        # CSV 出力ボタン
        def export_frequency_csv():
            filepath = filedialog.asksaveasfilename(
                defaultextension=".csv",
                initialfile="frequency.csv",
                filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
            )
            if not filepath:
                return
            try:
                with open(filepath, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(['単語', '出現回数'])
                    for word, count in sorted(word_freq.items(), key=lambda x: x[1], reverse=True):
                        writer.writerow([word, count])
                messagebox.showinfo("完了", f"保存しました: {filepath}")
            except Exception as e:
                messagebox.showerror("エラー", f"保存に失敗しました: {e}")
        
        ttk.Button(btn_frame, text="CSV出力", command=export_frequency_csv).pack(side=tk.LEFT, padx=5)

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
                # フォールバック：original_lines から
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
        """original_text を Sudachi で再解析して pre_tokens_lines を更新する（ストップワード除去前）"""
        self.pre_tokens_lines = []
        text = getattr(self, "original_text", "") or self.text_area.get(1.0, tk.END).strip()
        if not text:
            if hasattr(self, "pre_token_area"):
                self.pre_token_area.delete(1.0, tk.END)
            return

        lines = text.split('\n')
        for raw_line in lines:
            surfaces, _ = self.token_service.parse_with_pos(raw_line) if self.token_service else ([], [])
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
        return TokenizationService.apply_merge_rules_to_line(tokens_line, self.merge_rules)

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
        if not self.pre_tokens_lines:
            messagebox.showwarning("警告", "分かち書きの取得に失敗しました。テキストを入力してから再実行してください。")
            return

        # --- 変更: 編集領域を更新する前に Listbox と同期して最新の stop_words を反映 ---
        if hasattr(self, "stopword_listbox"):
            try:
                self.stop_words = set(self.stopword_listbox.get(0, tk.END))
            except Exception:
                # 万一の取得エラーは既存の self.stop_words を維持
                pass

        # 連語ルールを適用して分かち書き行を更新し、ストップワード除去後のトークンを取得
        merged_lines, filtered_tokens = TokenizationService.merge_lines(
            self.pre_tokens_lines,
            self.merge_rules,
            self.stop_words,
        )
        self.pre_tokens_lines = merged_lines

        # フィルタ済みトークンが空の場合は安全側で長さ1も残す
        merged_tokens_all = (
            filtered_tokens
            if filtered_tokens
            else [t for line in merged_lines for t in line if t not in self.stop_words and len(t) > 0]
        )

        # original_lines も結合後の内容に合わせて更新（行単位の表示や共起計算で利用）
        self.original_lines = [
            " ".join([t for t in line if t not in self.stop_words and len(t) > 1])
            for line in merged_lines
            if any(t for t in line if t not in self.stop_words and len(t) > 1)
        ]

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
