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
# import japanize_matplotlib  # 日本語フォント対応




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

        # MeCab
        try:
            self.mecab = MeCab.Tagger("-Owakati")
        except Exception:
            messagebox.showerror("警告", "MeCabが見つかりません")
            self.mecab = None


        # データ保持
        self.original_text = ""
        self.tokens = []
        self.word_freq = Counter()
        self.pos_cache = []

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
        ttk.Label(input_frame, text="分析したいテキストを入力してください:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.text_area = scrolledtext.ScrolledText(input_frame, width=100, height=25, wrap=tk.WORD)
        self.text_area.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        # 分かち書きボタン
        ttk.Button(input_frame, text="分かち書き実行 →", command=self.tokenize_text,
                   style="Accent.TButton").grid(row=3, column=0, pady=10)

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

        ttk.Label(param_frame, text="最小出現回数:").pack(side=tk.LEFT, padx=5)
        self.min_freq_var = tk.IntVar(value=2)
        ttk.Spinbox(param_frame, from_=1, to=20, textvariable=self.min_freq_var, width=5).pack(side=tk.LEFT, padx=5)

        ttk.Label(param_frame, text="共起ウィンドウ:").pack(side=tk.LEFT, padx=5)
        self.window_var = tk.IntVar(value=5)
        ttk.Spinbox(param_frame, from_=2, to=20, textvariable=self.window_var, width=5).pack(side=tk.LEFT, padx=5)

        ttk.Button(param_frame, text="可視化実行 →", command=self.visualize,
                   style="Accent.TButton").pack(side=tk.RIGHT, padx=5)

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
            filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")]
        )
        if filepath:
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    self.text_area.delete(1.0, tk.END)
                    self.text_area.insert(1.0, f.read())
            except Exception as e:
                messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました: {e}")

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

    def clear_text(self):
        self.text_area.delete(1.0, tk.END)

    def tokenize_text(self):
        if not self.mecab:
            messagebox.showerror("エラー", "MeCabが初期化されていません。")
            return

        text = self.text_area.get(1.0, tk.END).strip()
        if not text:
            messagebox.showwarning("警告", "テキストを入力してください。")
            return

        self.original_text = text

        # MeCabで分かち書き
        wakati = self.mecab.parse(text)
        # ゼロ幅文字を除去し、全角スペースを半角スペースへ
        wakati = re.sub(r'[\u200B\u200C\u200D\uFEFF]', '', wakati)
        wakati = wakati.replace('\u3000', ' ')
        if not wakati:
            messagebox.showerror("エラー", "MeCabの解析結果を取得できませんでした。")
            return
        words = wakati.strip().split()

        # ストップワード除去
        self.tokens = [w for w in words if w not in self.stop_words and len(w) > 1]

        # 品詞情報をキャッシュする
        self.pos_cache = [self.get_pos(w) for w in self.tokens]

        # 頻度カウント
        self.word_freq = Counter(self.tokens)

        # 編集エリアに表示
        self.edit_area.delete(1.0, tk.END)
        self.edit_area.insert(1.0, ' '.join(self.tokens))

        # 単語リスト更新
        self.refresh_word_list()

        # タブ切り替え
        self.notebook.select(1)
        messagebox.showinfo("完了", f"{len(self.tokens)}個の単語を抽出しました。")

    def refresh_word_list(self):
        # 編集エリアから単語を再カウント
        text = self.edit_area.get(1.0, tk.END).strip()
        self.tokens = text.split()
        self.word_freq = Counter(self.tokens)

        # 品詞キャッシュを更新（編集後の単語を再判定）
        self.pos_cache = [self.get_pos(w) for w in self.tokens]

        # リスト更新
        self.word_listbox.delete(0, tk.END)
        for word, count in self.word_freq.most_common():
            self.word_listbox.insert(tk.END, f"{word} ({count}回)")

    def filter_word_list(self, *args):
        search_term = self.search_var.get()
        self.word_listbox.delete(0, tk.END)

        for word, count in self.word_freq.most_common():
            if search_term.lower() in word.lower():
                self.word_listbox.insert(tk.END, f"{word} ({count}回)")

    @lru_cache(maxsize=4096)
    def get_pos(self, word: str) -> str:
        """単語から主要な品詞を取得する（キャッシュ付き）。"""
        if not word or not self.mecab:
            return ""

        try:
            # -Ochasen 形式は POS をタブ 4 列目に含む（例: 名詞-一般-...）
            tagger = MeCab.Tagger("-Ochasen")
            parsed = tagger.parse(word)
            if not parsed:
                return ""
            first_line = parsed.splitlines()[0]
            parts = first_line.split("\t")
            if len(parts) < 4:
                return ""
            pos_field = parts[3]  # 品詞フィールド
            pos_major = pos_field.split("-")[0] if pos_field else ""
            return pos_major
        except Exception:
            return ""

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

        # WordCloud生成
        wc = WordCloud(
            width=1000,
            height=600,
            background_color='white',
            font_path=self.font_path,
            relative_scaling=0.5,
            min_font_size=10,
            max_font_size=100,
            colormap='tab10'
        ).generate_from_frequencies(word_freq)

        # 描画
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

        # 共起ペア抽出
        cooc_pairs = []
        for i in range(len(tokens)):
            if tokens[i] not in word_freq:
                continue
            for j in range(i + 1, min(i + window_size, len(tokens))):
                if tokens[j] not in word_freq:
                    continue
                pair = tuple(sorted([tokens[i], tokens[j]]))
                cooc_pairs.append(pair)

        cooc_count = Counter(cooc_pairs)

        if not cooc_count:
            ttk.Label(self.network_frame, text="共起データがありません").pack(pady=20)
            return

        # ネットワーク構築
        G = nx.Graph()
        for (word1, word2), count in cooc_count.most_common(50):  # 上位50組
            G.add_edge(word1, word2, weight=count)

        # 描画
        fig, ax = plt.subplots(figsize=(12, 8))
        pos = nx.spring_layout(G, k=2, iterations=50)



        # after building G
        communities = list(nx.community.greedy_modularity_communities(G))
        comm_map = {}
        for idx, nodes in enumerate(communities):
            for n in nodes:
                comm_map[n] = idx
        colors = [cm.tab20(comm_map.get(n, 0) / max(len(communities), 1)) for n in G.nodes()]
        pos = nx.spring_layout(G, k=1.5, iterations=100, seed=42)

        node_sizes = [word_freq.get(node, 1) * 120 for node in G.nodes()]
        edges = G.edges()
        weights = [G[u][v]['weight'] for u, v in edges]
        max_weight = max(weights) if weights else 1
        edge_widths = [0.5 + (w / max_weight) * 4 for w in weights]

        nx.draw_networkx_nodes(G, pos, node_size=node_sizes, node_color=colors, alpha=0.8, ax=ax, linewidths=0.5, edgecolors="#333")
        nx.draw_networkx_edges(G, pos, width=edge_widths, alpha=0.4, edge_color="#888", ax=ax)
        font_family = self.font_prop.get_name() if getattr(self, "font_prop", None) else "sans-serif"
        nx.draw_networkx_labels(G, pos, font_size=10, font_family=font_family, ax=ax)
        ax.set_title("共起ネットワーク（上位50エッジ）", fontsize=16, pad=20)
        ax.axis("off")
        plt.tight_layout()


        canvas = FigureCanvasTkAgg(fig, self.network_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # 保存ボタン
        ttk.Button(self.network_frame, text="画像として保存",
                   command=lambda: self.save_figure(fig, "network")).pack(pady=5)

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

    def save_figure(self, fig, name):
        filepath = filedialog.asksaveasfilename(
            defaultextension=".png",
            initialfile=f"{name}.png",
            filetypes=[("PNG", "*.png"), ("PDF", "*.pdf"), ("SVG", "*.svg")]
        )
        if filepath:
            fig.savefig(filepath, dpi=300, bbox_inches='tight')
            messagebox.showinfo("完了", f"保存しました: {filepath}")


def main():
    root = tk.Tk()
    app = JapaneseTextAnalyzer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
