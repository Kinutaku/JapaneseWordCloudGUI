from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Iterable, List, Mapping, Sequence

import matplotlib.pyplot as plt
import networkx as nx
import numpy as np
from matplotlib import cm, font_manager
from wordcloud import WordCloud  # WordCloud is MIT-licensed
from PIL import Image


class VisualizationService:
    """Generate matplotlib figures without GUI coupling."""

    def build_wordcloud_figure(
        self,
        word_freq: Mapping[str, int],
        width: int,
        height: int,
        shape: str,
        font_path: str | None,
        custom_image_path: str | None = None,
    ):
        mask = None
        if shape == "ellipse":
            ellipse_path = Path(__file__).parent.parent / "frame_image" / "楕円.png"
            if ellipse_path.exists():
                img = Image.open(ellipse_path)
                img = img.resize((width, height))
                mask = np.array(img.convert("L"))
        elif shape == "custom" and custom_image_path:
            img_path = Path(custom_image_path)
            if img_path.exists():
                img = Image.open(img_path)
                img = img.resize((width, height))
                mask = np.array(img.convert("L"))

        wc_kwargs = {
            "width": width,
            "height": height,
            "background_color": "white",
            "font_path": font_path,
            "relative_scaling": 0.5,
            "min_font_size": 10,
            "max_font_size": 100,
            "colormap": "tab10",
        }
        if mask is not None:
            wc_kwargs["mask"] = mask
            wc_kwargs["contour_width"] = 0

        wc = WordCloud(**wc_kwargs).generate_from_frequencies(word_freq)

        fig, ax = plt.subplots(figsize=(12, 7))
        ax.imshow(wc, interpolation="bilinear")
        ax.axis("off")
        ax.set_title("WordCloud", fontsize=16, pad=20)
        return fig

    def build_network_figure(
        self,
        tokens: Sequence[str],
        word_freq: Mapping[str, int],
        pre_tokens_lines: Sequence[Sequence[str]] | None,
        original_lines: Sequence[str],
        window_mode: str,
        window_size: int,
        collapse_consecutive: bool,
        dedup_pairs_per_line: bool,
        self_loop_mode: str,
        edge_count: int,
        min_cooc: int,
        net_width: int,
        net_height: int,
        cmap_name: str,
        edge_cmap_name: str = "Blues",
        node_size_scale: float = 1.0,
        font_size_scale: float = 1.0,
        show_legend: bool = True,
        font_family: str | None = None,
    ):
        def _collapse_consecutive(seq: Iterable[str]) -> List[str]:
            result: List[str] = []
            prev = None
            for item in seq:
                if item != prev:
                    result.append(item)
                prev = item
            return result

        cooc_pairs = []
        if window_mode == "sliding":
            tokens_used = list(tokens)
            if collapse_consecutive:
                tokens_used = _collapse_consecutive(tokens_used)
            for i in range(len(tokens_used)):
                if tokens_used[i] not in word_freq:
                    continue
                for j in range(i + 1, min(i + window_size, len(tokens_used))):
                    if tokens_used[j] not in word_freq:
                        continue
                    pair = tuple(sorted([tokens_used[i], tokens_used[j]]))
                    cooc_pairs.append(pair)
        else:
            if pre_tokens_lines:
                for surfaces in pre_tokens_lines:
                    if not surfaces:
                        continue
                    line_tokens = [s for s in surfaces if s in word_freq]
                    if collapse_consecutive:
                        line_tokens = _collapse_consecutive(line_tokens)
                    seen_pairs = set() if dedup_pairs_per_line else None
                    for i in range(len(line_tokens)):
                        for j in range(i + 1, len(line_tokens)):
                            pair = tuple(sorted([line_tokens[i], line_tokens[j]]))
                            if dedup_pairs_per_line:
                                if pair not in seen_pairs:
                                    cooc_pairs.append(pair)
                                    seen_pairs.add(pair)
                            else:
                                cooc_pairs.append(pair)
            else:
                for line in original_lines:
                    if not line.strip():
                        continue
                    line_tokens = line.split()
                    if collapse_consecutive:
                        line_tokens = _collapse_consecutive(line_tokens)
                    seen_pairs = set() if dedup_pairs_per_line else None
                    for i in range(len(line_tokens)):
                        if line_tokens[i] not in word_freq:
                            continue
                        for j in range(i + 1, len(line_tokens)):
                            if line_tokens[j] not in word_freq:
                                continue
                            pair = tuple(sorted([line_tokens[i], line_tokens[j]]))
                            if dedup_pairs_per_line:
                                if pair not in seen_pairs:
                                    cooc_pairs.append(pair)
                                    seen_pairs.add(pair)
                            else:
                                cooc_pairs.append(pair)

        cooc_count = Counter(cooc_pairs)
        cooc_count = Counter({p: c for p, c in cooc_count.items() if c >= min_cooc})
        G = nx.Graph()
        for (word1, word2), count in cooc_count.most_common(edge_count):
            if word1 == word2 and self_loop_mode == "remove":
                continue
            if count < min_cooc:
                continue
            G.add_edge(word1, word2, weight=count)

        if len(G.nodes()) == 0:
            return None

        if not nx.is_connected(G):
            largest_cc = max(nx.connected_components(G), key=len)
            G = G.subgraph(largest_cc).copy()

        if len(G.edges()) > 0:
            max_weight_val = max([d["weight"] for _, _, d in G.edges(data=True)])
            min_weight = max(1, max_weight_val // 5)
            G = nx.Graph([(u, v, d) for u, v, d in G.edges(data=True) if d["weight"] >= min_weight])

        if len(G.nodes()) < 2:
            return None

        fig_w = net_width / 100
        fig_h = net_height / 100
        fig, ax = plt.subplots(figsize=(fig_w, fig_h), facecolor="white")
        ax.set_facecolor("white")

        try:
            pos = nx.kamada_kawai_layout(G, scale=2)
        except Exception:
            layout_k = 2 / (len(G.nodes()) ** 0.5)
            pos = nx.spring_layout(G, k=layout_k, iterations=500, seed=42, scale=2, weight="weight")

        communities = list(nx.community.greedy_modularity_communities(G))
        comm_map = {}
        for idx, nodes in enumerate(communities):
            for n in nodes:
                comm_map[n] = idx

        node_sizes = [max(300, word_freq.get(node, 1) * 150) * node_size_scale for node in G.nodes()]
        edges = G.edges()
        weights = [G[u][v]["weight"] for u, v in edges]
        max_weight = max(weights) if weights else 1
        normalized_weights = [w / max_weight for w in weights]

        try:
            cmap = cm.get_cmap(cmap_name)
        except Exception:
            cmap = cm.get_cmap("Pastel1")

        try:
            edge_cmap = cm.get_cmap(edge_cmap_name)
        except Exception:
            edge_cmap = cm.Blues

        nx.draw_networkx_nodes(
            G,
            pos,
            node_color=[comm_map.get(n, 0) for n in G.nodes()],
            cmap=cmap,
            node_size=node_sizes,
            ax=ax,
            alpha=0.9,
        )
        nx.draw_networkx_edges(
            G,
            pos,
            width=[1 + w * 4 for w in normalized_weights],
            edge_color=normalized_weights,
            edge_cmap=edge_cmap,
            alpha=0.6,
            ax=ax,
        )
        
        scaled_font_size = 10 * font_size_scale
        label_kwargs = {"font_size": scaled_font_size, "ax": ax}
        if font_family:
            label_kwargs["font_family"] = font_family
        nx.draw_networkx_labels(G, pos, **label_kwargs)
        ax.axis("off")
        title_kwargs = {"fontsize": 16, "pad": 20}
        if font_family:
            title_kwargs["fontname"] = font_family
        ax.set_title("共起ネットワーク", **title_kwargs)
        
        # 凡例表示（プロットと重ならないように右側へ退避）
        if show_legend:
            from matplotlib.lines import Line2D
            from matplotlib.patches import Patch
            
            # ノード頻度の範囲を取得
            node_freqs = [word_freq.get(node, 1) for node in G.nodes()]
            min_freq = min(node_freqs) if node_freqs else 1
            max_freq = max(node_freqs) if node_freqs else 1
            mid_freq = (min_freq + max_freq) // 2
            
            # ノードサイズの凡例
            min_node_size = max(300, min_freq * 150) * node_size_scale
            mid_node_size = max(300, mid_freq * 150) * node_size_scale
            max_node_size = max(300, max_freq * 150) * node_size_scale
            
            legend_elements = [
                Line2D([0], [0], marker='o', color='w', markerfacecolor='gray', markersize=np.sqrt(min_node_size/np.pi), 
                       label=f'ノード: 出現{min_freq}回 (最小)'),
                Line2D([0], [0], marker='o', color='w', markerfacecolor='gray', markersize=np.sqrt(mid_node_size/np.pi), 
                       label=f'ノード: 出現{mid_freq}回 (中央)'),
                Line2D([0], [0], marker='o', color='w', markerfacecolor='gray', markersize=np.sqrt(max_node_size/np.pi), 
                       label=f'ノード: 出現{max_freq}回 (最大)'),
            ]
            
            # エッジ（共起関係）の凡例
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

                    # 実際の描画ロジックと同じカラーマップを使用
                    edge_color_tuple = edge_cmap(norm_w)
                    linestyle = 'solid' if norm_w > 0.3 else 'dashed'
                    legend_elements.append(
                        Line2D([0], [0], color=edge_color_tuple, linewidth=3, linestyle=linestyle,
                               label=f'共起{int(w)}回 ({norm_w:.0%})')
                    )
            
            # 凡例を配置（自動調整）
            legend_font = font_manager.FontProperties(family=font_family) if font_family else None
            legend = ax.legend(
                handles=legend_elements,
                loc="upper left",
                bbox_to_anchor=(1.02, 1.0),
                borderaxespad=0.0,
                fontsize=7.5,
                title="凡例",
                title_fontsize=8,
                framealpha=0.95,
                labelspacing=1.2,
                handlelength=2.5,
                prop=legend_font,
            )
            if font_family and legend.get_title():
                legend.get_title().set_fontproperties(legend_font)

            # 凡例ぶんの右余白を確保（動的に計算）
            fig.canvas.draw()
            legend_bbox = legend.get_window_extent(renderer=fig.canvas.get_renderer())
            legend_width_inches = legend_bbox.width / fig.dpi
            right_margin = min(0.35, 0.05 + legend_width_inches / fig_w)
            fig.subplots_adjust(left=0.05, right=max(0.6, 1 - right_margin), top=0.95, bottom=0.05)
        
        return fig

    def build_frequency_figure(self, word_freq: Mapping[str, int]):
        top_words = dict(sorted(word_freq.items(), key=lambda x: x[1], reverse=True)[:30])
        fig, ax = plt.subplots(figsize=(12, 8))
        words = list(top_words.keys())
        counts = list(top_words.values())

        ax.barh(words, counts, color="steelblue")
        ax.set_xlabel("出現回数", fontsize=12)
        ax.set_title(f"単語出現頻度（全{len(word_freq)}単語中の上位30単語）", fontsize=16, pad=20)
        ax.invert_yaxis()
        plt.tight_layout()
        return fig
