from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from typing import Iterable, List, Sequence, Tuple

import MeCab
import ipadic


@dataclass
class TokenizationResult:
    tokens: List[str]
    pos_cache: List[str]
    word_freq: Counter
    pre_tokens_lines: List[List[str]]
    original_lines: List[str]
    surfaces: List[str]
    pos_list: List[str]


class TokenizationService:
    """Utility wrapper around MeCab tokenization logic without GUI side effects."""

    def __init__(self, mecab_tagger: MeCab.Tagger | None = None):
        self.mecab_tagger = mecab_tagger or MeCab.Tagger(f"{ipadic.MECAB_ARGS} -Ochasen")

    def parse_with_pos(self, text: str) -> Tuple[List[str], List[str]]:
        parsed = self.mecab_tagger.parse(text)
        lines = [l for l in parsed.splitlines() if l and l != "EOS"]
        surfaces, pos_list = [], []
        for line in lines:
            parts = line.split("\t")
            if len(parts) >= 4:
                surfaces.append(parts[0])
                pos_list.append(parts[3].split("-")[0])
        return surfaces, pos_list

    def tokenize_text(self, text: str, stop_words: Iterable[str]) -> TokenizationResult:
        stop_set = set(stop_words)
        lines = text.split("\n")

        pre_tokens_lines: List[List[str]] = []
        for raw_line in lines:
            surfaces, _ = self.parse_with_pos(raw_line)
            pre_tokens_lines.append(surfaces)

        original_lines: List[str] = []
        for surfaces in pre_tokens_lines:
            line_tokens = [s for s in surfaces if s not in stop_set and len(s) > 1]
            if line_tokens:
                original_lines.append(" ".join(line_tokens))

        surfaces, pos_list = self.parse_with_pos(text)
        tokens = [s for s in surfaces if s not in stop_set and len(s) > 1]
        pos_cache = [p for s, p in zip(surfaces, pos_list) if s not in stop_set and len(s) > 1]
        word_freq = Counter(tokens)

        return TokenizationResult(
            tokens=tokens,
            pos_cache=pos_cache,
            word_freq=word_freq,
            pre_tokens_lines=pre_tokens_lines,
            original_lines=original_lines,
            surfaces=surfaces,
            pos_list=pos_list,
        )

    @staticmethod
    def apply_merge_rules_to_line(tokens_line: Sequence[str], merge_rules: Sequence[dict]) -> List[str]:
        rules_sorted = sorted(merge_rules, key=lambda r: r["len"], reverse=True)
        out: List[str] = []
        i = 0
        n_tokens = len(tokens_line)
        while i < n_tokens:
            matched = False
            for r in rules_sorted:
                n = r["len"]
                if i + n <= n_tokens and tuple(tokens_line[i:i + n]) == r["seq"]:
                    out.append(r["merged"])
                    i += n
                    matched = True
                    break
            if not matched:
                out.append(tokens_line[i])
                i += 1
        return out

    @staticmethod
    def merge_lines(
        pre_tokens_lines: Sequence[Sequence[str]],
        merge_rules: Sequence[dict],
        stop_words: Iterable[str],
    ) -> Tuple[List[List[str]], List[str]]:
        stop_set = set(stop_words)
        merged_lines: List[List[str]] = []
        filtered_tokens: List[str] = []
        for tokens_line in pre_tokens_lines:
            new_line = (
                TokenizationService.apply_merge_rules_to_line(tokens_line, merge_rules)
                if merge_rules
                else list(tokens_line)
            )
            merged_lines.append(new_line)
            filtered_tokens.extend([t for t in new_line if t not in stop_set and len(t) > 1])
        return merged_lines, filtered_tokens
