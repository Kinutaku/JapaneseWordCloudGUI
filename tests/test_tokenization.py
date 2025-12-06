from collections import Counter
from pathlib import Path
import sys

sys.path.append(str(Path(__file__).resolve().parents[1]))

from services.tokenization import TokenizationService


class DummyTagger:
    def __init__(self, responses):
        self.responses = responses

    def parse(self, text):
        return self.responses.get(text, self.responses["default"])


def build_service():
    responses = {
        "default": "人工知能\t*\t*\t名詞-一般\n進化\t*\t*\t名詞-一般\nEOS",
        "人工知能 進化": "人工知能\t*\t*\t名詞-一般\n進化\t*\t*\t名詞-一般\nEOS",
        "人工 知能 進化": "人工\t*\t*\t名詞-一般\n知能\t*\t*\t名詞-一般\n進化\t*\t*\t名詞-一般\nEOS",
    }
    return TokenizationService(DummyTagger(responses))


def test_tokenize_filters_stopwords():
    service = build_service()
    result = service.tokenize_text("人工知能 進化", stop_words={"進化"})
    assert result.tokens == ["人工知能"]
    assert result.word_freq == Counter({"人工知能": 1})
    assert result.original_lines == ["人工知能"]


def test_apply_merge_rules_to_line_prefers_longer_match():
    service = build_service()
    rules = [{"len": 2, "seq": ("人工", "知能"), "merged": "人工知能"}]
    merged = service.apply_merge_rules_to_line(["人工", "知能", "進化"], rules)
    assert merged == ["人工知能", "進化"]


def test_merge_lines_applies_rules_and_filters_stopwords():
    service = build_service()
    pre_tokens_lines = [["人工", "知能", "AI"], ["進化", "未来"]]
    rules = [{"len": 2, "seq": ("人工", "知能"), "merged": "人工知能"}]
    merged_lines, filtered = service.merge_lines(pre_tokens_lines, rules, stop_words={"AI"})
    assert merged_lines[0] == ["人工知能", "AI"]
    assert filtered == ["人工知能", "進化", "未来"]
