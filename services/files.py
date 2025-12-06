from __future__ import annotations

import csv
import io
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Sequence


@dataclass
class CsvDetectionResult:
    rows: List[List[str]]
    used_encoding: str
    delimiter: str
    has_header_guess: bool


class FileService:
    """File-related helpers extracted from the GUI class."""

    def detect_csv_content(self, filepath: str) -> CsvDetectionResult:
        raw = Path(filepath).read_bytes()

        enc_candidates = [
            "utf-8-sig",
            "utf-8",
            "cp932",
            "shift_jis",
            "euc_jp",
            "utf-16",
            "utf-16-le",
            "utf-16-be",
        ]
        decoded: Optional[str] = None
        used_enc: Optional[str] = None
        for enc in enc_candidates:
            try:
                decoded = raw.decode(enc)
                used_enc = enc
                break
            except Exception:
                continue

        if decoded is None:
            decoded = raw.decode("utf-8", errors="replace")
            used_enc = "utf-8 (replace)"

        decoded = decoded.replace("\r\n", "\n").replace("\r", "\n")
        sample = decoded[:4096]
        delimiter = ","
        dialect = None
        try:
            dialect = csv.Sniffer().sniff(sample)
            delimiter = dialect.delimiter
        except Exception:
            for cand in [",", "\t", ";"]:
                try:
                    if cand in sample:
                        delimiter = cand
                        break
                except Exception:
                    continue

        if dialect:
            reader = csv.reader(io.StringIO(decoded), dialect=dialect)
        else:
            reader = csv.reader(io.StringIO(decoded), delimiter=delimiter)
        rows = list(reader)

        has_header_guess = False
        try:
            has_header_guess = csv.Sniffer().has_header(sample)
        except Exception:
            has_header_guess = False

        return CsvDetectionResult(
            rows=rows,
            used_encoding=used_enc or "unknown",
            delimiter=delimiter,
            has_header_guess=has_header_guess,
        )

    @staticmethod
    def combine_columns(rows: Sequence[Sequence[str]], selected_indices: Sequence[int], has_header: bool) -> str:
        if not rows:
            return ""

        data_start = 1 if has_header else 0
        lines: List[str] = []
        for row in rows[data_start:]:
            parts = [row[i] for i in selected_indices if i < len(row)]
            lines.append(" ".join(parts))
        return "\n".join(lines)
