from __future__ import annotations

import re
from dataclasses import dataclass


SENTENCE_SPLIT_REGEX = re.compile(r"(?<=[.!?])\s+|\n+")


@dataclass
class ByteCounterConfig:
    newline_bytes: int = 1


class ByteCounter:
    def __init__(self, config: ByteCounterConfig) -> None:
        if config.newline_bytes not in (1, 2):
            raise ValueError("newline_bytes must be 1 or 2")
        self.config = config

    def _char_bytes(self, ch: str) -> tuple[int, bool]:
        code = ord(ch)
        if ch == "\n":
            return self.config.newline_bytes, False
        if code <= 0x7F:
            return 1, False
        if 0xD800 <= code <= 0xDFFF:
            return 3, True
        if code > 0xFFFF:
            return 3, True
        return 3, False

    def analyze(self, text: str) -> tuple[int, list[int], bool]:
        normalized = text.replace("\r\n", "\n").replace("\r", "\n")
        total = 0
        has_non_bmp = False
        for ch in normalized:
            size, warn = self._char_bytes(ch)
            total += size
            has_non_bmp = has_non_bmp or warn

        sentences = [s for s in SENTENCE_SPLIT_REGEX.split(normalized) if s]
        sentence_bytes = [sum(self._char_bytes(c)[0] for c in sentence) for sentence in sentences]
        return total, sentence_bytes, has_non_bmp
