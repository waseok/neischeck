from dataclasses import dataclass, field
from typing import Dict, List


@dataclass
class ByteAnalysis:
    total_bytes: int
    sentence_bytes: List[int]
    has_non_bmp: bool


@dataclass
class RuleHit:
    term: str
    category: str
    severity: str
    reason: str
    start: int
    end: int


@dataclass
class CellAnalysisResult:
    source_column: str
    source_value: str
    byte_count: int
    byte_limit: int
    overflow_yn: str
    verdict: str
    hit_terms: List[str] = field(default_factory=list)
    suggested_rewrite: str = ""
    review_note: str = ""
    rule_hits: List[RuleHit] = field(default_factory=list)
    sentence_bytes: List[int] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


@dataclass
class AnalysisContext:
    item_type: str
    identifier: Dict[str, str]
