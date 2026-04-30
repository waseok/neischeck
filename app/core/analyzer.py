from __future__ import annotations

from app.core.byte_counter import ByteCounter
from app.core.rule_engine import RuleEngine
from app.core.suggestion_engine import SuggestionEngine
from app.models import CellAnalysisResult


class Analyzer:
    def __init__(
        self,
        byte_counter: ByteCounter,
        rule_engine: RuleEngine,
        suggestion_engine: SuggestionEngine,
    ) -> None:
        self.byte_counter = byte_counter
        self.rule_engine = rule_engine
        self.suggestion_engine = suggestion_engine

    def analyze_cell(self, source_column: str, value: str, byte_limit: int) -> CellAnalysisResult:
        text = "" if value is None else str(value)
        byte_count, sentence_bytes, has_non_bmp = self.byte_counter.analyze(text)
        verdict, hits = self.rule_engine.analyze(text)
        overflow_yn = "Y" if byte_count > byte_limit else "N"

        warnings = []
        if has_non_bmp:
            warnings.append("이모지/비BMP 문자가 포함되어 검토가 필요합니다.")
            if verdict == "허용":
                verdict = "검토 필요"

        suggested = self.suggestion_engine.suggest(hits)
        hit_terms = [h.term for h in hits]
        note = ", ".join(warnings) if warnings else ""
        return CellAnalysisResult(
            source_column=source_column,
            source_value=text,
            byte_count=byte_count,
            byte_limit=byte_limit,
            overflow_yn=overflow_yn,
            verdict=verdict,
            hit_terms=hit_terms,
            suggested_rewrite=suggested,
            review_note=note,
            rule_hits=hits,
            sentence_bytes=sentence_bytes,
            warnings=warnings,
        )
