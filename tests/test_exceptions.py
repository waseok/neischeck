import pandas as pd

from app.core.analyzer import Analyzer
from app.core.byte_counter import ByteCounter, ByteCounterConfig
from app.core.rule_engine import RuleEngine
from app.core.suggestion_engine import SuggestionEngine


def _analyzer():
    return Analyzer(
        ByteCounter(ByteCounterConfig(newline_bytes=1)),
        RuleEngine({"rules": []}, {"always_allow_terms": []}, {"context_markers": {"title_quotes": []}}),
        SuggestionEngine({"suggestions": {}}),
    )


def test_empty_cell():
    analyzer = _analyzer()
    r = analyzer.analyze_cell("세특", "", 500)
    assert r.byte_count == 0
    assert r.verdict == "허용"


def test_nan_cell():
    analyzer = _analyzer()
    value = pd.NA
    r = analyzer.analyze_cell("세특", "" if pd.isna(value) else str(value), 500)
    assert r.byte_count == 0
