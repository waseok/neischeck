from app.core.rule_engine import RuleEngine


def _engine():
    forbidden = {"rules": [{"term": "유튜브", "category": "브랜드명", "severity": "review"}]}
    allowlist = {"always_allow_terms": ["IT"]}
    category = {"context_markers": {"title_quotes": ['"', "'"]}}
    return RuleEngine(forbidden, allowlist, category)


def test_rule_hit_review():
    engine = _engine()
    verdict, hits = engine.analyze("유튜브로 학습함")
    assert verdict == "검토 필요"
    assert hits[0].term == "유튜브"


def test_allow_when_no_hit():
    engine = _engine()
    verdict, hits = engine.analyze("일반 수업 참여")
    assert verdict == "허용"
    assert hits == []
