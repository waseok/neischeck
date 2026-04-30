from __future__ import annotations

import re
from typing import Any, Dict, List

from app.models import RuleHit


class RuleEngine:
    def __init__(
        self,
        forbidden_rules: Dict[str, Any],
        allowlist: Dict[str, Any],
        category_rules: Dict[str, Any],
    ) -> None:
        self.forbidden_rules = forbidden_rules.get("rules", [])
        self.allow_terms = set(allowlist.get("always_allow_terms", []))
        self.title_quotes = set(category_rules.get("context_markers", {}).get("title_quotes", []))

    def _is_title_context(self, text: str, start: int, end: int) -> bool:
        left = text[max(0, start - 2) : start]
        right = text[end : min(len(text), end + 2)]
        if not left or not right:
            return False
        return any(q in left for q in self.title_quotes) and any(q in right for q in self.title_quotes)

    def analyze(self, text: str) -> tuple[str, List[RuleHit]]:
        hits: List[RuleHit] = []
        lowered = text.lower()
        for item in self.forbidden_rules:
            term = item["term"]
            if term in self.allow_terms:
                continue
            pattern = re.escape(term)
            for m in re.finditer(pattern, text, flags=re.IGNORECASE):
                severity = item.get("severity", "review")
                category = item.get("category", "기타")
                reason = "금지 표현 탐지"
                if self._is_title_context(text, m.start(), m.end()):
                    severity = "review"
                    reason = "제목/인용 문맥으로 검토 필요"
                if term.lower() == lowered.strip():
                    reason = "단독 표현"
                hits.append(
                    RuleHit(
                        term=term,
                        category=category,
                        severity=severity,
                        reason=reason,
                        start=m.start(),
                        end=m.end(),
                    )
                )

        if any(h.severity == "forbidden" for h in hits):
            verdict = "확정 위반"
        elif hits:
            verdict = "검토 필요"
        else:
            verdict = "허용"
        return verdict, hits
