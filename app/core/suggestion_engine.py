from __future__ import annotations

from typing import Any, Dict, List

from app.models import RuleHit


class SuggestionEngine:
    def __init__(self, suggestion_rules: Dict[str, Any]) -> None:
        self.term_suggestions = suggestion_rules.get("suggestions", {})
        self.category_suggestions = suggestion_rules.get("category_suggestions", {})

    def suggest(self, hits: List[RuleHit]) -> str:
        results: List[str] = []
        for hit in hits:
            if hit.term in self.term_suggestions:
                results.append(f"{hit.term} -> {self.term_suggestions[hit.term]}")
                continue
            if hit.category in self.category_suggestions:
                results.append(f"{hit.term} -> {self.category_suggestions[hit.category]}")
        unique = []
        for line in results:
            if line not in unique:
                unique.append(line)
        return "; ".join(unique)
