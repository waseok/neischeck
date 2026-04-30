import json
from pathlib import Path
from typing import Any, Dict

from app.utils.resource_path import get_resource_path


DEFAULT_CONFIG_FILES = [
    "settings.json",
    "forbidden_rules.json",
    "suggestion_rules.json",
    "allowlist.json",
    "category_rules.json",
]


class ConfigManager:
    def __init__(self, user_config_dir: Path | None = None) -> None:
        base = Path.cwd() / "data" / "config" if user_config_dir is None else user_config_dir
        self.user_config_dir = base
        self.user_config_dir.mkdir(parents=True, exist_ok=True)
        self.ensure_defaults()

    def ensure_defaults(self) -> None:
        for filename in DEFAULT_CONFIG_FILES:
            target = self.user_config_dir / filename
            if target.exists():
                continue
            source = get_resource_path(Path("config") / filename)
            target.write_text(source.read_text(encoding="utf-8"), encoding="utf-8")

    def load_json(self, filename: str) -> Dict[str, Any]:
        path = self.user_config_dir / filename
        return json.loads(path.read_text(encoding="utf-8"))

    def save_json(self, filename: str, payload: Dict[str, Any]) -> None:
        path = self.user_config_dir / filename
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
