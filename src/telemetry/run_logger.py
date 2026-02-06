import json
from pathlib import Path
from typing import Any, Dict

DEFAULT_LOG_PATH = Path("data/logs/psb_run.jsonl")


def append_run(record: Dict[str, Any], path: Path = DEFAULT_LOG_PATH) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(record, ensure_ascii=False)
        with path.open("a", encoding="utf-8") as f:
            f.write(payload + "\n")
    except Exception as e:
        print(f"   ⚠️ Log write failed: {e}")
