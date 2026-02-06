import json
from pathlib import Path

from src.telemetry.run_logger import append_run


def test_append_run_writes_jsonl(tmp_path):
    log_path = tmp_path / "runs.jsonl"
    record = {"run_id": "1", "result": {"status": "success"}}

    append_run(record, path=log_path)

    content = log_path.read_text(encoding="utf-8").strip()
    assert json.loads(content) == record
