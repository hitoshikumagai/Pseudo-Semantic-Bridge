import json
import threading
import time
from pathlib import Path

from src.schema.definitions import OutlookConfig


def load_rules(path: Path):
    if not path.exists():
        return []
    return json.loads(path.read_text(encoding="utf-8"))


def save_rules(path: Path, rules):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(rules, indent=2, ensure_ascii=False), encoding="utf-8")


def load_system_config(path: Path) -> OutlookConfig:
    data = json.loads(path.read_text(encoding="utf-8"))
    return OutlookConfig(**data)


def run_engine_job(jobs: dict, job_id: str, build_fn, config_path: Path, adapter_factory, engine_factory):
    jobs[job_id]["status"] = "running"
    try:
        build_fn()
        config = load_system_config(config_path)
        adapter = adapter_factory()
        engine = engine_factory(config, adapter)
        engine.run()
        jobs[job_id]["status"] = "done"
    except Exception as e:
        jobs[job_id]["status"] = f"error: {e}"


def start_job(jobs: dict, run_fn):
    job_id = f"job-{int(time.time())}"
    jobs[job_id] = {"status": "queued"}
    thread = threading.Thread(target=run_fn, args=(job_id,), daemon=True)
    thread.start()
    return job_id
