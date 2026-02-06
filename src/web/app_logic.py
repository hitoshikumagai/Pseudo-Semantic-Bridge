import json
import threading
import time
from pathlib import Path
from typing import Any, Dict, List, Tuple

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


def load_jsonl_runs(path: Path) -> List[Dict[str, Any]]:
    if not path.exists():
        return []
    runs: List[Dict[str, Any]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            runs.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return runs


def summarize_quality(runs: List[Dict[str, Any]]) -> Dict[str, Any]:
    total = len(runs)
    success = 0
    quality_ok = 0
    quality_labeled = 0
    for run in runs:
        if run.get("result", {}).get("status") == "success":
            success += 1
        quality = run.get("quality") or {}
        label = str(quality.get("label") or "").lower()
        score = quality.get("score")
        if label or score is not None:
            quality_labeled += 1
        if label in {"ok", "pass", "good"}:
            quality_ok += 1
        elif isinstance(score, (int, float)) and score >= 0.8:
            quality_ok += 1
    return {
        "total": total,
        "success": success,
        "quality_labeled": quality_labeled,
        "quality_ok": quality_ok,
    }


def _candidate_key(run: Dict[str, Any]) -> Tuple[str, str, bool, str]:
    input_meta = run.get("input") or {}
    subject = str(input_meta.get("subject") or "")
    ext = str(input_meta.get("attachment_ext") or "")
    has_attachment = bool(input_meta.get("has_attachment"))
    action_id = str(run.get("action_id") or "")
    return subject, ext, has_attachment, action_id


def propose_rule_candidates(
    runs: List[Dict[str, Any]],
    min_samples: int = 5,
    min_quality_rate: float = 0.8,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    buckets: Dict[Tuple[str, str, bool, str], Dict[str, Any]] = {}
    for run in runs:
        key = _candidate_key(run)
        bucket = buckets.setdefault(
            key,
            {"samples": 0, "success": 0, "quality_ok": 0, "quality_labeled": 0},
        )
        bucket["samples"] += 1
        if run.get("result", {}).get("status") == "success":
            bucket["success"] += 1
        quality = run.get("quality") or {}
        label = str(quality.get("label") or "").lower()
        score = quality.get("score")
        if label or score is not None:
            bucket["quality_labeled"] += 1
        if label in {"ok", "pass", "good"}:
            bucket["quality_ok"] += 1
        elif isinstance(score, (int, float)) and score >= 0.8:
            bucket["quality_ok"] += 1

    meta_rows: List[Dict[str, Any]] = []
    candidate_rows: List[Dict[str, Any]] = []
    for (subject, ext, has_attachment, action_id), stats in buckets.items():
        samples = stats["samples"]
        if samples < min_samples:
            continue
        quality_rate = (stats["quality_ok"] / samples) if samples else 0.0
        if quality_rate < min_quality_rate:
            continue
        meta_rows.append(
            {
                "subject_filter": subject,
                "target_ext": ext or "*",
                "require_attachment": has_attachment,
                "action_id": action_id,
                "samples": samples,
                "success_rate": round(stats["success"] / samples, 3) if samples else 0.0,
                "quality_rate": round(quality_rate, 3),
            }
        )
        candidate_rows.append(
            {
                "subject_filter": subject,
                "task_name": "AUTO",
                "require_attachment": has_attachment,
                "target_ext": ext or "*",
                "action_id": action_id or "save_only",
                "parameters": {},
            }
        )

    return meta_rows, candidate_rows
