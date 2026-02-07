import json
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

from src.schema.definitions import OutlookConfig


def load_system_config(path: Path) -> OutlookConfig:
    data = json.loads(path.read_text(encoding="utf-8"))
    return OutlookConfig(**data)


def run_pipeline_baseline(
    build_fn,
    config_path: Path,
    adapter_factory,
    engine_factory,
) -> Dict[str, Any]:
    """
    Canonical pipeline flow used by notebook and web app:
    1) build configs
    2) load main system config
    3) run engine
    """
    started_at = datetime.now(timezone.utc).isoformat()
    status = "done"
    error: Optional[str] = None
    try:
        build_fn()
        config = load_system_config(config_path)
        adapter = adapter_factory()
        engine = engine_factory(config, adapter)
        engine.run()
    except Exception as exc:
        status = "error"
        error = str(exc)
    ended_at = datetime.now(timezone.utc).isoformat()
    return {
        "status": status,
        "error": error,
        "started_at": started_at,
        "ended_at": ended_at,
        "artifacts": [_file_snapshot(config_path)],
    }


def run_engine_job(jobs: dict, job_id: str, build_fn, config_path: Path, adapter_factory, engine_factory):
    jobs.setdefault(job_id, {"status": "queued"})
    jobs[job_id]["status"] = "running"
    pythoncom = None
    try:
        import pythoncom as _pythoncom  # type: ignore
        pythoncom = _pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pythoncom = None
    pipeline_summary = run_pipeline_baseline(
        build_fn=build_fn,
        config_path=config_path,
        adapter_factory=adapter_factory,
        engine_factory=engine_factory,
    )
    jobs[job_id]["pipeline_summary"] = pipeline_summary
    jobs[job_id]["started_at"] = pipeline_summary["started_at"]
    jobs[job_id]["ended_at"] = pipeline_summary["ended_at"]
    if pipeline_summary["status"] == "done":
        jobs[job_id]["status"] = "done"
    else:
        jobs[job_id]["status"] = f"error: {pipeline_summary.get('error')}"
    if pythoncom is not None:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def summarize_run_window(
    runs: List[Dict[str, Any]],
    start_index: int = 0,
) -> Dict[str, Any]:
    scoped_runs = runs[max(start_index, 0) :]
    total = len(scoped_runs)
    success = 0
    error = 0
    with_output = 0
    latest_error: Optional[str] = None
    workflows: set[str] = set()
    latest_timestamp: Optional[str] = None

    for run in scoped_runs:
        result = run.get("result") or {}
        status = str(result.get("status") or "")
        if status == "success":
            success += 1
        elif status == "error":
            error += 1
            if result.get("error"):
                latest_error = str(result.get("error"))
        if result.get("output_path"):
            with_output += 1
        if run.get("workflow"):
            workflows.add(str(run["workflow"]))
        if run.get("timestamp"):
            latest_timestamp = str(run["timestamp"])

    return {
        "total": total,
        "success": success,
        "error": error,
        "with_output": with_output,
        "latest_error": latest_error,
        "latest_timestamp": latest_timestamp,
        "workflows": sorted(workflows),
    }


def summarize_run_detail_rows(
    runs: List[Dict[str, Any]],
    start_index: int = 0,
    limit: int = 100,
) -> List[Dict[str, Any]]:
    scoped_runs = runs[max(start_index, 0) :]
    rows: List[Dict[str, Any]] = []
    for run in scoped_runs[-max(limit, 1) :]:
        result = run.get("result") or {}
        quality = run.get("quality") or {}
        input_meta = run.get("input") or {}
        rows.append(
            {
                "timestamp": run.get("timestamp"),
                "workflow": run.get("workflow"),
                "action_id": run.get("action_id"),
                "status": result.get("status"),
                "error": result.get("error"),
                "output_path": result.get("output_path"),
                "subject": input_meta.get("subject"),
                "attachment_ext": input_meta.get("attachment_ext"),
                "quality_label": quality.get("label"),
                "quality_score": quality.get("score"),
            }
        )
    rows.reverse()
    return rows


def compute_job_duration_seconds(job: Dict[str, Any]) -> Optional[float]:
    started_at = job.get("started_at")
    ended_at = job.get("ended_at")
    if not started_at or not ended_at:
        return None
    try:
        start_dt = datetime.fromisoformat(str(started_at))
        end_dt = datetime.fromisoformat(str(ended_at))
        seconds = (end_dt - start_dt).total_seconds()
        return round(max(seconds, 0.0), 3)
    except Exception:
        return None


def append_jsonl_record(path: Path, record: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(record, ensure_ascii=False)
    with path.open("a", encoding="utf-8") as f:
        f.write(payload + "\n")


def _file_snapshot(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {"path": str(path), "exists": False, "size": None, "mtime": None}
    stat = path.stat()
    return {
        "path": str(path),
        "exists": True,
        "size": int(stat.st_size),
        "mtime": datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc).isoformat(),
    }


def run_bridge_compile_summary(
    build_fn,
    system_config_path: Path,
    rules_path: Path,
) -> Dict[str, Any]:
    started_at = datetime.now(timezone.utc).isoformat()
    status = "done"
    error: Optional[str] = None
    try:
        build_fn()
    except Exception as exc:
        status = "error"
        error = str(exc)
    ended_at = datetime.now(timezone.utc).isoformat()
    return {
        "status": status,
        "error": error,
        "started_at": started_at,
        "ended_at": ended_at,
        "artifacts": [
            _file_snapshot(system_config_path),
            _file_snapshot(rules_path),
        ],
    }


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


def load_jsonl_runs_tail(path: Path, max_lines: int = 200) -> List[Dict[str, Any]]:
    if not path.exists():
        return []
    try:
        with path.open("r", encoding="utf-8") as f:
            lines = f.readlines()
    except Exception:
        return []
    if max_lines and max_lines > 0:
        lines = lines[-max_lines:]
    runs: List[Dict[str, Any]] = []
    for line in lines:
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
