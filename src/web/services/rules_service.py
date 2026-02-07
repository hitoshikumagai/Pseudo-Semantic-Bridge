import json
from typing import Any, Dict, List, Optional, Tuple
from pathlib import Path


def load_rules(path: Path) -> List[Dict[str, Any]]:
    if not path.exists():
        return []
    return json.loads(path.read_text(encoding="utf-8"))


def save_rules(path: Path, rules: List[Dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(rules, indent=2, ensure_ascii=False), encoding="utf-8")


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


def parse_rules_input(text: str) -> Tuple[List[Dict[str, Any]], Optional[str]]:
    raw = (text or "").strip()
    if not raw:
        return [], None
    if not (raw.startswith("{") or raw.startswith("[")):
        return [], "Rules input is not JSON. Provide a JSON array of rule objects."
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as e:
        return [], f"Invalid JSON: {e}"
    if isinstance(parsed, dict):
        parsed = [parsed]
    if not isinstance(parsed, list):
        return [], "Rules JSON must be an array of rule objects."
    rules = [item for item in parsed if isinstance(item, dict)]
    if len(rules) != len(parsed):
        return rules, "Some entries were not objects and were skipped."
    return rules, None


def build_mail_rule_from_intent_spec(spec: Dict[str, Any]) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    if not isinstance(spec, dict):
        return None, "Intent spec is not a dict."
    inputs = spec.get("inputs") or {}
    if not isinstance(inputs, dict):
        return None, "Intent spec inputs must be a dict."
    mail_rule = inputs.get("mail_rule") or {}
    if not mail_rule:
        mail_rule = {
            "subject_filter": inputs.get("subject_filter"),
            "task_name": inputs.get("task_name"),
            "require_attachment": inputs.get("require_attachment"),
            "action_id": inputs.get("action_id"),
            "parameters": inputs.get("parameters"),
        }
    if not isinstance(mail_rule, dict):
        return None, "Intent spec mail_rule must be a dict."
    subject_filter = mail_rule.get("subject_filter")
    action_id = mail_rule.get("action_id")
    if not subject_filter or not action_id:
        return None, "mail_rule requires subject_filter and action_id."
    task_name = mail_rule.get("task_name") or "AUTO"
    require_attachment = mail_rule.get("require_attachment")
    if require_attachment is None:
        require_attachment = True
    parameters = mail_rule.get("parameters") or {}
    if isinstance(parameters, str):
        try:
            parameters = json.loads(parameters)
        except json.JSONDecodeError:
            return None, "mail_rule parameters must be valid JSON."
    if not isinstance(parameters, dict):
        return None, "mail_rule parameters must be a dict."
    return {
        "subject_filter": str(subject_filter),
        "task_name": str(task_name),
        "require_attachment": bool(require_attachment),
        "action_id": str(action_id),
        "parameters": parameters,
    }, None
