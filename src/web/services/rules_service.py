import json
from datetime import datetime, timezone
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


def _coerce_bool(value: Any, default: bool = True) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        return value.strip().lower() in {"true", "1", "yes", "y"}
    return default


def _infer_mail_rule_fields_from_steps(steps: Any) -> Dict[str, Any]:
    if not isinstance(steps, list):
        return {}
    action_map = {
        "ocr_process": "ocr_process",
        "save_result": "save_process",
        "save_only": "save_process",
        "save_process": "save_process",
        "unzip_file": "unzip_process",
        "unzip_process": "unzip_process",
    }
    inferred: Dict[str, Any] = {}
    for step in steps:
        if not isinstance(step, dict):
            continue
        action = str(step.get("action") or "")
        params = step.get("params") if isinstance(step.get("params"), dict) else {}
        if action == "fetch_mails" and not inferred.get("subject_filter"):
            subject_filter = params.get("subject_filter")
            if subject_filter:
                inferred["subject_filter"] = subject_filter
        if action == "extract_attachment":
            inferred.setdefault("require_attachment", True)
            target_ext = params.get("ext") or params.get("target_ext")
            if target_ext and not inferred.get("target_ext"):
                inferred["target_ext"] = target_ext
        if action in action_map and not inferred.get("action_id"):
            inferred["action_id"] = action_map[action]
            if action == "ocr_process" and "lang" in params:
                inferred.setdefault("parameters", {})["lang"] = params.get("lang")
            if action in {"save_result", "save_process", "save_only"} and "destination" in params:
                inferred.setdefault("parameters", {})["destination"] = params.get("destination")
    return inferred


def _merge_rule_fields(*sources: Any) -> Dict[str, Any]:
    merged: Dict[str, Any] = {}
    for source in sources:
        if not isinstance(source, dict):
            continue
        for key, value in source.items():
            if value is None:
                continue
            merged[key] = value
    return merged


def build_mail_rule_from_intent_spec(spec: Dict[str, Any]) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    if not isinstance(spec, dict):
        return None, "Intent spec is not a dict."
    inputs = spec.get("inputs") or {}
    if not isinstance(inputs, dict):
        return None, "Intent spec inputs must be a dict."
    mail_rule = inputs.get("mail_rule") or {}
    if mail_rule and not isinstance(mail_rule, dict):
        return None, "Intent spec mail_rule must be a dict."
    inferred = _infer_mail_rule_fields_from_steps(spec.get("steps"))
    fallback_inputs = {
        "subject_filter": inputs.get("subject_filter"),
        "task_name": inputs.get("task_name"),
        "require_attachment": inputs.get("require_attachment"),
        "action_id": inputs.get("action_id"),
        "parameters": inputs.get("parameters"),
        "target_ext": inputs.get("target_ext"),
    }
    merged = _merge_rule_fields(inferred, fallback_inputs, mail_rule)
    subject_filter = merged.get("subject_filter")
    action_id = merged.get("action_id")
    if not subject_filter or not action_id:
        return None, "mail_rule requires subject_filter and action_id."
    task_name = merged.get("task_name") or "AUTO"
    require_attachment = _coerce_bool(merged.get("require_attachment"), default=True)
    parameters = merged.get("parameters") or {}
    if isinstance(parameters, str):
        try:
            parameters = json.loads(parameters)
        except json.JSONDecodeError:
            return None, "mail_rule parameters must be valid JSON."
    if not isinstance(parameters, dict):
        return None, "mail_rule parameters must be a dict."
    rule = {
        "subject_filter": str(subject_filter),
        "task_name": str(task_name),
        "require_attachment": bool(require_attachment),
        "action_id": str(action_id),
        "parameters": parameters,
    }
    target_ext = merged.get("target_ext")
    if target_ext:
        rule["target_ext"] = str(target_ext)
    return rule, None


def _normalize_rule(rule: Dict[str, Any], include_metadata: bool = False) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    if not isinstance(rule, dict):
        return None, "Rule must be a dict."
    subject_filter = rule.get("subject_filter")
    action_id = rule.get("action_id")
    if not subject_filter or not action_id:
        return None, "Rule requires subject_filter and action_id."
    task_name = rule.get("task_name") or "AUTO"
    require_attachment = _coerce_bool(rule.get("require_attachment"), default=True)
    parameters = rule.get("parameters") or {}
    if isinstance(parameters, str):
        try:
            parameters = json.loads(parameters)
        except json.JSONDecodeError:
            return None, "Rule parameters must be valid JSON."
    if not isinstance(parameters, dict):
        return None, "Rule parameters must be a dict."
    normalized = {
        "subject_filter": str(subject_filter),
        "task_name": str(task_name),
        "require_attachment": bool(require_attachment),
        "action_id": str(action_id),
        "parameters": parameters,
    }
    target_ext = rule.get("target_ext") or rule.get("ext")
    if target_ext:
        normalized["target_ext"] = str(target_ext)
    if include_metadata and isinstance(rule.get("rule_source"), dict):
        normalized["rule_source"] = rule.get("rule_source")
    return normalized, None


def _rule_key(rule: Dict[str, Any]) -> Tuple[str, str, bool, str]:
    target_ext = rule.get("target_ext") or ""
    return (
        str(rule.get("subject_filter")),
        str(rule.get("action_id")),
        bool(rule.get("require_attachment")),
        str(target_ext),
    )


def _extract_quality_gate(rule: Dict[str, Any]) -> Optional[float]:
    source = rule.get("rule_source")
    if not isinstance(source, dict):
        return None
    gate = source.get("quality_gate")
    if isinstance(gate, (int, float)):
        return float(gate)
    return None


def append_unique_rules(
    existing_rules: List[Dict[str, Any]],
    incoming_rules: List[Dict[str, Any]],
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    merged = list(existing_rules)
    existing_keys = set()
    invalid = 0
    for rule in existing_rules:
        normalized, error = _normalize_rule(rule, include_metadata=True)
        if error:
            continue
        existing_keys.add(_rule_key(normalized))
    added = 0
    skipped_duplicates = 0
    for rule in incoming_rules:
        normalized, error = _normalize_rule(rule, include_metadata=True)
        if error:
            invalid += 1
            continue
        key = _rule_key(normalized)
        if key in existing_keys:
            skipped_duplicates += 1
            continue
        merged.append(normalized)
        existing_keys.add(key)
        added += 1
    return merged, {
        "added": added,
        "skipped_duplicates": skipped_duplicates,
        "skipped_invalid": invalid,
    }


def build_rule_proposals_from_intent_spec(
    spec: Dict[str, Any],
    min_quality_score_gate: float = 0.8,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    rule, error = build_mail_rule_from_intent_spec(spec)
    if error:
        return [], [error]
    verification = spec.get("verification") if isinstance(spec, dict) else {}
    if not isinstance(verification, dict):
        verification = {}
    quality_gate = verification.get("min_quality_score")
    warnings: List[str] = []
    if isinstance(quality_gate, (int, float)) and quality_gate < min_quality_score_gate:
        warnings.append(
            f"Spec min_quality_score ({quality_gate}) is below gate ({min_quality_score_gate})."
        )
    steps = spec.get("steps") if isinstance(spec, dict) else []
    step_ids = [step.get("id") for step in steps if isinstance(step, dict) and step.get("id")]
    rule_source = {
        "spec_id": spec.get("spec_id"),
        "spec_version": spec.get("spec_version"),
        "domain": spec.get("domain"),
        "intent": spec.get("intent"),
        "step_ids": step_ids,
        "quality_gate": quality_gate,
        "generated_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
    }
    proposed = dict(rule)
    proposed["rule_source"] = rule_source
    return [proposed], warnings


def merge_proposed_rules(
    existing_rules: List[Dict[str, Any]],
    proposed_rules: List[Dict[str, Any]],
    selected_indices: Optional[List[int]] = None,
    min_quality_score_gate: float = 0.8,
    allow_low_quality: bool = False,
    drop_metadata: bool = True,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, Any]]:
    selected = set(selected_indices or range(len(proposed_rules)))
    merged = list(existing_rules)
    remaining: List[Dict[str, Any]] = []
    summary = {
        "selected": len(selected),
        "merged": 0,
        "skipped_duplicates": 0,
        "skipped_conflicts": 0,
        "skipped_invalid": 0,
        "skipped_quality_gate": 0,
    }
    existing_keys = set()
    existing_subjects = set()
    for rule in existing_rules:
        normalized, error = _normalize_rule(rule, include_metadata=True)
        if error:
            continue
        existing_keys.add(_rule_key(normalized))
        subject = normalized.get("subject_filter")
        if subject:
            existing_subjects.add(subject)

    for idx, rule in enumerate(proposed_rules):
        if idx not in selected:
            remaining.append(rule)
            continue
        normalized, error = _normalize_rule(rule, include_metadata=not drop_metadata)
        if error:
            summary["skipped_invalid"] += 1
            remaining.append(rule)
            continue
        quality_gate = _extract_quality_gate(rule)
        if (
            quality_gate is not None
            and quality_gate < min_quality_score_gate
            and not allow_low_quality
        ):
            summary["skipped_quality_gate"] += 1
            remaining.append(rule)
            continue
        key = _rule_key(normalized)
        if key in existing_keys:
            summary["skipped_duplicates"] += 1
            remaining.append(rule)
            continue
        subject = normalized.get("subject_filter")
        if subject in existing_subjects:
            summary["skipped_conflicts"] += 1
            remaining.append(rule)
            continue
        merged.append(normalized)
        existing_keys.add(key)
        if subject:
            existing_subjects.add(subject)
        summary["merged"] += 1
    return merged, remaining, summary
