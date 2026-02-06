import json
import os
import threading
import time
from pathlib import Path
from typing import Any, Dict, List, Tuple, Optional

from src.schema.definitions import IntentSpecification, OutlookConfig


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
    jobs.setdefault(job_id, {"status": "queued"})
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


def parse_feedback_input(text: str) -> Dict[str, Any]:
    raw = (text or "").strip()
    if not raw:
        return {"raw": "", "format": "empty", "parsed": None}
    if raw.startswith("{") or raw.startswith("["):
        try:
            parsed = json.loads(raw)
            return {"raw": raw, "format": "json", "parsed": parsed}
        except json.JSONDecodeError:
            return {"raw": raw, "format": "text", "parsed": None}
    return {"raw": raw, "format": "text", "parsed": None}


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


def run_rule_check(feedback: Dict[str, Any], rules: List[Dict[str, Any]]) -> Dict[str, Any]:
    required_keys = {"subject_filter", "action_id"}
    missing = []
    for idx, rule in enumerate(rules):
        missing_keys = sorted(list(required_keys - set(rule.keys())))
        if missing_keys:
            missing.append({"index": idx, "missing_keys": missing_keys})
    text = (feedback.get("raw") or "").lower()
    hints = []
    if "error" in text or "fail" in text or "失敗" in text:
        hints.append("Feedback indicates errors/failures; confirm rules route to safe actions.")
    if "遅い" in text or "slow" in text:
        hints.append("Feedback mentions slowness; consider async or caching rules.")
    status = "ok" if not missing else "needs_fix"
    return {
        "agent": "rule_check",
        "status": status,
        "summary": f"{len(rules)} rules checked, {len(missing)} with missing keys.",
        "missing": missing,
        "hints": hints,
    }


def _prioritize_ideas(ideas: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    priority_order = {"high": 0, "medium": 1, "low": 2}
    return sorted(ideas, key=lambda item: priority_order.get(item.get("priority", "low"), 2))


def run_workflow_improvement(feedback: Dict[str, Any]) -> Dict[str, Any]:
    text = (feedback.get("raw") or "").lower()
    ideas = []
    if "手動" in text or "manual" in text:
        ideas.append(
            {
                "title": "Template Picker",
                "reason": "Reduce manual typing by reusing common patterns.",
                "priority": "high",
                "action": "Add a small template dropdown above the input.",
            }
        )
    if "遅い" in text or "slow" in text:
        ideas.append(
            {
                "title": "Async Status + Notify",
                "reason": "Users feel latency; visibility and notify reduce waiting.",
                "priority": "high",
                "action": "Show job status and add completion notification.",
            }
        )
    if "不明" in text or "unclear" in text:
        ideas.append(
            {
                "title": "Inline Glossary",
                "reason": "Clarify ambiguous terms at the point of input.",
                "priority": "medium",
                "action": "Add a small glossary popover for key fields.",
            }
        )
    if not ideas:
        ideas = [
            {
                "title": "One-Click Rerun",
                "reason": "Reduce repeat setup when iterating on rules.",
                "priority": "medium",
                "action": "Add a rerun button that reuses last inputs.",
            },
            {
                "title": "Before/After Diff",
                "reason": "Make rule changes visible and auditable.",
                "priority": "low",
                "action": "Render a diff view of results by rule version.",
            },
        ]
    ideas = _prioritize_ideas(ideas)
    return {
        "agent": "workflow_improvement",
        "status": "ok",
        "summary": f"{len(ideas)} improvement ideas generated.",
        "ideas": ideas,
    }


def run_skill_suggestion(feedback: Dict[str, Any]) -> Dict[str, Any]:
    text = (feedback.get("raw") or "").lower()
    suggestions = []
    if "ocr" in text:
        suggestions.append(
            {
                "name": "ocr_quality_audit",
                "description": "Score OCR results and flag low-confidence extracts.",
            }
        )
    if "分類" in text or "classify" in text:
        suggestions.append(
            {
                "name": "document_classifier",
                "description": "Predict document type and route to the right action.",
            }
        )
    if not suggestions:
        suggestions.append(
            {
                "name": "quality_summary",
                "description": "Summarize recent feedback and top failures.",
            }
        )
    return {
        "agent": "skill_suggestion",
        "status": "ok",
        "summary": f"{len(suggestions)} skill ideas generated.",
        "skills": suggestions,
    }


def run_quality_agents(
    feedback_text: str,
    context_text: str,
    rules_text: str,
    run_rule: bool = True,
    run_workflow: bool = True,
    run_skill: bool = True,
) -> Dict[str, Any]:
    feedback = parse_feedback_input(feedback_text)
    context = parse_feedback_input(context_text)
    rules, rules_error = parse_rules_input(rules_text)
    results = []
    if run_rule:
        results.append(run_rule_check(feedback, rules))
    if run_workflow:
        results.append(run_workflow_improvement(feedback))
    if run_skill:
        results.append(run_skill_suggestion(feedback))
    return {
        "input": {
            "feedback": feedback,
            "context": context,
            "rules_count": len(rules),
        },
        "rules_error": rules_error,
        "results": results,
    }


def _build_template_steps(goal: str) -> List[Dict[str, Any]]:
    lowered = (goal or "").lower()
    if "請求" in goal or "invoice" in lowered:
        return [
            {"id": "s1", "action": "fetch_mails", "params": {"subject_filter": "Invoice"}},
            {"id": "s2", "action": "extract_attachment", "params": {"ext": ".pdf"}},
            {"id": "s3", "action": "ocr_process", "params": {"lang": "jpn"}},
            {"id": "s4", "action": "save_result", "params": {"destination": "data/out"}},
        ]
    return [
        {"id": "s1", "action": "fetch_inputs", "params": {}},
        {"id": "s2", "action": "transform", "params": {}},
        {"id": "s3", "action": "save_result", "params": {"destination": "data/out"}},
    ]


def _build_template_spec(
    app_context: str,
    goal: str,
    scope: str,
    success: str,
    artifacts: str,
) -> Dict[str, Any]:
    return {
        "spec_id": f"spec-{int(time.time())}",
        "spec_version": "1.0",
        "domain": "accounting_mail_invoice",
        "intent": goal or "Define processing workflow from user intent.",
        "inputs": {
            "app_context": app_context,
            "scope": scope,
            "artifacts": artifacts,
        },
        "steps": _build_template_steps(goal),
        "verification": {
            "required_fields": ["invoice_number", "amount", "date"] if ("請求" in goal or "invoice" in goal.lower()) else [],
            "min_quality_score": 0.8,
            "success_criteria_note": success,
        },
        "fallback": {"on_failure": "route_manual_review"},
    }


def _generate_spec_with_openai(prompt_text: str, model: str) -> Dict[str, Any]:
    from openai import OpenAI

    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
    completion = client.chat.completions.create(
        model=model,
        temperature=0,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": (
                    "You generate strict JSON only. Output a valid intent spec object with fields: "
                    "spec_id, spec_version='1.0', domain, intent, inputs, steps, verification, fallback. "
                    "steps must have unique ids."
                ),
            },
            {"role": "user", "content": prompt_text},
        ],
    )
    content = completion.choices[0].message.content or "{}"
    return json.loads(content)


def generate_intent_spec(
    app_context: str,
    goal: str,
    scope: str,
    success: str,
    artifacts: str,
    use_ai: bool = False,
    model: str = "gpt-4o-mini",
) -> Tuple[Dict[str, Any], Optional[str], str]:
    template_spec = _build_template_spec(app_context, goal, scope, success, artifacts)

    if not use_ai:
        validated = IntentSpecification(**template_spec)
        return validated.model_dump(), None, "template"

    if not os.getenv("OPENAI_API_KEY"):
        validated = IntentSpecification(**template_spec)
        return validated.model_dump(), "OPENAI_API_KEY not found. Generated template spec instead.", "template"

    prompt_text = (
        "Create an intent spec JSON for this request.\n"
        f"app_context: {app_context}\n"
        f"goal: {goal}\n"
        f"scope: {scope}\n"
        f"success: {success}\n"
        f"artifacts: {artifacts}\n"
    )

    try:
        ai_spec = _generate_spec_with_openai(prompt_text, model=model)
        if "spec_id" not in ai_spec:
            ai_spec["spec_id"] = f"spec-{int(time.time())}"
        ai_spec["spec_version"] = "1.0"
        validated = IntentSpecification(**ai_spec)
        return validated.model_dump(), None, "llm"
    except Exception as exc:
        validated = IntentSpecification(**template_spec)
        return (
            validated.model_dump(),
            f"AI generation failed ({exc}). Generated template spec instead.",
            "template",
        )
