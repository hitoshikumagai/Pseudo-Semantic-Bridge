import json
from typing import Any, Dict, List


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
    rules, rules_error = _parse_rules_input_internal(rules_text)
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


def _parse_rules_input_internal(text: str) -> tuple[list[dict[str, Any]], str | None]:
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
