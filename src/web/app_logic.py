import json
import os
import threading
import time
from datetime import datetime, timezone
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


def _build_instruction_template(user_instruction: str, domain_hint: str) -> Dict[str, Any]:
    raw = (user_instruction or "").strip()
    lowered = raw.lower()
    tasks: List[str] = []
    constraints: List[str] = []
    follow_up_questions: List[str] = []

    if "請求" in raw or "invoice" in lowered:
        tasks.append("請求書メールを特定し、添付を抽出してOCR処理する")
        constraints.append("個人情報・機密情報を保持したまま保存先を管理する")
    if "分類" in raw or "classify" in lowered:
        tasks.append("文書種別を判定してルーティングする")
    if "遅い" in raw or "slow" in lowered:
        constraints.append("処理時間の短縮が必要")

    if not tasks:
        tasks.append("ユーザー要求を実行可能なステップに分解する")
    if not constraints:
        constraints.append("曖昧な要件を確認質問で補完する")

    follow_up_questions.extend(
        [
            "入力データの対象期間はどれくらいですか？",
            "成功とみなす品質基準は何ですか？",
            "失敗時の扱い（再試行/手動対応）はどうしますか？",
        ]
    )

    return {
        "record_type": "instruction_intake",
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "domain_hint": domain_hint or "accounting_mail_invoice",
        "instruction_raw": raw,
        "intent_summary": raw or "No instruction provided.",
        "tasks": tasks,
        "constraints": constraints,
        "missing_info": [
            "target data range",
            "acceptance criteria",
            "failure handling policy",
        ],
        "follow_up_questions": follow_up_questions,
    }


def _generate_instruction_with_openai(
    user_instruction: str,
    domain_hint: str,
    model: str,
) -> Dict[str, Any]:
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
                    "You extract user intent for automation design. Output strict JSON only with keys: "
                    "intent_summary, tasks, constraints, missing_info, follow_up_questions."
                ),
            },
            {
                "role": "user",
                "content": (
                    f"domain_hint: {domain_hint}\n"
                    f"user_instruction: {user_instruction}\n"
                    "Return concise Japanese text."
                ),
            },
        ],
    )
    content = completion.choices[0].message.content or "{}"
    data = json.loads(content)
    return {
        "record_type": "instruction_intake",
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "domain_hint": domain_hint or "accounting_mail_invoice",
        "instruction_raw": (user_instruction or "").strip(),
        "intent_summary": data.get("intent_summary") or "",
        "tasks": data.get("tasks") or [],
        "constraints": data.get("constraints") or [],
        "missing_info": data.get("missing_info") or [],
        "follow_up_questions": data.get("follow_up_questions") or [],
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
                    "Each step must include id, action, params. "
                    "Allowed action examples: fetch_mails, extract_attachment, ocr_process, save_result, custom_step. "
                    "steps must have unique ids."
                ),
            },
            {"role": "user", "content": prompt_text},
        ],
    )
    content = completion.choices[0].message.content or "{}"
    return json.loads(content)


def _conversation_to_text(conversation: List[Dict[str, Any]]) -> str:
    lines = []
    for msg in conversation or []:
        role = str(msg.get("role") or "user")
        content = str(msg.get("content") or "")
        if not content:
            continue
        lines.append(f"{role}: {content}")
    return "\n".join(lines)


def generate_followup_question(
    conversation: List[Dict[str, Any]],
    domain_hint: str = "accounting_mail_invoice",
    use_ai: bool = False,
    model: str = "gpt-4o-mini",
    round_index: int = 0,
) -> Tuple[str, Optional[str], str]:
    template_questions = [
        "対象メールの件名やキーワードは何ですか？",
        "添付ファイルの種類（PDF/画像など）を教えてください。",
        "OCR後の出力先・保存形式はどうしますか？",
        "成功条件（精度や形式）は何を想定していますか？",
    ]
    if not use_ai:
        return template_questions[round_index % len(template_questions)], None, "template"
    if not os.getenv("OPENAI_API_KEY"):
        return (
            template_questions[round_index % len(template_questions)],
            "OPENAI_API_KEY not found. Generated template question instead.",
            "template",
        )

    convo_text = _conversation_to_text(conversation)
    try:
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
                        "You are a requirements interviewer for mail processing + OCR automation. "
                        "Ask one concise Japanese question to clarify missing info. "
                        "Return JSON only: {\"question\": \"...\"}."
                    ),
                },
                {
                    "role": "user",
                    "content": f"domain_hint: {domain_hint}\nconversation:\n{convo_text}\n",
                },
            ],
        )
        content = completion.choices[0].message.content or "{}"
        payload = json.loads(content)
        question = str(payload.get("question") or "").strip()
        if not question:
            question = template_questions[round_index % len(template_questions)]
        return question, None, "llm"
    except Exception as exc:
        return (
            template_questions[round_index % len(template_questions)],
            f"AI question failed ({exc}). Generated template question instead.",
            "template",
        )


def summarize_conversation(
    conversation: List[Dict[str, Any]],
    use_ai: bool = False,
    model: str = "gpt-4o-mini",
) -> Tuple[List[str], Optional[str], str]:
    user_messages = [str(msg.get("content") or "") for msg in conversation or [] if msg.get("role") == "user"]
    if not user_messages:
        return [], "No user messages to summarize.", "template"
    if not use_ai:
        bullets = [
            f"目的: {user_messages[0]}",
        ]
        if len(user_messages) > 1:
            bullets.append(f"追加要件: {user_messages[-1]}")
        return bullets, None, "template"
    if not os.getenv("OPENAI_API_KEY"):
        bullets = [f"目的: {user_messages[0]}"]
        return bullets, "OPENAI_API_KEY not found. Generated template summary instead.", "template"

    convo_text = _conversation_to_text(conversation)
    try:
        from openai import OpenAI

        client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        completion = client.chat.completions.create(
            model=model,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {
                    "role": "system",
                    "content": "Summarize the conversation into concise Japanese bullet points. Return JSON only: {\"bullets\": [\"...\"]}.",
                },
                {"role": "user", "content": convo_text},
            ],
        )
        content = completion.choices[0].message.content or "{}"
        payload = json.loads(content)
        bullets = payload.get("bullets") or []
        bullets = [str(item).strip() for item in bullets if str(item).strip()]
        if not bullets:
            bullets = [f"目的: {user_messages[0]}"]
        return bullets, None, "llm"
    except Exception as exc:
        bullets = [f"目的: {user_messages[0]}"]
        return bullets, f"AI summary failed ({exc}). Generated template summary instead.", "template"


def _infer_action_from_description(description: str) -> str:
    text = (description or "").lower()
    if "ocr" in text or "文字" in description or "抽出" in description:
        return "ocr_process"
    if "保存" in description or "save" in text:
        return "save_result"
    if "添付" in description or "attachment" in text or "pdf" in text:
        return "extract_attachment"
    if "メール" in description or "mail" in text or "inbox" in text:
        return "fetch_mails"
    return "custom_step"


def _normalize_intent_spec(ai_spec: Dict[str, Any], template_spec: Dict[str, Any]) -> Dict[str, Any]:
    if not isinstance(ai_spec, dict):
        return template_spec
    normalized = dict(template_spec)
    for key, value in ai_spec.items():
        if value is not None:
            normalized[key] = value

    steps = ai_spec.get("steps")
    if not isinstance(steps, list) or not steps:
        steps = template_spec.get("steps", [])

    normalized_steps: List[Dict[str, Any]] = []
    for idx, step in enumerate(steps, start=1):
        if not isinstance(step, dict):
            continue
        step_id = step.get("id") or f"s{idx}"
        action = step.get("action")
        params = step.get("params") if isinstance(step.get("params"), dict) else {}
        if not action:
            desc = step.get("description") or step.get("desc") or step.get("detail") or ""
            action = _infer_action_from_description(desc)
            if desc:
                params = dict(params)
                params.setdefault("description", desc)
        if not action:
            action = "custom_step"
        normalized_steps.append(
            {
                "id": str(step_id),
                "action": str(action),
                "params": params,
            }
        )

    if normalized_steps:
        normalized["steps"] = normalized_steps
    if not isinstance(normalized.get("inputs"), dict):
        normalized["inputs"] = template_spec.get("inputs", {})
    if not isinstance(normalized.get("verification"), dict):
        normalized["verification"] = template_spec.get("verification", {})
    if not isinstance(normalized.get("fallback"), dict):
        normalized["fallback"] = template_spec.get("fallback", {})

    if not normalized.get("spec_id"):
        normalized["spec_id"] = f"spec-{int(time.time())}"
    normalized["spec_version"] = "1.0"
    return normalized


def generate_intent_spec_from_summary(
    summary_bullets: List[str],
    focus: str,
    use_ai: bool = False,
    model: str = "gpt-4o-mini",
) -> Tuple[Dict[str, Any], Optional[str], str]:
    focus_text = (focus or "").strip()
    if not focus_text and summary_bullets:
        focus_text = summary_bullets[0]
    scope = " / ".join(summary_bullets) if summary_bullets else ""
    template_spec = _build_template_spec(
        app_context="メール",
        goal=focus_text or "Define processing workflow from user intent.",
        scope=scope,
        success="",
        artifacts="",
    )

    if not use_ai:
        validated = IntentSpecification(**template_spec)
        return validated.model_dump(), None, "template"

    if not os.getenv("OPENAI_API_KEY"):
        validated = IntentSpecification(**template_spec)
        return validated.model_dump(), "OPENAI_API_KEY not found. Generated template spec instead.", "template"

    prompt_text = (
        "Create an intent spec JSON for this conversation summary.\n"
        f"summary_bullets: {summary_bullets}\n"
        f"focus: {focus_text}\n"
        "Ensure steps include id, action, params."
    )
    try:
        ai_spec = _generate_spec_with_openai(prompt_text, model=model)
        normalized = _normalize_intent_spec(ai_spec, template_spec)
        validated = IntentSpecification(**normalized)
        return validated.model_dump(), None, "llm"
    except Exception as exc:
        validated = IntentSpecification(**template_spec)
        return (
            validated.model_dump(),
            f"AI generation failed ({exc}). Generated template spec instead.",
            "template",
        )


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
        normalized = _normalize_intent_spec(ai_spec, template_spec)
        validated = IntentSpecification(**normalized)
        return validated.model_dump(), None, "llm"
    except Exception as exc:
        validated = IntentSpecification(**template_spec)
        return (
            validated.model_dump(),
            f"AI generation failed ({exc}). Generated template spec instead.",
            "template",
        )


def analyze_user_instruction(
    user_instruction: str,
    domain_hint: str = "accounting_mail_invoice",
    use_ai: bool = False,
    model: str = "gpt-4o-mini",
) -> Tuple[Dict[str, Any], Optional[str], str]:
    template_result = _build_instruction_template(user_instruction, domain_hint=domain_hint)
    if not use_ai:
        return template_result, None, "template"

    if not os.getenv("OPENAI_API_KEY"):
        return template_result, "OPENAI_API_KEY not found. Generated template analysis instead.", "template"

    try:
        ai_result = _generate_instruction_with_openai(
            user_instruction=user_instruction,
            domain_hint=domain_hint,
            model=model,
        )
        return ai_result, None, "llm"
    except Exception as exc:
        return (
            template_result,
            f"AI analysis failed ({exc}). Generated template analysis instead.",
            "template",
        )


def analyze_and_log_user_instruction(
    user_instruction: str,
    log_path: Path,
    domain_hint: str = "accounting_mail_invoice",
    use_ai: bool = False,
    model: str = "gpt-4o-mini",
) -> Tuple[Dict[str, Any], Optional[str], str]:
    result, error, source = analyze_user_instruction(
        user_instruction=user_instruction,
        domain_hint=domain_hint,
        use_ai=use_ai,
        model=model,
    )
    payload = dict(result)
    payload["source"] = source
    append_jsonl_record(log_path, payload)
    return payload, error, source
