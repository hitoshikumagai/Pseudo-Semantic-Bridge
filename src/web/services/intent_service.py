import json
import os
import time
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple

from src.schema.definitions import IntentSpecification
from src.web.services.run_service import append_jsonl_record


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
    log_path,
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
