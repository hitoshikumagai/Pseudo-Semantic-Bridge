import json
from datetime import datetime
from pathlib import Path

import streamlit as st

from src.web.app_logic import append_unique_rules


def _default_semantic_layer_spec() -> dict:
    return {
        "spec_id": "semantic-layer-blueprint",
        "spec_version": "1.0",
        "updated_at": None,
        "purpose": {
            "objective_statement": "",
            "success_metric": "",
            "priority_domain": "",
            "initial_scope": "",
            "priority_objectives": [
                {"objective": "", "target_metric": "", "owner": "", "priority": "high"}
            ],
        },
        "technical_metadata": {
            "auto_collection_enabled": True,
            "metadata_sources": [
                {"system_type": "warehouse", "system_name": "", "connector": "", "status": "planned"}
            ],
            "lineage_paths": [
                {"source_asset": "", "transform": "", "target_asset": "", "trust_level": "medium"}
            ],
        },
        "business_semantics": {
            "glossary_terms": [
                {"term_id": "", "business_name": "", "technical_field": "", "definition": "", "calc_logic": ""}
            ],
            "kpi_definitions": [
                {"kpi_name": "", "formula": "", "grain": "", "source_of_truth": ""}
            ],
        },
        "federation": {
            "integrated_tools": [
                {"category": "catalog", "tool_name": "", "integration_mode": "federated", "status": "planned"}
            ],
            "tacit_patterns": [{"pattern": "", "meaning": "", "domain": ""}],
        },
        "active_metadata": {
            "ai_enrichment_enabled": True,
            "human_review_required": True,
            "learning_cycle": "weekly",
            "learning_signals": [{"signal_name": "", "source": "", "action": ""}],
        },
        "ownership": {
            "ownership_model": "federated",
            "central_team": "",
            "domain_owners": [{"domain": "", "owner_team": "", "steward": "", "approval_sla_days": 5}],
            "guardrails": "",
        },
        "automation_assets": {
            "rules": [],
            "proposed_rules": [],
            "candidate_meta": [],
            "candidate_rows": [],
            "instruction_intake": None,
            "intent_spec": None,
            "intent_spec_source": None,
            "intent_spec_history": [],
        },
        "architecture_views": {
            "diagram_mode": "table",
            "mermaid_flow": "",
            "last_ai_prompt": "",
        },
    }


def _merge_semantic_spec(default_spec: dict, current_spec: dict) -> dict:
    merged = dict(current_spec) if isinstance(current_spec, dict) else {}
    for key, default_value in default_spec.items():
        current_value = current_spec.get(key) if isinstance(current_spec, dict) else None
        if isinstance(default_value, dict):
            section = dict(current_value) if isinstance(current_value, dict) else {}
            if isinstance(current_value, dict):
                for field, field_default in default_value.items():
                    field_value = current_value.get(field)
                    if isinstance(field_default, list):
                        section[field] = field_value if isinstance(field_value, list) else list(field_default)
                    elif isinstance(field_default, dict):
                        section[field] = field_value if isinstance(field_value, dict) else dict(field_default)
                    else:
                        section[field] = field_default if field_value is None else field_value
            merged[key] = section
            continue
        if current_value is None:
            merged[key] = default_value
            continue
        merged[key] = current_value
    return merged


def _load_semantic_layer_spec(path: Path) -> dict:
    default_spec = _default_semantic_layer_spec()
    if not path.exists():
        return default_spec
    try:
        loaded = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return default_spec
    if not isinstance(loaded, dict):
        return default_spec
    return _merge_semantic_spec(default_spec, loaded)


def _save_semantic_layer_spec(path: Path, spec: dict) -> None:
    payload = _merge_semantic_spec(_default_semantic_layer_spec(), spec if isinstance(spec, dict) else {})
    payload["updated_at"] = datetime.utcnow().replace(microsecond=0).isoformat() + "Z"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def _count_rows(value) -> int:
    if not isinstance(value, list):
        return 0
    return len([row for row in value if isinstance(row, dict)])


def _summarize_semantic_layer(spec: dict) -> dict:
    purpose = spec.get("purpose") if isinstance(spec.get("purpose"), dict) else {}
    technical_metadata = spec.get("technical_metadata") if isinstance(spec.get("technical_metadata"), dict) else {}
    business_semantics = spec.get("business_semantics") if isinstance(spec.get("business_semantics"), dict) else {}
    federation = spec.get("federation") if isinstance(spec.get("federation"), dict) else {}
    active_metadata = spec.get("active_metadata") if isinstance(spec.get("active_metadata"), dict) else {}
    ownership = spec.get("ownership") if isinstance(spec.get("ownership"), dict) else {}

    objective_ready = bool((purpose.get("objective_statement") or "").strip()) and bool(
        (purpose.get("priority_domain") or "").strip()
    )
    technical_ready = _count_rows(technical_metadata.get("metadata_sources")) > 0
    semantics_ready = _count_rows(business_semantics.get("glossary_terms")) > 0
    federation_ready = _count_rows(federation.get("integrated_tools")) > 0
    active_ready = _count_rows(active_metadata.get("learning_signals")) > 0
    ownership_ready = _count_rows(ownership.get("domain_owners")) > 0

    ready_steps = sum(
        [
            int(objective_ready),
            int(technical_ready),
            int(semantics_ready),
            int(federation_ready),
            int(active_ready),
            int(ownership_ready),
        ]
    )
    return {
        "ready_steps": ready_steps,
        "total_steps": 6,
        "readiness_pct": round((ready_steps / 6) * 100, 1),
        "glossary_terms": _count_rows(business_semantics.get("glossary_terms")),
        "metadata_sources": _count_rows(technical_metadata.get("metadata_sources")),
        "domain_owners": _count_rows(ownership.get("domain_owners")),
    }


def _build_semantic_context(spec: dict) -> dict:
    purpose = spec.get("purpose") if isinstance(spec.get("purpose"), dict) else {}
    business_semantics = spec.get("business_semantics") if isinstance(spec.get("business_semantics"), dict) else {}
    technical_metadata = spec.get("technical_metadata") if isinstance(spec.get("technical_metadata"), dict) else {}
    objective_rows = purpose.get("priority_objectives") if isinstance(purpose.get("priority_objectives"), list) else []
    objective_labels = []
    for row in objective_rows:
        if not isinstance(row, dict):
            continue
        objective = str(row.get("objective") or "").strip()
        if objective:
            objective_labels.append(objective)
    return {
        "spec_id": spec.get("spec_id"),
        "spec_version": spec.get("spec_version"),
        "priority_domain": purpose.get("priority_domain"),
        "objective_statement": purpose.get("objective_statement"),
        "priority_objectives": objective_labels,
        "glossary_terms_count": _count_rows(business_semantics.get("glossary_terms")),
        "metadata_sources_count": _count_rows(technical_metadata.get("metadata_sources")),
    }


def _normalize_dict_rows(value) -> list[dict]:
    if not isinstance(value, list):
        return []
    return [row for row in value if isinstance(row, dict)]


def _get_semantic_assets(spec: dict) -> dict:
    assets = spec.get("automation_assets")
    return assets if isinstance(assets, dict) else {}


def _load_working_state_from_semantic(spec: dict) -> None:
    assets = _get_semantic_assets(spec)
    semantic_rules = _normalize_dict_rows(assets.get("rules"))
    semantic_proposed = _normalize_dict_rows(assets.get("proposed_rules"))
    semantic_meta = _normalize_dict_rows(assets.get("candidate_meta"))
    semantic_candidates = _normalize_dict_rows(assets.get("candidate_rows"))
    semantic_instruction = assets.get("instruction_intake")
    semantic_intent_spec = assets.get("intent_spec")
    semantic_intent_source = assets.get("intent_spec_source")
    semantic_intent_history = _normalize_dict_rows(assets.get("intent_spec_history"))

    if semantic_rules:
        st.session_state["rules"] = semantic_rules
    if semantic_proposed:
        st.session_state["proposed_rules"] = semantic_proposed
    if semantic_meta:
        st.session_state["ai_meta"] = semantic_meta
    if semantic_candidates:
        st.session_state["ai_candidates"] = semantic_candidates
    if isinstance(semantic_instruction, dict):
        st.session_state["instruction_intake"] = semantic_instruction
    if isinstance(semantic_intent_spec, dict):
        st.session_state["intent_spec"] = semantic_intent_spec
    if isinstance(semantic_intent_source, str):
        st.session_state["intent_spec_source"] = semantic_intent_source
    if semantic_intent_history:
        st.session_state["intent_spec_history"] = semantic_intent_history
    elif isinstance(semantic_intent_spec, dict):
        _upsert_intent_history(semantic_intent_spec, semantic_intent_source if isinstance(semantic_intent_source, str) else None)


def _collect_semantic_payload_from_state(base_spec: dict) -> dict:
    merged = _merge_semantic_spec(_default_semantic_layer_spec(), base_spec if isinstance(base_spec, dict) else {})
    architecture_views = merged.get("architecture_views") if isinstance(merged.get("architecture_views"), dict) else {}
    architecture_views["diagram_mode"] = st.session_state.get("diagram_mode") or architecture_views.get("diagram_mode") or "table"
    architecture_views["mermaid_flow"] = st.session_state.get("mermaid_flow") or architecture_views.get("mermaid_flow") or ""
    architecture_views["last_ai_prompt"] = st.session_state.get("mermaid_ai_prompt") or architecture_views.get("last_ai_prompt") or ""
    merged["architecture_views"] = architecture_views
    merged["automation_assets"] = {
        "rules": _normalize_dict_rows(st.session_state.get("rules")),
        "proposed_rules": _normalize_dict_rows(st.session_state.get("proposed_rules")),
        "candidate_meta": _normalize_dict_rows(st.session_state.get("ai_meta")),
        "candidate_rows": _normalize_dict_rows(st.session_state.get("ai_candidates")),
        "instruction_intake": st.session_state.get("instruction_intake")
        if isinstance(st.session_state.get("instruction_intake"), dict)
        else None,
        "intent_spec": st.session_state.get("intent_spec") if isinstance(st.session_state.get("intent_spec"), dict) else None,
        "intent_spec_source": st.session_state.get("intent_spec_source"),
        "intent_spec_history": _normalize_dict_rows(st.session_state.get("intent_spec_history")),
    }
    return merged


def _prepare_runtime_from_semantic(save_rules_fn, rules_path: Path) -> dict:
    semantic_payload = _collect_semantic_payload_from_state(st.session_state.get("semantic_layer_spec") or {})
    st.session_state["semantic_layer_spec"] = semantic_payload
    assets = _get_semantic_assets(semantic_payload)
    runtime_rules = _normalize_dict_rows(assets.get("rules"))
    if runtime_rules:
        st.session_state["rules"] = runtime_rules
        save_rules_fn(rules_path, runtime_rules)
    runtime_spec = assets.get("intent_spec")
    if isinstance(runtime_spec, dict):
        st.session_state["intent_spec"] = runtime_spec
        return runtime_spec
    return st.session_state.get("intent_spec") if isinstance(st.session_state.get("intent_spec"), dict) else {}


def _runtime_readiness_issues(spec: dict) -> list[str]:
    issues = []
    purpose = spec.get("purpose") if isinstance(spec.get("purpose"), dict) else {}
    assets = _get_semantic_assets(spec)
    rules = _normalize_dict_rows(assets.get("rules"))
    intent_spec = assets.get("intent_spec")

    if not str(purpose.get("objective_statement") or "").strip():
        issues.append("Missing purpose.objective_statement in semantic layer.")
    if not str(purpose.get("priority_domain") or "").strip():
        issues.append("Missing purpose.priority_domain in semantic layer.")
    if not rules:
        issues.append("No semantic rules found in automation_assets.rules.")
    if not isinstance(intent_spec, dict) or not str(intent_spec.get("spec_id") or "").strip():
        issues.append("Missing semantic intent spec in automation_assets.intent_spec.")
    return issues


def _semantic_source_rows(spec: dict) -> list[dict]:
    purpose = spec.get("purpose") if isinstance(spec.get("purpose"), dict) else {}
    assets = _get_semantic_assets(spec)
    rules = _normalize_dict_rows(assets.get("rules"))
    proposed_rules = _normalize_dict_rows(assets.get("proposed_rules"))
    candidate_rows = _normalize_dict_rows(assets.get("candidate_rows"))
    intent_spec = assets.get("intent_spec")
    intent_history = _normalize_dict_rows(assets.get("intent_spec_history"))
    return [
        {
            "path": "purpose.objective_statement",
            "status": "ready" if str(purpose.get("objective_statement") or "").strip() else "missing",
            "value": str(purpose.get("objective_statement") or ""),
        },
        {
            "path": "purpose.priority_domain",
            "status": "ready" if str(purpose.get("priority_domain") or "").strip() else "missing",
            "value": str(purpose.get("priority_domain") or ""),
        },
        {
            "path": "automation_assets.rules",
            "status": "ready" if rules else "missing",
            "value": len(rules),
        },
        {
            "path": "automation_assets.proposed_rules",
            "status": "ready" if proposed_rules else "missing",
            "value": len(proposed_rules),
        },
        {
            "path": "automation_assets.candidate_rows",
            "status": "ready" if candidate_rows else "missing",
            "value": len(candidate_rows),
        },
        {
            "path": "automation_assets.intent_spec",
            "status": "ready"
            if isinstance(intent_spec, dict) and str(intent_spec.get("spec_id") or "").strip()
            else "missing",
            "value": intent_spec.get("spec_id") if isinstance(intent_spec, dict) else "",
        },
        {
            "path": "automation_assets.intent_spec_history",
            "status": "ready" if intent_history else "missing",
            "value": len(intent_history),
        },
    ]


def _render_input_hub_mapping(title: str, rows: list[dict]) -> None:
    st.markdown(f"<div class='psb-label'>{title}</div>", unsafe_allow_html=True)
    st.dataframe(rows, use_container_width=True, hide_index=True)


def _extract_instruction_text(record: dict) -> str:
    if not isinstance(record, dict):
        return ""
    parts = []
    raw = record.get("instruction_raw")
    summary = record.get("intent_summary")
    tasks = record.get("tasks")
    constraints = record.get("constraints")
    if isinstance(raw, str):
        parts.append(raw)
    if isinstance(summary, str):
        parts.append(summary)
    if isinstance(tasks, list):
        for item in tasks:
            if isinstance(item, str):
                parts.append(item)
            elif isinstance(item, dict):
                parts.append(str(item.get("task") or item.get("name") or item.get("action") or ""))
    if isinstance(constraints, list):
        for item in constraints:
            if isinstance(item, str):
                parts.append(item)
    return " ".join([part.strip() for part in parts if isinstance(part, str) and part.strip()]).lower()


def _fallback_rule_drafts_from_text(text: str) -> list[dict]:
    lowered = (text or "").lower()
    if not lowered:
        return []
    action_id = "save_process"
    if "ocr" in lowered:
        action_id = "ocr_process"
    elif "zip" in lowered or "unzip" in lowered:
        action_id = "unzip_process"
    subject_filter = "Invoice" if "invoice" in lowered else "Auto Draft"
    return [
        {
            "subject_filter": subject_filter,
            "task_name": "AUTO_DRAFT",
            "require_attachment": True,
            "target_ext": ".pdf",
            "action_id": action_id,
            "parameters": {"draft_source": "ai_instruction_fallback"},
            "rule_source": {"kind": "ai_draft", "confidence": 0.5},
        }
    ]


def _rule_signature(rule: dict) -> tuple[str, str, bool, str]:
    if not isinstance(rule, dict):
        return ("", "", False, "")
    target_ext = str(rule.get("target_ext") or rule.get("ext") or "")
    return (
        str(rule.get("subject_filter") or ""),
        str(rule.get("action_id") or ""),
        bool(rule.get("require_attachment")),
        target_ext,
    )


def _intent_rule_from_spec(spec: dict) -> dict | None:
    if not isinstance(spec, dict):
        return None
    inputs = spec.get("inputs")
    if not isinstance(inputs, dict):
        return None
    mail_rule = inputs.get("mail_rule")
    if not isinstance(mail_rule, dict):
        return None
    subject_filter = str(mail_rule.get("subject_filter") or "").strip()
    action_id = str(mail_rule.get("action_id") or "").strip()
    if not subject_filter or not action_id:
        return None
    return {
        "subject_filter": subject_filter,
        "action_id": action_id,
        "require_attachment": bool(mail_rule.get("require_attachment")),
        "target_ext": str(mail_rule.get("target_ext") or ""),
        "task_name": str(mail_rule.get("task_name") or "AUTO"),
        "parameters": mail_rule.get("parameters") if isinstance(mail_rule.get("parameters"), dict) else {},
    }


def _upsert_intent_history(spec: dict, source: str | None) -> None:
    if not isinstance(spec, dict):
        return
    spec_id = str(spec.get("spec_id") or "").strip()
    if not spec_id:
        return
    history = _normalize_dict_rows(st.session_state.get("intent_spec_history"))
    new_entry = {
        "spec_id": spec_id,
        "source": source or "unknown",
        "captured_at": datetime.utcnow().replace(microsecond=0).isoformat() + "Z",
        "domain": str(spec.get("domain") or ""),
        "intent": str(spec.get("intent") or ""),
        "rule": _intent_rule_from_spec(spec),
    }
    filtered = [entry for entry in history if str(entry.get("spec_id") or "") != spec_id]
    filtered.insert(0, new_entry)
    st.session_state["intent_spec_history"] = filtered[:30]


def _build_rule_ir_relationship_rows(
    active_rules: list[dict],
    proposed_rules: list[dict],
    intent_history: list[dict],
) -> list[dict]:
    intent_rule_map = {}
    for entry in _normalize_dict_rows(intent_history):
        rule = entry.get("rule")
        if not isinstance(rule, dict):
            continue
        signature = _rule_signature(rule)
        intent_rule_map.setdefault(signature, []).append(str(entry.get("spec_id") or ""))

    rows = []
    for source_label, rules in [("active_rule", active_rules), ("proposed_rule", proposed_rules)]:
        for index, rule in enumerate(_normalize_dict_rows(rules)):
            signature = _rule_signature(rule)
            linked_specs = intent_rule_map.get(signature, [])
            rows.append(
                {
                    "source": source_label,
                    "index": index,
                    "subject_filter": rule.get("subject_filter"),
                    "action_id": rule.get("action_id"),
                    "task_name": rule.get("task_name"),
                    "linked_ir_count": len(linked_specs),
                    "linked_ir_ids": ", ".join(linked_specs),
                    "status": "linked" if linked_specs else "unlinked",
                }
            )
    return rows


def _build_ir_coverage_rows(intent_history: list[dict], active_rules: list[dict], proposed_rules: list[dict]) -> list[dict]:
    active_signatures = {_rule_signature(rule) for rule in _normalize_dict_rows(active_rules)}
    proposed_signatures = {_rule_signature(rule) for rule in _normalize_dict_rows(proposed_rules)}
    rows = []
    for entry in _normalize_dict_rows(intent_history):
        rule = entry.get("rule")
        if not isinstance(rule, dict):
            continue
        signature = _rule_signature(rule)
        in_active = signature in active_signatures
        in_proposed = signature in proposed_signatures
        rows.append(
            {
                "spec_id": entry.get("spec_id"),
                "source": entry.get("source"),
                "subject_filter": rule.get("subject_filter"),
                "action_id": rule.get("action_id"),
                "active_rule": in_active,
                "proposed_rule": in_proposed,
                "coverage": "active" if in_active else ("proposed" if in_proposed else "missing"),
            }
        )
    return rows


def _queue_missing_ir_rules(intent_history: list[dict], existing_rules: list[dict], proposed_rules: list[dict]) -> tuple[list[dict], dict]:
    existing_signatures = {_rule_signature(rule) for rule in _normalize_dict_rows(existing_rules)}
    proposed_signatures = {_rule_signature(rule) for rule in _normalize_dict_rows(proposed_rules)}
    candidates = []
    for entry in _normalize_dict_rows(intent_history):
        rule = entry.get("rule")
        if not isinstance(rule, dict):
            continue
        signature = _rule_signature(rule)
        if signature in existing_signatures or signature in proposed_signatures:
            continue
        candidate = dict(rule)
        candidate["rule_source"] = {
            "kind": "intent_history",
            "spec_id": entry.get("spec_id"),
            "captured_at": entry.get("captured_at"),
        }
        candidates.append(candidate)
    return append_unique_rules(proposed_rules, candidates)


def _build_hub_node_rows(spec: dict) -> list[dict]:
    assets = _get_semantic_assets(spec)
    intent_spec = assets.get("intent_spec") if isinstance(assets.get("intent_spec"), dict) else {}
    return [
        {"node": "rules", "count": len(_normalize_dict_rows(assets.get("rules"))), "status": "ready" if _normalize_dict_rows(assets.get("rules")) else "missing"},
        {"node": "proposed_rules", "count": len(_normalize_dict_rows(assets.get("proposed_rules"))), "status": "ready" if _normalize_dict_rows(assets.get("proposed_rules")) else "missing"},
        {"node": "candidate_rows", "count": len(_normalize_dict_rows(assets.get("candidate_rows"))), "status": "ready" if _normalize_dict_rows(assets.get("candidate_rows")) else "missing"},
        {"node": "intent_spec", "count": 1 if intent_spec else 0, "status": "ready" if intent_spec else "missing"},
        {"node": "intent_history", "count": len(_normalize_dict_rows(assets.get("intent_spec_history"))), "status": "ready" if _normalize_dict_rows(assets.get("intent_spec_history")) else "missing"},
        {"node": "run_gateway", "count": 1, "status": "ready"},
    ]


def _build_hub_edge_rows(spec: dict) -> list[dict]:
    assets = _get_semantic_assets(spec)
    return [
        {"from": "input_rules", "to": "automation_assets.rules", "status": "ready" if _normalize_dict_rows(assets.get("rules")) else "missing"},
        {"from": "input_candidates", "to": "automation_assets.candidate_rows", "status": "ready" if _normalize_dict_rows(assets.get("candidate_rows")) else "missing"},
        {"from": "input_intent", "to": "automation_assets.intent_spec", "status": "ready" if isinstance(assets.get("intent_spec"), dict) else "missing"},
        {"from": "automation_assets.rules", "to": "run_automation", "status": "ready" if _normalize_dict_rows(assets.get("rules")) else "blocked"},
        {"from": "automation_assets.intent_spec", "to": "run_automation", "status": "ready" if isinstance(assets.get("intent_spec"), dict) else "blocked"},
    ]


def _semantic_mermaid_from_spec(spec: dict) -> str:
    assets = _get_semantic_assets(spec)
    rules_count = len(_normalize_dict_rows(assets.get("rules")))
    proposed_count = len(_normalize_dict_rows(assets.get("proposed_rules")))
    candidate_count = len(_normalize_dict_rows(assets.get("candidate_rows")))
    intent_spec = assets.get("intent_spec") if isinstance(assets.get("intent_spec"), dict) else {}
    intent_id = str(intent_spec.get("spec_id") or "missing")
    return "\n".join(
        [
            "flowchart TD",
            '  IRules["Input: Rules"] --> HubRules["Hub rules (' + str(rules_count) + ')"]',
            '  ICand["Input: Candidates"] --> HubCand["Hub candidates (' + str(candidate_count) + ')"]',
            '  IIntent["Input: Intent"] --> HubIntent["Hub intent (' + intent_id + ')"]',
            '  HubCand --> HubProposed["Hub proposed (' + str(proposed_count) + ')"]',
            '  HubRules --> Run["Run Automation"]',
            '  HubIntent --> Run',
        ]
    )


def _semantic_mermaid_from_intent_steps(spec: dict) -> str:
    if not isinstance(spec, dict):
        return ""
    steps = spec.get("steps")
    if not isinstance(steps, list) or not steps:
        return ""
    lines = ["flowchart TD", '  Start["AI Prompt"] --> S0["Intent Spec"]']
    for index, step in enumerate(steps):
        if not isinstance(step, dict):
            continue
        label = str(step.get("action") or step.get("id") or f"step_{index + 1}")
        lines.append(f'  S{index} --> S{index + 1}["{label}"]')
    lines.append(f'  S{len(steps)} --> End["Hub Draft"]')
    return "\n".join(lines)
