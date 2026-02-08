import json
import re
import time
from datetime import datetime, timedelta
from pathlib import Path

import streamlit as st

from src.bridge.builder import build_all_configs
from src.engine.core import GenericEtlEngine
from src.web.app_logic import (
    analyze_and_log_user_instruction,
    append_unique_rules,
    build_rule_proposals_from_intent_spec,
    compute_job_duration_seconds,
    generate_intent_spec,
    generate_intent_spec_from_summary,
    generate_followup_question,
    summarize_conversation,
    load_jsonl_runs,
    load_jsonl_runs_tail,
    load_rules,
    merge_proposed_rules,
    propose_rule_candidates,
    run_bridge_compile_summary,
    run_engine_job,
    save_rules,
    start_job,
    summarize_run_detail_rows,
    summarize_quality,
    summarize_run_window,
)


APP_TITLE = "Pseudo Semantic Bridge"
RULES_PATH = Path("configs/accounting/mail_business_rules.json")
PROPOSED_RULES_PATH = Path("configs/accounting/mail_rules_proposed.json")
SYSTEM_CONFIG_PATH = Path("configs/accounting/invoice_bot_v2.json")
LOGS_PATH = Path("data/logs/psb_run.jsonl")
INTAKE_LOGS_PATH = Path("data/logs/intent_intake.jsonl")
SEMANTIC_LAYER_PATH = Path("configs/accounting/semantic_layer_definition.json")
OUTLOOK_IMPORT_ERROR = None
OutlookAdapter = None

try:
    from src.adapter.outlook import OutlookAdapter as _OutlookAdapter
    OutlookAdapter = _OutlookAdapter
except Exception as exc:
    OUTLOOK_IMPORT_ERROR = exc


st.set_page_config(page_title=APP_TITLE, layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;600&family=JetBrains+Mono:wght@400;600&display=swap');
    :root {
        --bg: #0a1118;
        --bg-soft: #0f1722;
        --ink: #e6edf6;
        --accent: #22d3a6;
        --accent-2: #fb923c;
        --card: #152131;
        --line: #2b3a4f;
        --muted: #94a3b8;
    }
    html, body, [class*="stApp"] {
        font-family: "Space Grotesk", sans-serif;
        color: var(--ink);
        background: var(--bg);
    }
    [data-testid="stAppViewContainer"] {
        background:
            radial-gradient(circle at 85% -20%, #134e4a55 0%, transparent 35%),
            radial-gradient(circle at 10% -10%, #7c2d1250 0%, transparent 32%),
            var(--bg);
    }
    .psb-hero {
        background: linear-gradient(135deg, #0f1f2f 0%, #17263a 100%);
        padding: 24px 28px;
        border-radius: 16px;
        border: 1px solid var(--line);
    }
    .psb-title {
        font-size: 32px;
        font-weight: 600;
        margin: 0 0 8px 0;
    }
    .psb-sub {
        color: var(--muted);
        margin: 0;
    }
    .psb-card {
        background: var(--card);
        padding: 16px 18px;
        border-radius: 12px;
        border: 1px solid var(--line);
    }
    .psb-kpi {
        background: var(--card);
        border: 1px solid var(--line);
        border-radius: 12px;
        padding: 14px 16px;
    }
    .psb-kpi-label {
        color: var(--muted);
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.06em;
    }
    .psb-kpi-value {
        color: var(--ink);
        font-size: 28px;
        font-weight: 600;
        margin-top: 4px;
    }
    .psb-skeleton-card {
        background: linear-gradient(180deg, #172334 0%, #131d2c 100%);
        padding: 14px 16px;
        border-radius: 12px;
        border: 1px solid var(--line);
    }
    .psb-skeleton-title {
        margin: 0 0 10px 0;
        color: var(--ink);
        font-weight: 600;
        font-size: 14px;
    }
    .psb-skeleton-bar {
        height: 10px;
        border-radius: 999px;
        margin-bottom: 8px;
        background: linear-gradient(90deg, #233246 10%, #304661 40%, #233246 70%);
        background-size: 220% 100%;
        animation: shimmer 1.6s infinite linear;
    }
    @keyframes shimmer {
        0% { background-position: 200% 0; }
        100% { background-position: -20% 0; }
    }
    .psb-label {
        font-size: 12px;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: var(--muted);
        margin-bottom: 6px;
    }
    .psb-mono {
        font-family: "JetBrains Mono", monospace;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_runs_cached(path_str: str, mtime: float):
    return load_jsonl_runs(Path(path_str))


@st.cache_data(show_spinner=False)
def load_runs_tail_cached(path_str: str, mtime: float, max_lines: int):
    return load_jsonl_runs_tail(Path(path_str), max_lines=max_lines)


def _log_mtime(path: Path) -> float:
    try:
        return path.stat().st_mtime
    except FileNotFoundError:
        return 0.0


def _split_chunks(text: str) -> list[str]:
    cleaned = (text or "").strip()
    if not cleaned:
        return []
    cleaned = cleaned.replace("・", " ")
    parts = re.split(r"[、。/|\n]+", cleaned)
    chunks = [part.strip() for part in parts if part.strip()]
    return chunks if chunks else [cleaned]


def _dedupe_chunks(chunks: list[str]) -> list[str]:
    seen = set()
    unique = []
    for chunk in chunks:
        if chunk in seen:
            continue
        seen.add(chunk)
        unique.append(chunk)
    return unique


def _strip_select(rows: list[dict]) -> list[dict]:
    cleaned = []
    for row in rows or []:
        if not isinstance(row, dict):
            continue
        row_copy = dict(row)
        row_copy.pop("select", None)
        cleaned.append(row_copy)
    return cleaned


def _record_perf(label: str, seconds: float, count: int) -> None:
    marks = st.session_state.get("perf_marks") or []
    marks.append((label, seconds, count))
    if len(marks) > 200:
        marks = marks[-200:]
    st.session_state["perf_marks"] = marks


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


def _collect_semantic_payload_from_state(base_spec: dict) -> dict:
    merged = _merge_semantic_spec(_default_semantic_layer_spec(), base_spec if isinstance(base_spec, dict) else {})
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
    }
    return merged


def _prepare_runtime_from_semantic() -> dict:
    semantic_payload = _collect_semantic_payload_from_state(st.session_state.get("semantic_layer_spec") or {})
    st.session_state["semantic_layer_spec"] = semantic_payload
    assets = _get_semantic_assets(semantic_payload)
    runtime_rules = _normalize_dict_rows(assets.get("rules"))
    if runtime_rules:
        st.session_state["rules"] = runtime_rules
        save_rules(RULES_PATH, runtime_rules)
    runtime_spec = assets.get("intent_spec")
    if isinstance(runtime_spec, dict):
        st.session_state["intent_spec"] = runtime_spec
        return runtime_spec
    return st.session_state.get("intent_spec") if isinstance(st.session_state.get("intent_spec"), dict) else {}


def render_kpi(label: str, value: str):
    st.markdown(
        f"""
        <div class="psb-kpi">
            <div class="psb-kpi-label">{label}</div>
            <div class="psb-kpi-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_skeleton_card(title: str, widths: list[int]):
    bars = "".join([f'<div class="psb-skeleton-bar" style="width:{width}%"></div>' for width in widths])
    st.markdown(
        f"""
        <div class="psb-skeleton-card">
            <div class="psb-skeleton-title">{title}</div>
            {bars}
        </div>
        """,
        unsafe_allow_html=True,
    )


st.markdown(
    f"""
    <div class="psb-hero">
        <div class="psb-title">{APP_TITLE}</div>
        <p class="psb-sub">Dark-mode skeleton for planning Rules, Quality, and Pipeline flow.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.write("")

tabs = st.tabs(
    [
        "1) Semantic Overview",
        "2) Semantic Input: Rules",
        "3) Semantic Input: Candidates",
        "4) Semantic Input: Intent",
        "5) Semantic Layer Hub",
        "6) Run Automation",
    ]
)

if "perf_debug" not in st.session_state:
    st.session_state["perf_debug"] = False
if "perf_marks" not in st.session_state:
    st.session_state["perf_marks"] = []

if "rules" not in st.session_state:
    existing_rules = load_rules(RULES_PATH)
    st.session_state["rules"] = existing_rules or [
        {
            "subject_filter": "Invoice",
            "task_name": "INVOICE",
            "require_attachment": True,
            "target_ext": ".pdf",
            "action_id": "ocr_process",
            "parameters": {"lang": "jpn"},
        }
    ]

if "proposed_rules" not in st.session_state:
    st.session_state["proposed_rules"] = load_rules(PROPOSED_RULES_PATH)

if "ai_meta" not in st.session_state:
    st.session_state["ai_meta"] = []

if "ai_candidates" not in st.session_state:
    st.session_state["ai_candidates"] = []

if "intent_spec" not in st.session_state:
    st.session_state["intent_spec"] = None

if "intent_spec_source" not in st.session_state:
    st.session_state["intent_spec_source"] = None

if "intent_spec_error" not in st.session_state:
    st.session_state["intent_spec_error"] = None

if "instruction_intake" not in st.session_state:
    st.session_state["instruction_intake"] = None

if "compile_summary" not in st.session_state:
    st.session_state["compile_summary"] = None
if "semantic_layer_spec" not in st.session_state:
    st.session_state["semantic_layer_spec"] = _load_semantic_layer_spec(SEMANTIC_LAYER_PATH)
if "conversation_log" not in st.session_state:
    st.session_state["conversation_log"] = []
if "conversation_rounds" not in st.session_state:
    st.session_state["conversation_rounds"] = 0
if "conversation_allow_more" not in st.session_state:
    st.session_state["conversation_allow_more"] = False
if "conversation_summary" not in st.session_state:
    st.session_state["conversation_summary"] = []
if "conversation_focus" not in st.session_state:
    st.session_state["conversation_focus"] = None
if "conversation_marked" not in st.session_state:
    st.session_state["conversation_marked"] = []

_load_working_state_from_semantic(st.session_state.get("semantic_layer_spec") or {})
st.session_state["semantic_layer_spec"] = _collect_semantic_payload_from_state(
    st.session_state.get("semantic_layer_spec") or {}
)

with tabs[0]:
    st.caption("Workflow: Semantic Inputs -> Semantic Layer Hub -> Run Automation")
    log_mtime = _log_mtime(LOGS_PATH)
    start = time.perf_counter()
    runs = load_runs_cached(str(LOGS_PATH), log_mtime)
    _record_perf("overview.load_runs_full", time.perf_counter() - start, len(runs))
    summary = summarize_quality(runs) if runs else {"total": 0, "success": 0, "quality_labeled": 0, "quality_ok": 0}
    semantic_summary = _summarize_semantic_layer(st.session_state.get("semantic_layer_spec") or {})
    success_rate = round((summary["success"] / summary["total"]) * 100, 1) if summary["total"] else 0.0
    quality_rate = round((summary["quality_ok"] / summary["total"]) * 100, 1) if summary["total"] else 0.0

    st.markdown("<div class='psb-label'>Control Tower</div>", unsafe_allow_html=True)
    col_a, col_b, col_c, col_d, col_e, col_f = st.columns(6)
    with col_a:
        render_kpi("Rules", str(len(st.session_state["rules"])))
    with col_b:
        render_kpi("Runs", str(summary["total"]))
    with col_c:
        render_kpi("Success Rate", f"{success_rate}%")
    with col_d:
        render_kpi("Quality OK", f"{quality_rate}%")
    with col_e:
        render_kpi("Semantic Readiness", f"{semantic_summary['readiness_pct']}%")
    with col_f:
        render_kpi("Glossary Terms", str(semantic_summary["glossary_terms"]))

    st.write("")
    st.markdown("<div class='psb-label'>Planned Modules (Skeleton)</div>", unsafe_allow_html=True)
    col_g, col_h, col_i = st.columns(3)
    with col_g:
        render_skeleton_card("Intake Flow", [92, 75, 60, 84])
    with col_h:
        render_skeleton_card("Rule Suggestion", [88, 82, 68, 54])
    with col_i:
        render_skeleton_card("Run Timeline", [90, 65, 72, 58])

    st.caption(
        "Semantic progress: "
        f"{semantic_summary['ready_steps']}/{semantic_summary['total_steps']} steps complete | "
        f"metadata sources {semantic_summary['metadata_sources']} | "
        f"domain owners {semantic_summary['domain_owners']}"
    )

with tabs[1]:
    st.markdown("<div class='psb-label'>Semantic Input / Rule Assets</div>", unsafe_allow_html=True)
    st.caption("Writes into semantic path: automation_assets.rules / automation_assets.proposed_rules")
    edited = st.data_editor(
        st.session_state["rules"],
        num_rows="dynamic",
        use_container_width=True,
        key="rules_editor",
    )
    st.session_state["rules"] = edited

    col_a, col_b = st.columns([1, 1])
    with col_a:
        if st.button("Save Rules", type="primary"):
            save_rules(RULES_PATH, st.session_state["rules"])
            st.success("Rules saved.")
    with col_b:
        if st.button("Reload"):
            st.session_state["rules"] = load_rules(RULES_PATH)
            st.rerun()

    st.write("")
    st.markdown("<div class='psb-label'>Proposed Rules</div>", unsafe_allow_html=True)
    st.caption(f"Proposed rules path: {PROPOSED_RULES_PATH}")
    proposed_with_select = []
    for row in st.session_state.get("proposed_rules") or []:
        row_copy = dict(row)
        row_copy["select"] = False
        proposed_with_select.append(row_copy)
    edited_proposed = st.data_editor(
        proposed_with_select,
        num_rows="dynamic",
        use_container_width=True,
        key="proposed_rules_editor",
    )
    selected_indices = [idx for idx, row in enumerate(edited_proposed) if row.get("select")]
    st.session_state["proposed_rules"] = _strip_select(edited_proposed)

    merge_col_a, merge_col_b, merge_col_c = st.columns([1, 1, 1])
    with merge_col_a:
        merge_gate = st.number_input(
            "Merge quality gate",
            min_value=0.0,
            max_value=1.0,
            value=0.8,
            step=0.05,
        )
    with merge_col_b:
        allow_low_quality = st.checkbox("Allow below gate", value=False)
    with merge_col_c:
        drop_metadata = st.checkbox("Drop rule metadata", value=True)

    col_p1, col_p2, col_p3 = st.columns([1, 1, 1])
    with col_p1:
        if st.button("Save Proposed"):
            save_rules(PROPOSED_RULES_PATH, st.session_state["proposed_rules"])
            st.success("Proposed rules saved.")
    with col_p2:
        if st.button("Reload Proposed"):
            st.session_state["proposed_rules"] = load_rules(PROPOSED_RULES_PATH)
            st.rerun()
    with col_p3:
        if st.button("Merge Selected"):
            if not selected_indices:
                st.warning("No proposed rules selected.")
            else:
                merged_rules, remaining, summary = merge_proposed_rules(
                    existing_rules=st.session_state["rules"],
                    proposed_rules=st.session_state["proposed_rules"],
                    selected_indices=selected_indices,
                    min_quality_score_gate=float(merge_gate),
                    allow_low_quality=allow_low_quality,
                    drop_metadata=drop_metadata,
                )
                st.session_state["rules"] = merged_rules
                st.session_state["proposed_rules"] = remaining
                st.success(
                    "Merged "
                    f"{summary['merged']} rules "
                    f"(skipped duplicates {summary['skipped_duplicates']}, "
                    f"conflicts {summary['skipped_conflicts']}, "
                    f"quality gate {summary['skipped_quality_gate']}, "
                    f"invalid {summary['skipped_invalid']})."
                )
                st.rerun()

with tabs[2]:
    st.markdown("<div class='psb-label'>Semantic Input / Candidate Mining</div>", unsafe_allow_html=True)
    st.caption("Writes into semantic path: automation_assets.candidate_meta / automation_assets.candidate_rows")
    st.write("実行ログからルール候補を生成し、Rulesへ反映するための設計タブです。")
    st.caption("Output: executable rule rows (subject_filter, action_id, etc.)")

    col_a, col_b, col_c = st.columns([1, 1, 1])
    with col_a:
        lookback_days = st.number_input("Lookback (days)", min_value=1, max_value=365, value=30, step=1)
    with col_b:
        min_samples = st.number_input("Min samples", min_value=1, max_value=1000, value=5, step=1)
    with col_c:
        min_quality = st.number_input(
            "Min quality rate",
            min_value=0.0,
            max_value=1.0,
            value=0.8,
            step=0.05,
        )

    user_instruction = st.text_area(
        "User instruction",
        placeholder="例: 請求書はOCRを優先。日報は保存のみ。",
        height=80,
    )
    col_i1, col_i2 = st.columns([1, 1])
    with col_i1:
        domain_hint = st.text_input("Domain hint", value="accounting_mail_invoice")
    with col_i2:
        instruction_model = st.text_input("Instruction model", value="gpt-4o-mini")

    col_i3, col_i4 = st.columns([1, 1])
    with col_i3:
        if st.button("Instruction 解析 (AI)"):
            analyzed, error, source = analyze_and_log_user_instruction(
                user_instruction=user_instruction,
                log_path=INTAKE_LOGS_PATH,
                domain_hint=domain_hint,
                use_ai=True,
                model=instruction_model,
            )
            st.session_state["instruction_intake"] = analyzed
            if error:
                st.warning(error)
            else:
                st.success("Instruction analyzed by AI and logged.")
    with col_i4:
        if st.button("Instruction 解析 (Template)"):
            analyzed, error, source = analyze_and_log_user_instruction(
                user_instruction=user_instruction,
                log_path=INTAKE_LOGS_PATH,
                domain_hint=domain_hint,
                use_ai=False,
            )
            st.session_state["instruction_intake"] = analyzed
            if error:
                st.warning(error)
            else:
                st.success("Instruction analyzed by template and logged.")

    st.caption(f"Instruction intake log: {INTAKE_LOGS_PATH}")
    if st.session_state["instruction_intake"]:
        st.markdown("<div class='psb-label'>Instruction Analysis</div>", unsafe_allow_html=True)
        st.json(st.session_state["instruction_intake"])

    if st.button("Generate Candidates", type="primary"):
        log_mtime = _log_mtime(LOGS_PATH)
        start = time.perf_counter()
        runs = load_runs_cached(str(LOGS_PATH), log_mtime)
        _record_perf("rule_builder.load_runs_full", time.perf_counter() - start, len(runs))
        if not runs:
            st.session_state["ai_meta"] = []
            st.session_state["ai_candidates"] = []
            st.info(f"No logs found at {LOGS_PATH}.")
        else:
            cutoff = datetime.utcnow() - timedelta(days=int(lookback_days))
            filtered = []
            for run in runs:
                ts = run.get("timestamp")
                if not ts:
                    filtered.append(run)
                    continue
                try:
                    ts_clean = str(ts).replace("Z", "+00:00")
                    parsed = datetime.fromisoformat(ts_clean)
                    if parsed >= cutoff:
                        filtered.append(run)
                except ValueError:
                    filtered.append(run)

            meta_rows, candidate_rows = propose_rule_candidates(
                filtered,
                min_samples=int(min_samples),
                min_quality_rate=float(min_quality),
            )
            st.session_state["ai_meta"] = meta_rows
            st.session_state["ai_candidates"] = candidate_rows
            st.session_state["ai_instruction"] = user_instruction

    log_mtime = _log_mtime(LOGS_PATH)
    start = time.perf_counter()
    runs = load_runs_cached(str(LOGS_PATH), log_mtime)
    _record_perf("rule_builder.summary_runs_full", time.perf_counter() - start, len(runs))
    summary = summarize_quality(runs) if runs else {"total": 0, "success": 0, "quality_labeled": 0, "quality_ok": 0}
    st.caption(
        f"Logs: {summary['total']} | Success: {summary['success']} | "
        f"Quality labeled: {summary['quality_labeled']} | Quality OK: {summary['quality_ok']}"
    )
    st.caption(f"Log path: {LOGS_PATH}")

    if st.session_state["ai_meta"]:
        st.markdown("<div class='psb-label'>Candidate Meta</div>", unsafe_allow_html=True)
        st.dataframe(st.session_state["ai_meta"], use_container_width=True)
    else:
        render_skeleton_card("Candidate Meta Table", [96, 84, 74])

    if st.session_state["ai_candidates"]:
        st.markdown("<div class='psb-label'>Candidate Rows (Excel compatible)</div>", unsafe_allow_html=True)
        candidates_with_select = []
        for row in st.session_state["ai_candidates"]:
            row_copy = dict(row)
            row_copy["select"] = False
            candidates_with_select.append(row_copy)

        edited_candidates = st.data_editor(
            candidates_with_select,
            num_rows="dynamic",
            use_container_width=True,
            key="ai_candidate_editor",
        )

        if st.button("Append Selected To Rules"):
            selected = [
                {
                    "subject_filter": row.get("subject_filter"),
                    "task_name": row.get("task_name"),
                    "require_attachment": row.get("require_attachment"),
                    "target_ext": row.get("target_ext"),
                    "action_id": row.get("action_id"),
                    "parameters": row.get("parameters") or {},
                }
                for row in edited_candidates
                if row.get("select")
            ]
            if not selected:
                st.warning("No rows selected.")
            else:
                st.session_state["rules"].extend(selected)
                st.success(f"Appended {len(selected)} rows to rules (draft).")
                st.rerun()
    else:
        render_skeleton_card("Candidate Rows Table", [92, 66, 82, 54])

with tabs[3]:
    st.markdown("<div class='psb-label'>Semantic Input / Intent Specification</div>", unsafe_allow_html=True)
    st.caption("Writes into semantic path: automation_assets.intent_spec / automation_assets.intent_spec_source")
    st.write("曖昧な要求を Intent Spec (IR) に定式化し、Run へ渡すための設計タブです。")
    st.caption("Output: intent spec JSON (spec_id, steps, verification, fallback)")

    llm_model = st.session_state.get("intent_llm_model", "gpt-4o-mini")

    mail_subject_filter = st.session_state.get("mail_subject_filter", "請求書")
    mail_task_name = st.session_state.get("mail_task_name", "INVOICE")
    mail_require_attachment = st.session_state.get("mail_require_attachment", True)
    mail_action_id = st.session_state.get("mail_action_id", "ocr_process")
    mail_params = st.session_state.get("mail_params", {"lang": "jpn"})
    if not isinstance(mail_params, dict):
        mail_params = {}

    st.markdown("<div class='psb-label'>Minimal Flow</div>", unsafe_allow_html=True)
    st.markdown("<div class='psb-label'>Conversation Assist</div>", unsafe_allow_html=True)
    if st.session_state["conversation_log"]:
        for entry in st.session_state["conversation_log"]:
            role = entry.get("role", "user")
            content = entry.get("content", "")
            with st.chat_message(role):
                st.markdown(content)
    else:
        st.caption("Conversation log is empty.")

    user_message = st.chat_input("やりたいことを入力してください")
    if user_message:
        st.session_state["conversation_log"].append({"role": "user", "content": user_message})
        allow_more_needed = st.session_state["conversation_rounds"] >= 3 and not st.session_state["conversation_allow_more"]
        if allow_more_needed:
            st.warning("質問は3回まで。さらに必要なら許可してください。")
        else:
            question, error, source = generate_followup_question(
                conversation=st.session_state["conversation_log"],
                domain_hint="accounting_mail_invoice",
                use_ai=True,
                model=llm_model,
                round_index=st.session_state["conversation_rounds"],
                memory_bullets=st.session_state.get("conversation_summary") or [],
            )
            if error:
                st.warning(error)
            st.session_state["conversation_log"].append({"role": "assistant", "content": question})
            st.session_state["conversation_rounds"] += 1
        st.rerun()

    convo_col_a, convo_col_b = st.columns([1, 1])
    with convo_col_a:
        if st.button("Allow More Questions"):
            st.session_state["conversation_allow_more"] = True
            st.rerun()
    with convo_col_b:
        if st.button("Reset Conversation"):
            st.session_state["conversation_log"] = []
            st.session_state["conversation_rounds"] = 0
            st.session_state["conversation_allow_more"] = False
            st.session_state["conversation_summary"] = []
            st.session_state["conversation_focus"] = None
            st.session_state["conversation_marked"] = []
            st.rerun()

    if st.button("Summarize Conversation"):
        summary, error, source = summarize_conversation(
            conversation=st.session_state["conversation_log"],
            use_ai=True,
            model=llm_model,
        )
        if error:
            st.warning(error)
        st.session_state["conversation_summary"] = summary
        st.session_state["conversation_focus"] = None
        st.session_state["conversation_marked"] = []

    summary_bullets = st.session_state.get("conversation_summary") or []
    if summary_bullets:
        st.markdown("<div class='psb-label'>Summary (Bullet)</div>", unsafe_allow_html=True)
        for bullet in summary_bullets:
            st.write(f"- {bullet}")
        chunks = []
        for bullet in summary_bullets:
            chunks.extend(_split_chunks(bullet))
        chunks = _dedupe_chunks(chunks)
        marked = st.multiselect(
            "興味のある文字の塊をマーキング",
            chunks,
            default=st.session_state.get("conversation_marked") or [],
        )
        st.session_state["conversation_marked"] = marked
        if st.button("Use Marked Chunks as Focus"):
            if not marked:
                st.warning("Marked chunks are empty.")
            else:
                st.session_state["conversation_focus"] = " / ".join(marked)
        focus_choice = st.selectbox("興味のあるポイントを選択", summary_bullets)
        if st.button("Confirm Focus"):
            st.session_state["conversation_focus"] = focus_choice

    if st.button("Intent Spec 生成 (Conversation)", type="primary"):
        focus = st.session_state.get("conversation_focus")
        if not focus:
            st.warning("Focus is not confirmed yet.")
        else:
            spec, error, source = generate_intent_spec_from_summary(
                summary_bullets=st.session_state.get("conversation_summary") or [],
                focus=focus,
                use_ai=True,
                model=llm_model,
            )
            spec_inputs = spec.get("inputs") or {}
            spec_inputs["mail_rule"] = {
                "subject_filter": mail_subject_filter,
                "task_name": mail_task_name,
                "require_attachment": mail_require_attachment,
                "action_id": mail_action_id,
                "parameters": mail_params,
            }
            spec_inputs["semantic_layer"] = _build_semantic_context(st.session_state.get("semantic_layer_spec") or {})
            spec["inputs"] = spec_inputs
            st.session_state["intent_spec"] = spec
            st.session_state["intent_spec_source"] = source
            st.session_state["intent_spec_error"] = error
            if error:
                st.warning(error)
            else:
                st.success("Intent Spec generated from conversation.")

    with st.expander("Advanced Inputs", expanded=False):
        llm_model = st.text_input("OpenAI model", value=llm_model)
        st.session_state["intent_llm_model"] = llm_model
        col_a, col_b = st.columns([2, 1])
        with col_a:
            app_context = st.text_input(
                "対象アプリ/領域",
                value="メール",
                help="例: メール, 受発注, 請求書管理",
            )
            goal = st.text_area(
                "何がしたい？（目的）",
                placeholder="例: 添付の写真から文字抽出して、処理フローに組み込みたい",
                height=100,
            )
            scope = st.text_area(
                "想定シナリオ/制約",
                placeholder="例: まずは単体で出来栄えを見たい。機密情報あり。",
                height=100,
            )
            success = st.text_area(
                "成功条件/評価基準",
                placeholder="例: 95%の抽出精度、3秒以内の処理",
                height=80,
            )
        with col_b:
            st.markdown("<div class='psb-label'>進め方</div>", unsafe_allow_html=True)
            path = st.radio(
                "どの形で検証する？",
                ["まずは単体の出来栄えを見る", "ワークフローに組み込みたい"],
            )
            customization = st.radio(
                "ユーザーが自分でやってよい範囲",
                ["簡単な開発/カスタムはユーザーに任せる", "基本は運用チームで対応"],
            )
            st.markdown("<div class='psb-label' style='margin-top:12px'>任意情報</div>", unsafe_allow_html=True)
            artifacts = st.text_area(
                "参考情報（任意）",
                placeholder="例: 既存ルール、ログ、サンプル画像の説明",
                height=120,
            )

        col_c, col_d = st.columns([1, 1])
        with col_c:
            if st.button("受付内容を保存"):
                st.session_state["quality_intake"] = {
                    "app_context": app_context,
                    "goal": goal,
                    "scope": scope,
                    "success": success,
                    "path": path,
                    "customization": customization,
                    "artifacts": artifacts,
                }
                st.success("受付内容を保存しました。")
        with col_d:
            if st.button("Intent Spec 生成 (AI)"):
                spec, error, source = generate_intent_spec(
                    app_context=app_context,
                    goal=goal,
                    scope=scope,
                    success=success,
                    artifacts=artifacts,
                    use_ai=True,
                    model=llm_model,
                )
                spec_inputs = spec.get("inputs") or {}
                spec_inputs["mail_rule"] = {
                    "subject_filter": mail_subject_filter,
                    "task_name": mail_task_name,
                    "require_attachment": mail_require_attachment,
                    "action_id": mail_action_id,
                    "parameters": mail_params,
                }
                spec_inputs["semantic_layer"] = _build_semantic_context(st.session_state.get("semantic_layer_spec") or {})
                spec["inputs"] = spec_inputs
                st.session_state["intent_spec"] = spec
                st.session_state["intent_spec_source"] = source
                st.session_state["intent_spec_error"] = error
                if error:
                    st.warning(error)
                else:
                    st.success("Intent Spec generated by AI.")

        if st.button("Intent Spec 生成 (Template)"):
            spec, error, source = generate_intent_spec(
                app_context=app_context,
                goal=goal,
                scope=scope,
                success=success,
                artifacts=artifacts,
                use_ai=False,
            )
            spec_inputs = spec.get("inputs") or {}
            spec_inputs["mail_rule"] = {
                "subject_filter": mail_subject_filter,
                "task_name": mail_task_name,
                "require_attachment": mail_require_attachment,
                "action_id": mail_action_id,
                "parameters": mail_params,
            }
            spec_inputs["semantic_layer"] = _build_semantic_context(st.session_state.get("semantic_layer_spec") or {})
            spec["inputs"] = spec_inputs
            st.session_state["intent_spec"] = spec
            st.session_state["intent_spec_source"] = source
            st.session_state["intent_spec_error"] = error
            st.success("Intent Spec generated from template.")

    with st.expander("Mail Rule Mapping (Optional)", expanded=False):
        rule_col_a, rule_col_b, rule_col_c = st.columns([1, 1, 1])
        with rule_col_a:
            mail_subject_filter = st.text_input("Mail subject filter", value=mail_subject_filter)
        with rule_col_b:
            mail_task_name = st.text_input("Mail task name", value=mail_task_name)
        with rule_col_c:
            mail_require_attachment = st.checkbox("Require attachment", value=mail_require_attachment)
        rule_col_d, rule_col_e = st.columns([1, 1])
        with rule_col_d:
            mail_action_id = st.selectbox(
                "Action (mail rule)",
                ["ocr_process", "save_process", "unzip_process"],
                index=["ocr_process", "save_process", "unzip_process"].index(mail_action_id)
                if mail_action_id in {"ocr_process", "save_process", "unzip_process"}
                else 0,
            )
        with rule_col_e:
            mail_params_raw = st.text_area(
                "Action parameters (JSON)",
                value=json.dumps(mail_params or {}, ensure_ascii=False),
                height=80,
            )

        mail_params = {}
        if mail_params_raw.strip():
            try:
                mail_params = json.loads(mail_params_raw)
            except json.JSONDecodeError:
                st.warning("Action parameters JSON is invalid. Using empty parameters.")
                mail_params = {}

        st.session_state["mail_subject_filter"] = mail_subject_filter
        st.session_state["mail_task_name"] = mail_task_name
        st.session_state["mail_require_attachment"] = mail_require_attachment
        st.session_state["mail_action_id"] = mail_action_id
        st.session_state["mail_params"] = mail_params

        proposal_gate = st.number_input(
            "Proposal quality gate",
            min_value=0.0,
            max_value=1.0,
            value=0.8,
            step=0.05,
            key="proposal_quality_gate",
        )
        if st.button("Generate Proposed Rule", type="primary"):
            current_spec = st.session_state.get("intent_spec")
            if not current_spec:
                st.warning("Intent Spec not generated yet.")
            else:
                proposals, warnings = build_rule_proposals_from_intent_spec(
                    current_spec,
                    min_quality_score_gate=float(proposal_gate),
                )
                for warning in warnings:
                    st.warning(warning)
                if not proposals:
                    st.warning("No rule proposals generated from Intent Spec.")
                else:
                    merged, summary = append_unique_rules(
                        st.session_state.get("proposed_rules") or [],
                        proposals,
                    )
                    st.session_state["proposed_rules"] = merged
                    save_rules(PROPOSED_RULES_PATH, merged)
                    st.success(
                        f"Added {summary['added']} proposed rules "
                        f"(skipped duplicates {summary['skipped_duplicates']}, "
                        f"invalid {summary['skipped_invalid']})."
                    )
                    st.rerun()

    st.write("")
    st.markdown("<div class='psb-label'>Intent Specification (IR)</div>", unsafe_allow_html=True)
    if st.session_state["intent_spec"]:
        st.caption(
            "source: "
            f"{st.session_state.get('intent_spec_source', 'unknown')} | "
            f"spec_id: {st.session_state['intent_spec'].get('spec_id', '-')}"
        )
        st.json(st.session_state["intent_spec"])
    else:
        render_skeleton_card("Intent Spec Preview", [92, 78, 70, 86])

    if st.session_state.get("quality_intake"):
        st.json(st.session_state["quality_intake"])

with tabs[4]:
    st.markdown("<div class='psb-label'>Semantic Layer Hub</div>", unsafe_allow_html=True)
    st.write(
        "Central hub. This consolidates all semantic inputs with business meaning, lineage, ownership, "
        "and active metadata into one portable spec."
    )
    st.caption(f"Output path: {SEMANTIC_LAYER_PATH}")

    semantic_spec = _collect_semantic_payload_from_state(st.session_state.get("semantic_layer_spec") or {})

    purpose = semantic_spec.get("purpose") if isinstance(semantic_spec.get("purpose"), dict) else {}
    technical_metadata = (
        semantic_spec.get("technical_metadata") if isinstance(semantic_spec.get("technical_metadata"), dict) else {}
    )
    business_semantics = (
        semantic_spec.get("business_semantics") if isinstance(semantic_spec.get("business_semantics"), dict) else {}
    )
    federation = semantic_spec.get("federation") if isinstance(semantic_spec.get("federation"), dict) else {}
    active_metadata = (
        semantic_spec.get("active_metadata") if isinstance(semantic_spec.get("active_metadata"), dict) else {}
    )
    ownership = semantic_spec.get("ownership") if isinstance(semantic_spec.get("ownership"), dict) else {}

    st.markdown("<div class='psb-label'>Step 1: Business Objective and Scope</div>", unsafe_allow_html=True)
    objective_statement = st.text_area(
        "Objective statement",
        value=str(purpose.get("objective_statement") or ""),
        placeholder="e.g. Improve customer retention while reducing compliance risk.",
        height=80,
    )
    metric_col_a, metric_col_b = st.columns([1, 1])
    with metric_col_a:
        success_metric = st.text_input("Primary success metric", value=str(purpose.get("success_metric") or ""))
    with metric_col_b:
        priority_domain = st.text_input("Priority domain", value=str(purpose.get("priority_domain") or ""))
    initial_scope = st.text_area(
        "Initial high-value scope",
        value=str(purpose.get("initial_scope") or ""),
        placeholder="e.g. Start with finance and customer billing events only.",
        height=70,
    )
    priority_objectives = purpose.get("priority_objectives") if isinstance(purpose.get("priority_objectives"), list) else []
    priority_objectives = st.data_editor(
        priority_objectives,
        num_rows="dynamic",
        use_container_width=True,
        key="semantic_priority_objectives_editor",
    )

    st.markdown("<div class='psb-label'>Step 2: Automated Technical Metadata</div>", unsafe_allow_html=True)
    auto_collection_enabled = st.checkbox(
        "Enable automated metadata collection",
        value=bool(technical_metadata.get("auto_collection_enabled", True)),
    )
    metadata_sources = technical_metadata.get("metadata_sources") if isinstance(technical_metadata.get("metadata_sources"), list) else []
    metadata_sources = st.data_editor(
        metadata_sources,
        num_rows="dynamic",
        use_container_width=True,
        key="semantic_metadata_sources_editor",
    )
    lineage_paths = technical_metadata.get("lineage_paths") if isinstance(technical_metadata.get("lineage_paths"), list) else []
    lineage_paths = st.data_editor(
        lineage_paths,
        num_rows="dynamic",
        use_container_width=True,
        key="semantic_lineage_paths_editor",
    )

    st.markdown("<div class='psb-label'>Step 3: Business Terms and KPI Logic</div>", unsafe_allow_html=True)
    glossary_terms = business_semantics.get("glossary_terms") if isinstance(business_semantics.get("glossary_terms"), list) else []
    glossary_terms = st.data_editor(
        glossary_terms,
        num_rows="dynamic",
        use_container_width=True,
        key="semantic_glossary_editor",
    )
    kpi_definitions = business_semantics.get("kpi_definitions") if isinstance(business_semantics.get("kpi_definitions"), list) else []
    kpi_definitions = st.data_editor(
        kpi_definitions,
        num_rows="dynamic",
        use_container_width=True,
        key="semantic_kpi_editor",
    )

    st.markdown("<div class='psb-label'>Step 4: Federation of Existing Investments</div>", unsafe_allow_html=True)
    integrated_tools = federation.get("integrated_tools") if isinstance(federation.get("integrated_tools"), list) else []
    integrated_tools = st.data_editor(
        integrated_tools,
        num_rows="dynamic",
        use_container_width=True,
        key="semantic_integrated_tools_editor",
    )
    tacit_patterns = federation.get("tacit_patterns") if isinstance(federation.get("tacit_patterns"), list) else []
    tacit_patterns = st.data_editor(
        tacit_patterns,
        num_rows="dynamic",
        use_container_width=True,
        key="semantic_tacit_patterns_editor",
    )

    st.markdown("<div class='psb-label'>Step 5: Active Metadata and Learning Loop</div>", unsafe_allow_html=True)
    active_col_a, active_col_b, active_col_c = st.columns([1, 1, 1])
    with active_col_a:
        ai_enrichment_enabled = st.checkbox(
            "Enable AI enrichment",
            value=bool(active_metadata.get("ai_enrichment_enabled", True)),
        )
    with active_col_b:
        human_review_required = st.checkbox(
            "Require human review",
            value=bool(active_metadata.get("human_review_required", True)),
        )
    with active_col_c:
        learning_cycle_options = ["daily", "weekly", "biweekly", "monthly"]
        learning_cycle_current = str(active_metadata.get("learning_cycle") or "weekly")
        learning_cycle = st.selectbox(
            "Learning cycle",
            learning_cycle_options,
            index=learning_cycle_options.index(learning_cycle_current)
            if learning_cycle_current in learning_cycle_options
            else 1,
        )
    learning_signals = active_metadata.get("learning_signals") if isinstance(active_metadata.get("learning_signals"), list) else []
    learning_signals = st.data_editor(
        learning_signals,
        num_rows="dynamic",
        use_container_width=True,
        key="semantic_learning_signals_editor",
    )

    st.markdown("<div class='psb-label'>Step 6: Decentralized Ownership</div>", unsafe_allow_html=True)
    ownership_col_a, ownership_col_b = st.columns([1, 1])
    ownership_model_options = ["federated", "hybrid", "centralized"]
    ownership_model_current = str(ownership.get("ownership_model") or "federated")
    with ownership_col_a:
        ownership_model = st.selectbox(
            "Ownership model",
            ownership_model_options,
            index=ownership_model_options.index(ownership_model_current)
            if ownership_model_current in ownership_model_options
            else 0,
        )
    with ownership_col_b:
        central_team = st.text_input("Central team", value=str(ownership.get("central_team") or ""))
    domain_owners = ownership.get("domain_owners") if isinstance(ownership.get("domain_owners"), list) else []
    domain_owners = st.data_editor(
        domain_owners,
        num_rows="dynamic",
        use_container_width=True,
        key="semantic_domain_owners_editor",
    )
    guardrails = st.text_area(
        "Governance guardrails",
        value=str(ownership.get("guardrails") or ""),
        placeholder="e.g. Naming standards, SLA for approvals, and compliance checks.",
        height=70,
    )

    semantic_payload = {
        "spec_id": str(semantic_spec.get("spec_id") or "semantic-layer-blueprint"),
        "spec_version": str(semantic_spec.get("spec_version") or "1.0"),
        "updated_at": semantic_spec.get("updated_at"),
        "purpose": {
            "objective_statement": objective_statement,
            "success_metric": success_metric,
            "priority_domain": priority_domain,
            "initial_scope": initial_scope,
            "priority_objectives": [row for row in priority_objectives if isinstance(row, dict)],
        },
        "technical_metadata": {
            "auto_collection_enabled": bool(auto_collection_enabled),
            "metadata_sources": [row for row in metadata_sources if isinstance(row, dict)],
            "lineage_paths": [row for row in lineage_paths if isinstance(row, dict)],
        },
        "business_semantics": {
            "glossary_terms": [row for row in glossary_terms if isinstance(row, dict)],
            "kpi_definitions": [row for row in kpi_definitions if isinstance(row, dict)],
        },
        "federation": {
            "integrated_tools": [row for row in integrated_tools if isinstance(row, dict)],
            "tacit_patterns": [row for row in tacit_patterns if isinstance(row, dict)],
        },
        "active_metadata": {
            "ai_enrichment_enabled": bool(ai_enrichment_enabled),
            "human_review_required": bool(human_review_required),
            "learning_cycle": learning_cycle,
            "learning_signals": [row for row in learning_signals if isinstance(row, dict)],
        },
        "ownership": {
            "ownership_model": ownership_model,
            "central_team": central_team,
            "domain_owners": [row for row in domain_owners if isinstance(row, dict)],
            "guardrails": guardrails,
        },
    }
    semantic_payload = _collect_semantic_payload_from_state(semantic_payload)
    st.session_state["semantic_layer_spec"] = semantic_payload
    semantic_summary = _summarize_semantic_layer(semantic_payload)

    status_col_a, status_col_b, status_col_c = st.columns(3)
    with status_col_a:
        st.metric("Semantic readiness", f"{semantic_summary['readiness_pct']}%")
    with status_col_b:
        st.metric("Glossary terms", semantic_summary["glossary_terms"])
    with status_col_c:
        st.metric("Domain owners", semantic_summary["domain_owners"])

    st.caption(
        f"Step completion: {semantic_summary['ready_steps']}/{semantic_summary['total_steps']} | "
        f"metadata sources: {semantic_summary['metadata_sources']}"
    )
    if semantic_summary["ready_steps"] < semantic_summary["total_steps"]:
        st.warning("Some semantic-layer steps are still incomplete.")

    save_col_a, save_col_b, save_col_c = st.columns([1, 1, 1])
    with save_col_a:
        if st.button("Save Semantic Layer Definition", type="primary"):
            _save_semantic_layer_spec(SEMANTIC_LAYER_PATH, semantic_payload)
            st.session_state["semantic_layer_spec"] = _load_semantic_layer_spec(SEMANTIC_LAYER_PATH)
            st.success(f"Semantic layer definition saved: {SEMANTIC_LAYER_PATH}")
    with save_col_b:
        if st.button("Reload Semantic Layer Definition"):
            st.session_state["semantic_layer_spec"] = _load_semantic_layer_spec(SEMANTIC_LAYER_PATH)
            _load_working_state_from_semantic(st.session_state.get("semantic_layer_spec") or {})
            st.rerun()
    with save_col_c:
        if st.button("Apply Semantic Assets To Tabs"):
            _load_working_state_from_semantic(st.session_state.get("semantic_layer_spec") or {})
            st.success("Semantic assets applied to working tabs.")

    st.markdown("<div class='psb-label'>Semantic Layer Spec Preview</div>", unsafe_allow_html=True)
    st.json(semantic_payload)

with tabs[5]:
    jobs = st.session_state.setdefault("jobs", {})
    st.markdown("<div class='psb-label'>Run Automation</div>", unsafe_allow_html=True)
    st.caption("Run uses the semantic-layer aggregated assets as the final automation input.")
    st.markdown(
        "<div class='psb-card'>"
        "<div class='psb-label'>Config</div>"
        f"<div class='psb-mono'>{SYSTEM_CONFIG_PATH}</div>"
        "<div class='psb-label' style='margin-top:12px'>Rules</div>"
        f"<div class='psb-mono'>{RULES_PATH}</div>"
        "</div>",
        unsafe_allow_html=True,
    )

    st.write("")
    run_col_a, run_col_b = st.columns([1, 1])
    with run_col_a:
        if st.button("Compile Specs Only"):
            _prepare_runtime_from_semantic()
            compile_summary = run_bridge_compile_summary(
                build_all_configs,
                SYSTEM_CONFIG_PATH,
                RULES_PATH,
            )
            st.session_state["compile_summary"] = compile_summary
            if compile_summary["status"] == "done":
                st.success("Bridge compile completed.")
            else:
                st.error(f"Bridge compile failed: {compile_summary.get('error')}")
    with run_col_b:
        if st.button("Run Pipeline"):
            if OutlookAdapter is None:
                st.error(f"Outlook adapter is unavailable: {OUTLOOK_IMPORT_ERROR}")
                st.stop()
            current_spec = _prepare_runtime_from_semantic()
            log_mtime = _log_mtime(LOGS_PATH)
            start = time.perf_counter()
            baseline_count = len(load_runs_cached(str(LOGS_PATH), log_mtime))
            _record_perf("run.baseline_count", time.perf_counter() - start, baseline_count)

            def _run(current_job_id: str):
                run_engine_job(
                    jobs,
                    current_job_id,
                    build_all_configs,
                    SYSTEM_CONFIG_PATH,
                    OutlookAdapter,
                    GenericEtlEngine,
                )

            job_id = start_job(
                jobs,
                _run,
            )
            if current_spec.get("spec_id"):
                jobs[job_id]["spec_id"] = current_spec["spec_id"]
                jobs[job_id]["spec_source"] = st.session_state.get("intent_spec_source")
            jobs[job_id]["log_start_index"] = baseline_count
            st.session_state["last_job_id"] = job_id

    if st.button("Refresh Results"):
        st.rerun()

    st.write("")
    current_spec = st.session_state.get("intent_spec") or {}
    if current_spec.get("spec_id"):
        st.caption(
            f"Current spec: {current_spec['spec_id']} "
            f"(source: {st.session_state.get('intent_spec_source', '-')})"
        )
    else:
        st.caption("Current spec: not generated yet")

    compile_summary = st.session_state.get("compile_summary")
    if compile_summary:
        st.markdown("<div class='psb-label'>Last Compile Summary</div>", unsafe_allow_html=True)
        st.caption(
            f"status: {compile_summary.get('status')} | "
            f"started_at: {compile_summary.get('started_at')} | "
            f"ended_at: {compile_summary.get('ended_at')}"
        )
        st.dataframe(compile_summary.get("artifacts", []), use_container_width=True)
        if compile_summary.get("error"):
            st.warning(f"Compile error: {compile_summary.get('error')}")

    # Always show global log view so notebook-triggered runs are visible in Web.
    summary_col, detail_col = st.columns([1, 1])
    with summary_col:
        tail_limit = st.number_input("Global log lines (tail)", min_value=50, max_value=2000, value=300, step=50)
    with detail_col:
        show_global_detail = st.checkbox("Show global detail table", value=False)

    log_mtime = _log_mtime(LOGS_PATH)
    start = time.perf_counter()
    all_runs = load_runs_tail_cached(str(LOGS_PATH), log_mtime, max_lines=int(tail_limit))
    _record_perf("run.load_runs_tail", time.perf_counter() - start, len(all_runs))
    global_summary = summarize_run_window(all_runs, start_index=0)
    st.markdown("<div class='psb-label'>Global Run Summary (Jupyter + Web)</div>", unsafe_allow_html=True)
    g1, g2, g3, g4 = st.columns(4)
    with g1:
        st.metric("Processed", global_summary["total"])
    with g2:
        st.metric("Success", global_summary["success"])
    with g3:
        st.metric("Error", global_summary["error"])
    with g4:
        st.metric("Artifacts", global_summary["with_output"])
    if global_summary["workflows"]:
        st.caption(f"Workflows: {', '.join(global_summary['workflows'])}")
    if global_summary["latest_timestamp"]:
        st.caption(f"Latest log time: {global_summary['latest_timestamp']}")
    if global_summary["latest_error"]:
        st.warning(f"Latest error: {global_summary['latest_error']}")

    if show_global_detail:
        st.markdown("<div class='psb-label'>Global Run Detail</div>", unsafe_allow_html=True)
        global_detail_rows = summarize_run_detail_rows(
            all_runs,
            start_index=0,
            limit=100,
        )
        if global_detail_rows:
            st.dataframe(global_detail_rows, use_container_width=True)
        else:
            st.info("No logs found yet.")

    last_job_id = st.session_state.get("last_job_id")
    if last_job_id and last_job_id in jobs:
        log_mtime = _log_mtime(LOGS_PATH)
        start = time.perf_counter()
        current_runs = load_runs_cached(str(LOGS_PATH), log_mtime)
        _record_perf("run.load_runs_full", time.perf_counter() - start, len(current_runs))
        last_job = jobs[last_job_id]
        duration_sec = compute_job_duration_seconds(last_job)
        st.markdown("<div class='psb-label'>Last Triggered Job Summary (Web only)</div>", unsafe_allow_html=True)
        summary = summarize_run_window(
            current_runs,
            start_index=int(last_job.get("log_start_index", 0)),
        )
        s1, s2, s3, s4 = st.columns(4)
        with s1:
            st.metric("Processed", summary["total"])
        with s2:
            st.metric("Success", summary["success"])
        with s3:
            st.metric("Error", summary["error"])
        with s4:
            st.metric("Artifacts", summary["with_output"])
        if summary["workflows"]:
            st.caption(f"Workflows: {', '.join(summary['workflows'])}")
        if summary["latest_timestamp"]:
            st.caption(f"Latest log time: {summary['latest_timestamp']}")
        if summary["latest_error"]:
            st.warning(f"Latest error: {summary['latest_error']}")

        st.markdown("<div class='psb-label'>Last Triggered Job Detail (Web only)</div>", unsafe_allow_html=True)
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.caption(f"job_id: {last_job_id}")
        with c2:
            st.caption(f"status: {last_job.get('status', '-')}")
        with c3:
            st.caption(f"started_at: {last_job.get('started_at', '-')}")
        with c4:
            st.caption(f"ended_at: {last_job.get('ended_at', '-')}")
        with c5:
            st.caption(f"duration_sec: {duration_sec if duration_sec is not None else '-'}")

        detail_rows = summarize_run_detail_rows(
            current_runs,
            start_index=int(last_job.get("log_start_index", 0)),
            limit=100,
        )
        if detail_rows:
            st.dataframe(detail_rows, use_container_width=True)
        else:
            st.info("No detailed logs found for this run window yet.")

        pipeline_summary = last_job.get("pipeline_summary")
        if pipeline_summary:
            st.markdown("<div class='psb-label'>Notebook Baseline Summary</div>", unsafe_allow_html=True)
            st.caption(
                f"status: {pipeline_summary.get('status')} | "
                f"started_at: {pipeline_summary.get('started_at')} | "
                f"ended_at: {pipeline_summary.get('ended_at')}"
            )
            st.dataframe(pipeline_summary.get("artifacts", []), use_container_width=True)
            if pipeline_summary.get("error"):
                st.warning(f"Pipeline error: {pipeline_summary.get('error')}")

    st.markdown("<div class='psb-label'>Job Status</div>", unsafe_allow_html=True)
    if not jobs:
        st.info("No jobs yet.")
    else:
        for job_id, info in list(jobs.items())[::-1]:
            spec_label = info.get("spec_id", "-")
            st.write(f"{job_id} | spec: {spec_label} | status: {info['status']}")

    with st.expander("Performance (Debug)", expanded=False):
        st.session_state["perf_debug"] = st.checkbox("Show timings", value=st.session_state.get("perf_debug", False))
        if st.session_state.get("perf_debug"):
            marks = st.session_state.get("perf_marks") or []
            if marks:
                rows = [
                    {"label": label, "seconds": round(seconds, 4), "items": count}
                    for label, seconds, count in marks[-20:]
                ]
                st.dataframe(rows, use_container_width=True)
            else:
                st.info("No timings captured yet.")
