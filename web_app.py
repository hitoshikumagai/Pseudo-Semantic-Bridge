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
from src.web.ui.semantic_helpers import (
    _build_hub_edge_rows,
    _build_hub_node_rows,
    _build_ir_coverage_rows,
    _build_rule_ir_relationship_rows,
    _build_semantic_context,
    _collect_semantic_payload_from_state,
    _count_rows,
    _decision_support_rows,
    _default_semantic_layer_spec,
    _extract_instruction_text,
    _fallback_rule_drafts_from_text,
    _get_semantic_assets,
    _intent_rule_from_spec,
    _load_semantic_layer_spec,
    _load_working_state_from_semantic,
    _merge_semantic_spec,
    _normalize_dict_rows,
    _prepare_runtime_from_semantic as _prepare_runtime_from_semantic_impl,
    _runtime_prerequisite_rows,
    _queue_missing_ir_rules,
    _render_input_hub_mapping,
    _rule_signature,
    _runtime_readiness_issues,
    _save_semantic_layer_spec,
    _semantic_mermaid_from_intent_steps,
    _semantic_mermaid_from_spec,
    _semantic_source_rows,
    _summarize_semantic_layer,
    _upsert_intent_history,
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


def _prepare_runtime_from_semantic() -> dict:
    return _prepare_runtime_from_semantic_impl(save_rules, RULES_PATH)

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
if "intent_spec_history" not in st.session_state:
    st.session_state["intent_spec_history"] = []

if "instruction_intake" not in st.session_state:
    st.session_state["instruction_intake"] = None

if "compile_summary" not in st.session_state:
    st.session_state["compile_summary"] = None
if "semantic_layer_spec" not in st.session_state:
    st.session_state["semantic_layer_spec"] = _load_semantic_layer_spec(SEMANTIC_LAYER_PATH)
architecture_views_init = st.session_state.get("semantic_layer_spec", {}).get("architecture_views")
if "diagram_mode" not in st.session_state:
    if isinstance(architecture_views_init, dict):
        st.session_state["diagram_mode"] = architecture_views_init.get("diagram_mode", "table")
    else:
        st.session_state["diagram_mode"] = "table"
if "mermaid_flow" not in st.session_state:
    if isinstance(architecture_views_init, dict):
        st.session_state["mermaid_flow"] = architecture_views_init.get("mermaid_flow", "")
    else:
        st.session_state["mermaid_flow"] = ""
if "mermaid_ai_prompt" not in st.session_state:
    if isinstance(architecture_views_init, dict):
        st.session_state["mermaid_ai_prompt"] = architecture_views_init.get("last_ai_prompt", "")
    else:
        st.session_state["mermaid_ai_prompt"] = ""
if "intent_help_message" not in st.session_state:
    st.session_state["intent_help_message"] = ""
if "candidate_help_message" not in st.session_state:
    st.session_state["candidate_help_message"] = ""
if "run_decision_help_message" not in st.session_state:
    st.session_state["run_decision_help_message"] = ""
if "run_decision_help_spec" not in st.session_state:
    st.session_state["run_decision_help_spec"] = None
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
    semantic_payload = st.session_state.get("semantic_layer_spec") or {}
    log_mtime = _log_mtime(LOGS_PATH)
    start = time.perf_counter()
    runs = load_runs_cached(str(LOGS_PATH), log_mtime)
    _record_perf("overview.load_runs_full", time.perf_counter() - start, len(runs))
    summary = summarize_quality(runs) if runs else {"total": 0, "success": 0, "quality_labeled": 0, "quality_ok": 0}
    semantic_summary = _summarize_semantic_layer(semantic_payload)
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

    st.write("")
    st.markdown("<div class='psb-label'>Semantic Source Of Truth</div>", unsafe_allow_html=True)
    source_rows = _semantic_source_rows(semantic_payload)
    st.dataframe(source_rows, use_container_width=True)

with tabs[1]:
    st.markdown("<div class='psb-label'>Semantic Input / Rule Assets</div>", unsafe_allow_html=True)
    st.caption("Writes into semantic path: automation_assets.rules / automation_assets.proposed_rules")
    _render_input_hub_mapping(
        "Input -> Hub Linkage (Rules)",
        [
            {
                "input": "Rules editor",
                "hub_path": "automation_assets.rules",
                "current": len(st.session_state.get("rules") or []),
            },
            {
                "input": "Proposed rules editor",
                "hub_path": "automation_assets.proposed_rules",
                "current": len(st.session_state.get("proposed_rules") or []),
            },
            {
                "input": "Merge Selected",
                "hub_path": "automation_assets.rules",
                "current": "review -> approve -> merge",
            },
        ],
    )
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
    st.markdown("<div class='psb-label'>Proposed Rules (Review Queue)</div>", unsafe_allow_html=True)
    st.caption("AI and mining drafts land here first. Human review is required before merge.")
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
    _render_input_hub_mapping(
        "Input -> Hub Linkage (Candidates)",
        [
            {
                "input": "Instruction analysis",
                "hub_path": "automation_assets.instruction_intake",
                "current": "ready" if st.session_state.get("instruction_intake") else "missing",
            },
            {
                "input": "Generate Candidates",
                "hub_path": "automation_assets.candidate_meta / candidate_rows",
                "current": len(st.session_state.get("ai_candidates") or []),
            },
            {
                "input": "AI Draft Rules",
                "hub_path": "automation_assets.proposed_rules",
                "current": len(st.session_state.get("proposed_rules") or []),
            },
        ],
    )
    st.write("Generate rule candidates from run logs and route them into review-first rule drafts.")
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
        placeholder="e.g. Prioritize OCR for invoices; store daily reports without OCR.",
        height=80,
        key="candidate_user_instruction",
    )
    col_i1, col_i2 = st.columns([1, 1])
    with col_i1:
        domain_hint = st.text_input("Domain hint", value="accounting_mail_invoice")
    with col_i2:
        instruction_model = st.text_input("Instruction model", value="gpt-4o-mini")

    help_col_a, help_col_b = st.columns([1, 1])
    with help_col_a:
        if st.button("AI Help: Improve Rule Instruction"):
            help_seed = (user_instruction or "").strip()
            if not help_seed:
                help_seed = "I want to automate invoice processing with clear review steps."
            analyzed, help_error, help_source = analyze_and_log_user_instruction(
                user_instruction=help_seed,
                log_path=INTAKE_LOGS_PATH,
                domain_hint=domain_hint,
                use_ai=True,
                model=instruction_model,
            )
            if isinstance(analyzed, dict):
                tasks = analyzed.get("tasks") if isinstance(analyzed.get("tasks"), list) else []
                constraints = analyzed.get("constraints") if isinstance(analyzed.get("constraints"), list) else []
                task_text = ", ".join([str(item) for item in tasks[:3]]) if tasks else "(none)"
                constraint_text = ", ".join([str(item) for item in constraints[:3]]) if constraints else "(none)"
                st.session_state["candidate_help_message"] = (
                    f"AI help ({help_source}): tasks={task_text} | constraints={constraint_text}"
                )
            if help_error:
                st.warning(help_error)
    with help_col_b:
        if st.button("Use Instruction As Draft Prompt"):
            st.session_state["sync_rule_draft_prompt"] = True
            st.rerun()

    if st.session_state.get("candidate_help_message"):
        st.info(st.session_state["candidate_help_message"])

    col_i3, col_i4 = st.columns([1, 1])
    with col_i3:
        if st.button("Analyze Instruction (AI)"):
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
        if st.button("Analyze Instruction (Template)"):
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

    st.write("")
    st.markdown("<div class='psb-label'>AI Draft Rules From Human Input</div>", unsafe_allow_html=True)
    st.caption("Human writes intent -> AI generates draft rules -> human reviews in Proposed Rules queue.")
    if st.session_state.pop("sync_rule_draft_prompt", False):
        st.session_state["rule_draft_prompt_input"] = st.session_state.get("candidate_user_instruction") or ""
    draft_prompt = st.text_area(
        "Draft instruction",
        value=user_instruction,
        placeholder="e.g. For invoice emails, prioritize OCR and keep zipped receipts extractable.",
        height=80,
        key="rule_draft_prompt_input",
    )
    if st.button("Generate AI Draft Rules (Review Queue)"):
        prompt = (draft_prompt or user_instruction or "").strip()
        if not prompt:
            st.warning("Draft instruction is empty.")
        else:
            semantic_context = _build_semantic_context(st.session_state.get("semantic_layer_spec") or {})
            app_context = semantic_context.get("priority_domain") or "accounting_mail_invoice"
            draft_spec, draft_error, draft_source = generate_intent_spec(
                app_context=str(app_context),
                goal=prompt,
                scope="Generate reviewable business rules",
                success="Rules can be reviewed and merged into active set",
                artifacts="semantic-first automation",
                use_ai=True,
                model=instruction_model,
            )
            if draft_error:
                st.warning(draft_error)

            draft_rows = []
            if isinstance(draft_spec, dict):
                draft_inputs = draft_spec.get("inputs") or {}
                draft_inputs["semantic_layer"] = semantic_context
                draft_spec["inputs"] = draft_inputs
                draft_rows, draft_warnings = build_rule_proposals_from_intent_spec(
                    draft_spec,
                    min_quality_score_gate=0.0,
                )
                for warning in draft_warnings:
                    st.warning(warning)

            if not draft_rows:
                analyzed, intake_error, _ = analyze_and_log_user_instruction(
                    user_instruction=prompt,
                    log_path=INTAKE_LOGS_PATH,
                    domain_hint=str(app_context),
                    use_ai=True,
                    model=instruction_model,
                )
                if isinstance(analyzed, dict):
                    st.session_state["instruction_intake"] = analyzed
                if intake_error:
                    st.warning(intake_error)
                fallback_text = _extract_instruction_text(analyzed if isinstance(analyzed, dict) else {}) or prompt
                draft_rows = _fallback_rule_drafts_from_text(fallback_text)

            if not draft_rows:
                st.warning("No draft rules generated.")
            else:
                merged, summary = append_unique_rules(
                    st.session_state.get("proposed_rules") or [],
                    draft_rows,
                )
                st.session_state["proposed_rules"] = merged
                if isinstance(draft_spec, dict):
                    st.session_state["intent_spec"] = draft_spec
                    st.session_state["intent_spec_source"] = draft_source
                    _upsert_intent_history(draft_spec, draft_source)
                st.success(
                    "AI draft rules queued for review: "
                    f"added {summary['added']}, "
                    f"skipped duplicates {summary['skipped_duplicates']}, "
                    f"invalid {summary['skipped_invalid']}."
                )
                st.rerun()

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
    _render_input_hub_mapping(
        "Input -> Hub Linkage (Intent)",
        [
            {
                "input": "Conversation and focus selection",
                "hub_path": "automation_assets.intent_spec",
                "current": "ready"
                if isinstance(st.session_state.get("intent_spec"), dict)
                and str((st.session_state.get("intent_spec") or {}).get("spec_id") or "").strip()
                else "missing",
            },
            {
                "input": "Intent generation source",
                "hub_path": "automation_assets.intent_spec_source",
                "current": st.session_state.get("intent_spec_source") or "missing",
            },
            {
                "input": "Mail rule mapping",
                "hub_path": "automation_assets.intent_spec.inputs.mail_rule",
                "current": "attached to intent spec",
            },
        ],
    )
    st.write("Convert ambiguous requests into an intent spec that is persisted in semantic hub.")
    st.caption("Output: intent spec JSON (spec_id, steps, verification, fallback)")

    llm_model = st.session_state.get("intent_llm_model", "gpt-4o-mini")
    intent_help_seed = st.text_input(
        "Intent help seed",
        value="",
        placeholder="e.g. I need an approval-aware invoice automation flow.",
        key="intent_help_seed",
    )
    if st.button("AI Help: Suggest Intent Structure"):
        help_goal = (intent_help_seed or "").strip() or "Design a practical intent spec for invoice automation."
        help_spec, help_error, help_source = generate_intent_spec(
            app_context="semantic_hub_intent_help",
            goal=help_goal,
            scope="Give minimal but maintainable workflow",
            success="Intent spec is understandable and reviewable",
            artifacts="semantic-first automation",
            use_ai=True,
            model=llm_model,
        )
        if help_error:
            st.warning(help_error)
        help_steps = help_spec.get("steps") if isinstance(help_spec, dict) else []
        step_labels = []
        if isinstance(help_steps, list):
            for step in help_steps[:4]:
                if isinstance(step, dict):
                    step_labels.append(str(step.get("action") or step.get("id") or "step"))
        st.session_state["intent_help_message"] = (
            f"AI help ({help_source}): suggested steps -> {', '.join(step_labels) if step_labels else '(none)'}"
        )
    if st.session_state.get("intent_help_message"):
        st.info(st.session_state["intent_help_message"])

    mail_subject_filter = st.session_state.get("mail_subject_filter", "invoice")
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

    user_message = st.chat_input("Describe what you want to automate")
    if user_message:
        st.session_state["conversation_log"].append({"role": "user", "content": user_message})
        allow_more_needed = st.session_state["conversation_rounds"] >= 3 and not st.session_state["conversation_allow_more"]
        if allow_more_needed:
            st.warning("Follow-up questions are limited to 3 rounds. Allow more if needed.")
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
            "Mark relevant text chunks",
            chunks,
            default=st.session_state.get("conversation_marked") or [],
        )
        st.session_state["conversation_marked"] = marked
        if st.button("Use Marked Chunks as Focus"):
            if not marked:
                st.warning("Marked chunks are empty.")
            else:
                st.session_state["conversation_focus"] = " / ".join(marked)
        focus_choice = st.selectbox("Select the focus point", summary_bullets)
        if st.button("Confirm Focus"):
            st.session_state["conversation_focus"] = focus_choice

    if st.button("Generate Intent Spec (Conversation)", type="primary"):
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
            _upsert_intent_history(spec, source)
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
                "Target app/domain",
                value="mail",
                help="e.g. mail, order processing, invoice management",
            )
            goal = st.text_area(
                "What do you want to do? (Goal)",
                placeholder="e.g. Extract text from attached images and add it to the processing flow",
                height=100,
            )
            scope = st.text_area(
                "Expected scenarios / constraints",
                placeholder="e.g. Start with standalone validation; includes sensitive data",
                height=100,
            )
            success = st.text_area(
                "Success criteria / metrics",
                placeholder="e.g. 95% extraction accuracy and under 3 seconds processing time",
                height=80,
            )
        with col_b:
            st.markdown("<div class='psb-label'>Validation Path</div>", unsafe_allow_html=True)
            path = st.radio(
                "How should we validate?",
                ["Start with standalone quality checks", "Integrate into workflow"],
            )
            customization = st.radio(
                "User customization ownership",
                ["User can handle lightweight development/customization", "Operations team handles most changes"],
            )
            st.markdown("<div class='psb-label' style='margin-top:12px'>Optional Info</div>", unsafe_allow_html=True)
            artifacts = st.text_area(
                "Reference information (optional)",
                placeholder="e.g. existing rules, logs, sample image notes",
                height=120,
            )

        col_c, col_d = st.columns([1, 1])
        with col_c:
            if st.button("Save Intake"):
                st.session_state["quality_intake"] = {
                    "app_context": app_context,
                    "goal": goal,
                    "scope": scope,
                    "success": success,
                    "path": path,
                    "customization": customization,
                    "artifacts": artifacts,
                }
                st.success("Intake saved.")
        with col_d:
            if st.button("Generate Intent Spec (AI)"):
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
                _upsert_intent_history(spec, source)
                st.session_state["intent_spec_error"] = error
                if error:
                    st.warning(error)
                else:
                    st.success("Intent Spec generated by AI.")

        if st.button("Generate Intent Spec (Template)"):
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
            _upsert_intent_history(spec, source)
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
    _render_input_hub_mapping(
        "Upstream Inputs -> Hub Paths",
        [
            {
                "input_tab": "2) Semantic Input: Rules",
                "hub_path": "automation_assets.rules / proposed_rules",
                "current": f"rules {len(st.session_state.get('rules') or [])}, proposed {len(st.session_state.get('proposed_rules') or [])}",
            },
            {
                "input_tab": "3) Semantic Input: Candidates",
                "hub_path": "automation_assets.candidate_meta / candidate_rows / instruction_intake",
                "current": f"candidates {len(st.session_state.get('ai_candidates') or [])}",
            },
            {
                "input_tab": "4) Semantic Input: Intent",
                "hub_path": "automation_assets.intent_spec / intent_spec_source",
                "current": (st.session_state.get("intent_spec") or {}).get("spec_id", "missing")
                if isinstance(st.session_state.get("intent_spec"), dict)
                else "missing",
            },
            {
                "input_tab": "4) Semantic Input: Intent",
                "hub_path": "automation_assets.intent_spec_history",
                "current": len(st.session_state.get("intent_spec_history") or []),
            },
        ],
    )

    semantic_spec = _collect_semantic_payload_from_state(st.session_state.get("semantic_layer_spec") or {})
    intent_history = _normalize_dict_rows(st.session_state.get("intent_spec_history"))
    rule_ir_rows = _build_rule_ir_relationship_rows(
        st.session_state.get("rules") or [],
        st.session_state.get("proposed_rules") or [],
        intent_history,
    )
    ir_coverage_rows = _build_ir_coverage_rows(
        intent_history,
        st.session_state.get("rules") or [],
        st.session_state.get("proposed_rules") or [],
    )

    st.markdown("<div class='psb-label'>Rule / IR Relationship Map</div>", unsafe_allow_html=True)
    map_col_a, map_col_b, map_col_c = st.columns(3)
    with map_col_a:
        st.metric("IR history", len(intent_history))
    with map_col_b:
        st.metric("Linked rules", len([row for row in rule_ir_rows if row.get("status") == "linked"]))
    with map_col_c:
        st.metric(
            "Unlinked proposed",
            len([row for row in rule_ir_rows if row.get("source") == "proposed_rule" and row.get("status") == "unlinked"]),
        )

    if rule_ir_rows:
        st.dataframe(rule_ir_rows, use_container_width=True, hide_index=True)
    else:
        st.info("No rule/IR relationships available yet.")

    st.markdown("<div class='psb-label'>IR Coverage</div>", unsafe_allow_html=True)
    if ir_coverage_rows:
        st.dataframe(ir_coverage_rows, use_container_width=True, hide_index=True)
    else:
        st.info("No IR coverage rows yet.")

    st.markdown("<div class='psb-label'>Relationship Maintenance</div>", unsafe_allow_html=True)
    unlinked_proposed = [
        row for row in rule_ir_rows if row.get("source") == "proposed_rule" and row.get("status") == "unlinked"
    ]
    maintenance_rows = []
    for row in unlinked_proposed:
        maintenance_rows.append(
            {
                "select": False,
                "proposed_index": row.get("index"),
                "subject_filter": row.get("subject_filter"),
                "action_id": row.get("action_id"),
                "task_name": row.get("task_name"),
            }
        )
    edited_unlinked = st.data_editor(
        maintenance_rows,
        num_rows="dynamic",
        use_container_width=True,
        key="unlinked_proposed_maintenance_editor",
    )
    selected_unlinked_indices = [
        int(row.get("proposed_index"))
        for row in edited_unlinked
        if row.get("select") and isinstance(row.get("proposed_index"), int)
    ]
    maintenance_col_a, maintenance_col_b = st.columns([1, 1])
    with maintenance_col_a:
        if st.button("Remove Selected Unlinked Proposed"):
            if not selected_unlinked_indices:
                st.warning("No unlinked proposed rules selected.")
            else:
                remaining = [
                    rule
                    for idx, rule in enumerate(st.session_state.get("proposed_rules") or [])
                    if idx not in set(selected_unlinked_indices)
                ]
                removed = len((st.session_state.get("proposed_rules") or [])) - len(remaining)
                st.session_state["proposed_rules"] = remaining
                st.success(f"Removed {removed} unlinked proposed rules.")
                st.rerun()
    with maintenance_col_b:
        if st.button("Queue Missing IR Rules To Proposed"):
            merged, summary = _queue_missing_ir_rules(
                intent_history,
                st.session_state.get("rules") or [],
                st.session_state.get("proposed_rules") or [],
            )
            st.session_state["proposed_rules"] = merged
            st.success(
                f"Queued {summary['added']} IR-derived rules "
                f"(skipped duplicates {summary['skipped_duplicates']}, invalid {summary['skipped_invalid']})."
            )
            st.rerun()

    st.markdown("<div class='psb-label'>Architecture Views</div>", unsafe_allow_html=True)
    st.caption("Switch between system-design table view and Mermaid flow view for hub maintenance.")
    diagram_mode = st.radio(
        "Visualization mode",
        ["table", "mermaid"],
        key="diagram_mode",
    )
    if diagram_mode == "table":
        st.markdown("<div class='psb-label'>Node Table</div>", unsafe_allow_html=True)
        st.dataframe(_build_hub_node_rows(semantic_spec), use_container_width=True, hide_index=True)
        st.markdown("<div class='psb-label'>Edge Table</div>", unsafe_allow_html=True)
        st.dataframe(_build_hub_edge_rows(semantic_spec), use_container_width=True, hide_index=True)
    else:
        if not st.session_state.get("mermaid_flow"):
            st.session_state["mermaid_flow"] = _semantic_mermaid_from_spec(semantic_spec)
        st.session_state["mermaid_flow"] = st.text_area(
            "Mermaid flow",
            value=st.session_state.get("mermaid_flow") or "",
            height=220,
            key="mermaid_flow_editor",
        )
        st.markdown(
            "```mermaid\n"
            f"{st.session_state.get('mermaid_flow') or ''}\n"
            "```"
        )
        st.session_state["mermaid_ai_prompt"] = st.text_area(
            "AI prompt for Mermaid update",
            value=st.session_state.get("mermaid_ai_prompt") or "",
            placeholder="e.g. Add a review gate between proposed rules and active rules.",
            height=70,
            key="mermaid_ai_prompt_editor",
        )
        mermaid_col_a, mermaid_col_b = st.columns([1, 1])
        with mermaid_col_a:
            if st.button("Refresh Mermaid From Hub Data"):
                st.session_state["mermaid_flow"] = _semantic_mermaid_from_spec(semantic_spec)
                st.success("Mermaid refreshed from current hub state.")
                st.rerun()
        with mermaid_col_b:
            if st.button("AI Help: Generate Mermaid Draft"):
                ai_prompt = (st.session_state.get("mermaid_ai_prompt") or "").strip()
                if not ai_prompt:
                    st.warning("AI prompt for Mermaid is empty.")
                else:
                    ai_spec, ai_error, ai_source = generate_intent_spec(
                        app_context="semantic_hub_visualization",
                        goal=ai_prompt,
                        scope="Generate mermaid-ready maintenance flow",
                        success="Flow is useful for rule and IR maintenance",
                        artifacts="semantic hub",
                        use_ai=True,
                        model=st.session_state.get("intent_llm_model", "gpt-4o-mini"),
                    )
                    if ai_error:
                        st.warning(ai_error)
                    ai_mermaid = _semantic_mermaid_from_intent_steps(ai_spec)
                    if not ai_mermaid:
                        ai_mermaid = _semantic_mermaid_from_spec(semantic_spec)
                    st.session_state["mermaid_flow"] = ai_mermaid
                    st.success(f"AI Mermaid draft generated ({ai_source}).")
                    st.rerun()

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
    runtime_semantic_payload = _collect_semantic_payload_from_state(st.session_state.get("semantic_layer_spec") or {})
    prerequisite_rows = _runtime_prerequisite_rows(runtime_semantic_payload)
    run_issues = _runtime_readiness_issues(runtime_semantic_payload)
    st.markdown("<div class='psb-label'>Run Prerequisites</div>", unsafe_allow_html=True)
    ready_count = len([row for row in prerequisite_rows if row.get("status") == "ready"])
    st.caption(f"Ready {ready_count}/{len(prerequisite_rows)} prerequisites.")
    st.dataframe(prerequisite_rows, use_container_width=True, hide_index=True)
    if run_issues:
        st.warning("Run prerequisites are not complete yet.")
        for issue in run_issues:
            st.write(f"- {issue}")
    else:
        st.success("Run prerequisites are satisfied.")

    st.markdown("<div class='psb-label'>Decision Support Agent</div>", unsafe_allow_html=True)
    st.caption("Use AI suggestions to decide the next actions and draft missing prerequisite fields.")
    st.dataframe(_decision_support_rows(runtime_semantic_payload), use_container_width=True, hide_index=True)
    decision_context = st.text_area(
        "Decision context (optional)",
        value="",
        placeholder="e.g. Keep this minimal and executable for invoice automation.",
        height=70,
        key="run_decision_context",
    )
    decision_model = st.text_input(
        "Decision support model",
        value=st.session_state.get("intent_llm_model", "gpt-4o-mini"),
        key="run_decision_model",
    )
    decision_col_a, decision_col_b = st.columns([1, 1])
    with decision_col_a:
        if st.button("AI Help: Suggest Next Decisions"):
            missing_labels = [
                f"{row.get('path')} ({row.get('fix_in_tab')})"
                for row in prerequisite_rows
                if row.get("status") != "ready"
            ]
            decision_goal = (decision_context or "").strip() or (
                "Unblock run prerequisites: " + (", ".join(missing_labels) if missing_labels else "already ready")
            )
            ai_spec, ai_error, ai_source = generate_intent_spec(
                app_context="run_decision_support",
                goal=decision_goal,
                scope="Provide concise next decisions to unblock automation run",
                success="User can take clear next steps in current tabs",
                artifacts="semantic-layer workflow",
                use_ai=True,
                model=decision_model,
            )
            if ai_error:
                st.warning(ai_error)
            ai_steps = ai_spec.get("steps") if isinstance(ai_spec, dict) else []
            step_labels = []
            if isinstance(ai_steps, list):
                for step in ai_steps[:5]:
                    if isinstance(step, dict):
                        step_labels.append(str(step.get("action") or step.get("id") or "step"))
            st.session_state["run_decision_help_message"] = (
                f"AI decision support ({ai_source}): "
                f"{', '.join(step_labels) if step_labels else '(no step suggestions)'}"
            )
            st.session_state["run_decision_help_spec"] = ai_spec if isinstance(ai_spec, dict) else None
    with decision_col_b:
        if st.button("AI Help: Draft Missing Prerequisites"):
            missing_paths = {
                str(row.get("path") or "")
                for row in prerequisite_rows
                if row.get("status") != "ready"
            }
            if not missing_paths:
                st.session_state["run_decision_help_message"] = "All prerequisites are already satisfied."
                st.rerun()
            decision_goal = (decision_context or "").strip() or "Draft missing run prerequisites for semantic workflow."
            ai_spec, ai_error, ai_source = generate_intent_spec(
                app_context="run_prerequisite_draft",
                goal=decision_goal,
                scope="Draft objective/domain/intent context for immediate next-step execution",
                success="Missing prerequisites get draft values for user review",
                artifacts="semantic-layer run gate",
                use_ai=True,
                model=decision_model,
            )
            if ai_error:
                st.warning(ai_error)
            updated_payload = _merge_semantic_spec(_default_semantic_layer_spec(), runtime_semantic_payload)
            applied = []
            purpose = updated_payload.get("purpose") if isinstance(updated_payload.get("purpose"), dict) else {}
            assets = updated_payload.get("automation_assets") if isinstance(updated_payload.get("automation_assets"), dict) else {}

            if "purpose.objective_statement" in missing_paths:
                objective_draft = str((ai_spec or {}).get("intent") or decision_goal).strip()
                if objective_draft:
                    purpose["objective_statement"] = objective_draft
                    applied.append("purpose.objective_statement")
            if "purpose.priority_domain" in missing_paths:
                domain_draft = str((ai_spec or {}).get("domain") or "operations").strip()
                if domain_draft:
                    purpose["priority_domain"] = domain_draft
                    applied.append("purpose.priority_domain")
            if "automation_assets.intent_spec" in missing_paths and isinstance(ai_spec, dict):
                spec_id = str(ai_spec.get("spec_id") or "").strip()
                if not spec_id:
                    ai_spec["spec_id"] = f"spec-draft-{int(time.time())}"
                assets["intent_spec"] = ai_spec
                assets["intent_spec_source"] = ai_source
                st.session_state["intent_spec"] = ai_spec
                st.session_state["intent_spec_source"] = ai_source
                _upsert_intent_history(ai_spec, ai_source)
                applied.append("automation_assets.intent_spec")
            if "automation_assets.rules" in missing_paths and isinstance(ai_spec, dict):
                draft_rows, _draft_warnings = build_rule_proposals_from_intent_spec(
                    ai_spec,
                    min_quality_score_gate=0.0,
                )
                if not draft_rows:
                    draft_rows = _fallback_rule_drafts_from_text(decision_goal)
                merged_proposed, summary = append_unique_rules(st.session_state.get("proposed_rules") or [], draft_rows)
                st.session_state["proposed_rules"] = merged_proposed
                applied.append(
                    "automation_assets.proposed_rules "
                    f"(queued {summary['added']}, duplicates {summary['skipped_duplicates']})"
                )

            updated_payload["purpose"] = purpose
            updated_payload["automation_assets"] = assets
            st.session_state["semantic_layer_spec"] = updated_payload
            _load_working_state_from_semantic(updated_payload)
            st.session_state["run_decision_help_spec"] = ai_spec if isinstance(ai_spec, dict) else None
            st.session_state["run_decision_help_message"] = (
                "Drafted fields: " + (", ".join(applied) if applied else "none applied")
            )
            st.rerun()

    if st.session_state.get("run_decision_help_message"):
        st.info(st.session_state["run_decision_help_message"])
    if isinstance(st.session_state.get("run_decision_help_spec"), dict):
        with st.expander("Decision Agent Draft Spec", expanded=False):
            st.json(st.session_state["run_decision_help_spec"])

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
            if run_issues:
                st.warning("Cannot compile until run prerequisites are satisfied.")
            else:
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
            if run_issues:
                st.warning("Cannot run until prerequisites are satisfied.")
            elif OutlookAdapter is None:
                st.error(f"Outlook adapter is unavailable: {OUTLOOK_IMPORT_ERROR}")
                st.stop()
            else:
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
