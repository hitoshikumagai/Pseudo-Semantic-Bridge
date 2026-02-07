import json
from datetime import datetime, timedelta
from pathlib import Path

import streamlit as st

from src.adapter.outlook import OutlookAdapter
from src.bridge.builder import build_all_configs
from src.engine.core import GenericEtlEngine
from src.web.app_logic import (
    analyze_and_log_user_instruction,
    build_mail_rule_from_intent_spec,
    compute_job_duration_seconds,
    generate_intent_spec,
    generate_intent_spec_from_summary,
    generate_followup_question,
    summarize_conversation,
    load_jsonl_runs,
    load_jsonl_runs_tail,
    load_rules,
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
SYSTEM_CONFIG_PATH = Path("configs/accounting/invoice_bot_v2.json")
LOGS_PATH = Path("data/logs/psb_run.jsonl")
INTAKE_LOGS_PATH = Path("data/logs/intent_intake.jsonl")


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

tabs = st.tabs(["Overview", "Rules", "Design: Rule Builder", "Design: Intent Spec", "Run"])

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

with tabs[0]:
    runs = load_jsonl_runs(LOGS_PATH)
    summary = summarize_quality(runs) if runs else {"total": 0, "success": 0, "quality_labeled": 0, "quality_ok": 0}
    success_rate = round((summary["success"] / summary["total"]) * 100, 1) if summary["total"] else 0.0
    quality_rate = round((summary["quality_ok"] / summary["total"]) * 100, 1) if summary["total"] else 0.0

    st.markdown("<div class='psb-label'>Control Tower</div>", unsafe_allow_html=True)
    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        render_kpi("Rules", str(len(st.session_state["rules"])))
    with col_b:
        render_kpi("Runs", str(summary["total"]))
    with col_c:
        render_kpi("Success Rate", f"{success_rate}%")
    with col_d:
        render_kpi("Quality OK", f"{quality_rate}%")

    st.write("")
    st.markdown("<div class='psb-label'>Planned Modules (Skeleton)</div>", unsafe_allow_html=True)
    col_e, col_f, col_g = st.columns(3)
    with col_e:
        render_skeleton_card("Intake Flow", [92, 75, 60, 84])
    with col_f:
        render_skeleton_card("Rule Suggestion", [88, 82, 68, 54])
    with col_g:
        render_skeleton_card("Run Timeline", [90, 65, 72, 58])

with tabs[1]:
    st.markdown("<div class='psb-label'>Business Rules</div>", unsafe_allow_html=True)
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

with tabs[2]:
    st.markdown("<div class='psb-label'>Design / Rule Builder</div>", unsafe_allow_html=True)
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
        runs = load_jsonl_runs(LOGS_PATH)
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

    runs = load_jsonl_runs(LOGS_PATH)
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
    st.markdown("<div class='psb-label'>Design / Intent Specification</div>", unsafe_allow_html=True)
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
            st.markdown(f"**{role}**: {content}")
    else:
        st.caption("Conversation log is empty.")

    convo_input = st.text_area("User message", placeholder="やりたいことを短く入力してください", height=80)
    convo_col_a, convo_col_b, convo_col_c = st.columns([1, 1, 1])
    with convo_col_a:
        if st.button("Add Message"):
            if convo_input.strip():
                st.session_state["conversation_log"].append({"role": "user", "content": convo_input.strip()})
                st.rerun()
            else:
                st.warning("User message is empty.")
    with convo_col_b:
        allow_more_needed = st.session_state["conversation_rounds"] >= 3 and not st.session_state["conversation_allow_more"]
        if allow_more_needed:
            st.info("質問は3回まで。さらに必要なら許可してください。")
            if st.button("Allow More Questions"):
                st.session_state["conversation_allow_more"] = True
                st.rerun()
        else:
            if st.button("Ask Next Question (AI)"):
                question, error, source = generate_followup_question(
                    conversation=st.session_state["conversation_log"],
                    domain_hint="accounting_mail_invoice",
                    use_ai=True,
                    model=llm_model,
                    round_index=st.session_state["conversation_rounds"],
                )
                if error:
                    st.warning(error)
                st.session_state["conversation_log"].append({"role": "assistant", "content": question})
                st.session_state["conversation_rounds"] += 1
                st.rerun()
    with convo_col_c:
        if st.button("Reset Conversation"):
            st.session_state["conversation_log"] = []
            st.session_state["conversation_rounds"] = 0
            st.session_state["conversation_allow_more"] = False
            st.session_state["conversation_summary"] = []
            st.session_state["conversation_focus"] = None
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

    summary_bullets = st.session_state.get("conversation_summary") or []
    if summary_bullets:
        st.markdown("<div class='psb-label'>Summary (Bullet)</div>", unsafe_allow_html=True)
        for bullet in summary_bullets:
            st.write(f"- {bullet}")
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

        if st.button("Append Mail Rule To Rules", type="primary"):
            current_spec = st.session_state.get("intent_spec")
            if not current_spec:
                st.warning("Intent Spec not generated yet.")
            else:
                rule, rule_error = build_mail_rule_from_intent_spec(current_spec)
                if rule_error:
                    st.warning(rule_error)
                else:
                    st.session_state["rules"].append(rule)
                    st.success("Rule appended to Rules (draft).")
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
    jobs = st.session_state.setdefault("jobs", {})
    st.markdown("<div class='psb-label'>Run</div>", unsafe_allow_html=True)
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
            baseline_count = len(load_jsonl_runs(LOGS_PATH))
            current_spec = st.session_state.get("intent_spec") or {}

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

    all_runs = load_jsonl_runs_tail(LOGS_PATH, max_lines=int(tail_limit))
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
        current_runs = load_jsonl_runs(LOGS_PATH)
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
