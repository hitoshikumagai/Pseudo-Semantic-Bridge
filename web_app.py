from datetime import datetime, timedelta
from pathlib import Path

import streamlit as st

from src.bridge.builder import build_all_configs
from src.engine.core import GenericEtlEngine
from src.adapter.outlook import OutlookAdapter
from src.web.app_logic import (
    load_rules,
    load_jsonl_runs,
    propose_rule_candidates,
    save_rules,
    run_engine_job,
    start_job,
    summarize_quality,
)


APP_TITLE = "Pseudo Semantic Bridge"
RULES_PATH = Path("configs/accounting/mail_business_rules.json")
SYSTEM_CONFIG_PATH = Path("configs/accounting/invoice_bot_v2.json")
LOGS_PATH = Path("data/logs/psb_run.jsonl")


st.set_page_config(page_title=APP_TITLE, layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;600&family=JetBrains+Mono:wght@400;600&display=swap');
    :root {
        --bg: #f6f4ef;
        --ink: #111111;
        --accent: #157a6e;
        --accent-2: #f59e0b;
        --card: #ffffff;
        --muted: #6b7280;
    }
    html, body, [class*="stApp"] {
        font-family: "Space Grotesk", sans-serif;
        color: var(--ink);
        background: var(--bg);
    }
    .psb-hero {
        background: linear-gradient(135deg, #e6f4f1 0%, #fff7e6 100%);
        padding: 24px 28px;
        border-radius: 16px;
        border: 1px solid #e5e7eb;
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
        border: 1px solid #e5e7eb;
    }
    .psb-idea-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 12px;
    }
    .psb-idea-card {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 14px 16px;
        box-shadow: 0 8px 20px rgba(15, 23, 42, 0.05);
    }
    .psb-idea-title {
        font-weight: 600;
        margin: 0 0 6px 0;
    }
    .psb-idea-meta {
        font-size: 12px;
        color: var(--muted);
        margin-bottom: 10px;
    }
    .psb-pill {
        display: inline-block;
        font-size: 11px;
        padding: 2px 8px;
        border-radius: 999px;
        background: #f3f4f6;
        color: #111827;
        margin-right: 6px;
    }
    .psb-pill.high { background: #fee2e2; color: #991b1b; }
    .psb-pill.medium { background: #fef3c7; color: #92400e; }
    .psb-pill.low { background: #e0f2fe; color: #075985; }
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

st.markdown(
    f"""
    <div class="psb-hero">
        <div class="psb-title">{APP_TITLE}</div>
        <p class="psb-sub">Edit business rules, then run the pipeline as a background job.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.write("")

tabs = st.tabs(["Rules", "AI Rule Builder", "Quality Feedback", "Run"])

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

with tabs[0]:
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
            st.experimental_rerun()

with tabs[1]:
    st.markdown("<div class='psb-label'>AI Rule Builder (Preview)</div>", unsafe_allow_html=True)
    st.write("Generate Excel-compatible rule rows from JSONL logs with user guidance.")

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

    if st.session_state.get("ai_meta"):
        st.markdown("<div class='psb-label'>Candidate Meta</div>", unsafe_allow_html=True)
        st.dataframe(st.session_state["ai_meta"], use_container_width=True)

    if st.session_state.get("ai_candidates"):
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
                st.experimental_rerun()

with tabs[2]:
    st.markdown("<div class='psb-label'>Quality Intake</div>", unsafe_allow_html=True)
    st.write("受付フォームで目的を整理し、AIがワークフロー化/スキル化を提案します。")

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
        if st.button("AIで整理（準備）", type="primary"):
            st.session_state["quality_intake_ready"] = True

    st.write("")
    st.markdown("<div class='psb-label'>AI解析結果（予定）</div>", unsafe_allow_html=True)
    st.caption("次のフェーズで生成AIを接続し、要約・ワークフロー案・スキル案を表示します。")
    if st.session_state.get("quality_intake"):
        st.json(st.session_state["quality_intake"])

with tabs[3]:
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
    if st.button("Run Pipeline"):
        st.session_state.setdefault("jobs", {})
        job_id = start_job(
            st.session_state["jobs"],
            lambda job_id: run_engine_job(
                st.session_state["jobs"],
                job_id,
                build_all_configs,
                SYSTEM_CONFIG_PATH,
                OutlookAdapter,
                GenericEtlEngine,
            ),
        )
        st.session_state["last_job_id"] = job_id

    st.write("")
    st.markdown("<div class='psb-label'>Job Status</div>", unsafe_allow_html=True)
    jobs = st.session_state.get("jobs", {})
    if not jobs:
        st.info("No jobs yet.")
    else:
        for job_id, info in list(jobs.items())[::-1]:
            st.write(f"{job_id}: {info['status']}")
