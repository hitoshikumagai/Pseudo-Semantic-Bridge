from pathlib import Path

import streamlit as st

from src.bridge.builder import build_all_configs
from src.engine.core import GenericEtlEngine
from src.adapter.outlook import OutlookAdapter
from src.web.app_logic import (
    load_rules,
    save_rules,
    run_engine_job,
    start_job,
)


APP_TITLE = "Pseudo Semantic Bridge"
RULES_PATH = Path("configs/accounting/mail_business_rules.json")
SYSTEM_CONFIG_PATH = Path("configs/accounting/invoice_bot_v2.json")


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

left, right = st.columns([2, 1], gap="large")

with left:
    st.markdown("<div class='psb-label'>Business Rules</div>", unsafe_allow_html=True)
    rules = load_rules(RULES_PATH)
    if not rules:
        rules = [
            {
                "subject_filter": "Invoice",
                "task_name": "INVOICE",
                "require_attachment": True,
                "target_ext": ".pdf",
                "action_id": "ocr_process",
                "parameters": {"lang": "jpn"},
            }
        ]

    edited = st.data_editor(
        rules,
        num_rows="dynamic",
        use_container_width=True,
    )

    col_a, col_b = st.columns([1, 1])
    with col_a:
        if st.button("Save Rules", type="primary"):
            save_rules(RULES_PATH, edited)
            st.success("Rules saved.")
    with col_b:
        if st.button("Reload"):
            st.experimental_rerun()

with right:
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
