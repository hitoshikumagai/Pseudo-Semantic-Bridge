"""Facade module for web services.

This module re-exports functions from smaller service modules so existing imports remain stable.
"""

import threading
import time

from src.web.services.run_service import (
    load_system_config,
    run_pipeline_baseline,
    run_engine_job,
    summarize_run_window,
    summarize_run_detail_rows,
    compute_job_duration_seconds,
    append_jsonl_record,
    run_bridge_compile_summary,
    load_jsonl_runs,
    load_jsonl_runs_tail,
    summarize_quality,
)
from src.web.services.rules_service import (
    load_rules,
    save_rules,
    parse_rules_input,
    propose_rule_candidates,
    build_mail_rule_from_intent_spec,
    append_unique_rules,
    build_rule_proposals_from_intent_spec,
    merge_proposed_rules,
)
from src.web.services.quality_service import (
    parse_feedback_input,
    run_rule_check,
    run_workflow_improvement,
    run_skill_suggestion,
    run_quality_agents,
)
from src.web.services.intent_service import (
    _build_template_spec,
    _normalize_intent_spec,
    generate_followup_question,
    summarize_conversation,
    generate_intent_spec_from_summary,
    generate_intent_spec,
    analyze_user_instruction,
    analyze_and_log_user_instruction,
)


def start_job(jobs: dict, run_fn):
    job_id = f"job-{int(time.time())}"
    jobs[job_id] = {"status": "queued"}
    thread = threading.Thread(target=run_fn, args=(job_id,), daemon=True)
    thread.start()
    return job_id


__all__ = [
    "load_system_config",
    "run_pipeline_baseline",
    "run_engine_job",
    "summarize_run_window",
    "summarize_run_detail_rows",
    "compute_job_duration_seconds",
    "append_jsonl_record",
    "run_bridge_compile_summary",
    "start_job",
    "load_jsonl_runs",
    "load_jsonl_runs_tail",
    "summarize_quality",
    "load_rules",
    "save_rules",
    "parse_rules_input",
    "propose_rule_candidates",
    "build_mail_rule_from_intent_spec",
    "append_unique_rules",
    "build_rule_proposals_from_intent_spec",
    "merge_proposed_rules",
    "parse_feedback_input",
    "run_rule_check",
    "run_workflow_improvement",
    "run_skill_suggestion",
    "run_quality_agents",
    "_build_template_spec",
    "_normalize_intent_spec",
    "generate_followup_question",
    "summarize_conversation",
    "generate_intent_spec_from_summary",
    "generate_intent_spec",
    "analyze_user_instruction",
    "analyze_and_log_user_instruction",
    "time",
]
