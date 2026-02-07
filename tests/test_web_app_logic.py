import json
import threading
from pathlib import Path

import pytest

from src.web import app_logic
from src.schema.definitions import OutlookConfig


def test_load_rules_missing_file_returns_empty(tmp_path):
    rules = app_logic.load_rules(tmp_path / "missing.json")
    assert rules == []


def test_save_and_load_rules(tmp_path):
    rules_path = tmp_path / "rules.json"
    rules = [{"subject_filter": "A", "task_name": "X"}]
    app_logic.save_rules(rules_path, rules)
    loaded = app_logic.load_rules(rules_path)
    assert loaded == rules


def test_load_system_config(tmp_path):
    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps(
            {
                "job_name": "test",
                "version": "2.0",
                "domain": "test",
                "search_keywords": ["x"],
                "destination_path": "/tmp",
                "rules": [],
            }
        ),
        encoding="utf-8",
    )

    config = app_logic.load_system_config(config_path)
    assert isinstance(config, OutlookConfig)
    assert config.job_name == "test"


def test_run_engine_job_sets_done(monkeypatch, tmp_path):
    jobs = {"job-1": {"status": "queued"}}

    def build_fn():
        return None

    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps(
            {
                "job_name": "test",
                "version": "2.0",
                "domain": "test",
                "search_keywords": ["x"],
                "destination_path": "/tmp",
                "rules": [],
            }
        ),
        encoding="utf-8",
    )

    class DummyEngine:
        def __init__(self, config, adapter):
            self.config = config
            self.adapter = adapter

        def run(self):
            return None

    app_logic.run_engine_job(
        jobs,
        "job-1",
        build_fn,
        config_path,
        adapter_factory=lambda: object(),
        engine_factory=DummyEngine,
    )

    assert jobs["job-1"]["status"] == "done"
    assert jobs["job-1"]["pipeline_summary"]["status"] == "done"


def test_run_engine_job_sets_error(tmp_path):
    jobs = {"job-1": {"status": "queued"}}

    def build_fn():
        raise RuntimeError("boom")

    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps(
            {
                "job_name": "test",
                "version": "2.0",
                "domain": "test",
                "search_keywords": ["x"],
                "destination_path": "/tmp",
                "rules": [],
            }
        ),
        encoding="utf-8",
    )

    app_logic.run_engine_job(
        jobs,
        "job-1",
        build_fn,
        config_path,
        adapter_factory=lambda: object(),
        engine_factory=lambda config, adapter: object(),
    )

    assert jobs["job-1"]["status"].startswith("error:")
    assert jobs["job-1"]["pipeline_summary"]["status"] == "error"


def test_run_engine_job_initializes_missing_job_entry(tmp_path):
    jobs = {}

    def build_fn():
        return None

    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps(
            {
                "job_name": "test",
                "version": "2.0",
                "domain": "test",
                "search_keywords": ["x"],
                "destination_path": "/tmp",
                "rules": [],
            }
        ),
        encoding="utf-8",
    )

    class DummyEngine:
        def __init__(self, config, adapter):
            self.config = config
            self.adapter = adapter

        def run(self):
            return None

    app_logic.run_engine_job(
        jobs,
        "job-1",
        build_fn,
        config_path,
        adapter_factory=lambda: object(),
        engine_factory=DummyEngine,
    )

    assert jobs["job-1"]["status"] == "done"


def test_run_pipeline_baseline_success(tmp_path):
    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps(
            {
                "job_name": "test",
                "version": "2.0",
                "domain": "test",
                "search_keywords": ["x"],
                "destination_path": "/tmp",
                "rules": [],
            }
        ),
        encoding="utf-8",
    )

    class DummyEngine:
        def __init__(self, config, adapter):
            self.config = config
            self.adapter = adapter

        def run(self):
            return None

    summary = app_logic.run_pipeline_baseline(
        build_fn=lambda: None,
        config_path=config_path,
        adapter_factory=lambda: object(),
        engine_factory=DummyEngine,
    )
    assert summary["status"] == "done"
    assert summary["error"] is None
    assert summary["artifacts"][0]["exists"] is True


def test_run_pipeline_baseline_error(tmp_path):
    config_path = tmp_path / "missing.json"
    summary = app_logic.run_pipeline_baseline(
        build_fn=lambda: None,
        config_path=config_path,
        adapter_factory=lambda: object(),
        engine_factory=lambda config, adapter: object(),
    )
    assert summary["status"] == "error"
    assert summary["error"]
    assert summary["artifacts"][0]["exists"] is False


def test_load_jsonl_runs_skips_invalid_lines(tmp_path):
    log_path = tmp_path / "runs.jsonl"
    log_path.write_text(
        "\n".join(
            [
                '{"run_id":"1","result":{"status":"success"}}',
                "not-json",
                "",
                '{"run_id":"2","result":{"status":"fail"}}',
            ]
        ),
        encoding="utf-8",
    )

    runs = app_logic.load_jsonl_runs(log_path)
    assert [run["run_id"] for run in runs] == ["1", "2"]


def test_load_jsonl_runs_tail_limits_lines(tmp_path):
    log_path = tmp_path / "runs.jsonl"
    log_path.write_text(
        "\n".join(
            [
                '{"run_id":"1","result":{"status":"success"}}',
                '{"run_id":"2","result":{"status":"success"}}',
                '{"run_id":"3","result":{"status":"success"}}',
                '{"run_id":"4","result":{"status":"success"}}',
            ]
        ),
        encoding="utf-8",
    )

    runs = app_logic.load_jsonl_runs_tail(log_path, max_lines=2)
    assert [run["run_id"] for run in runs] == ["3", "4"]


def test_summarize_quality_counts_success_and_quality_labels():
    runs = [
        {"result": {"status": "success"}, "quality": {"label": "OK"}},
        {"result": {"status": "success"}, "quality": {"label": "NG"}},
        {"result": {"status": "fail"}, "quality": {"score": 0.9}},
        {"result": {"status": "fail"}, "quality": {}},
    ]

    summary = app_logic.summarize_quality(runs)
    assert summary["total"] == 4
    assert summary["success"] == 2
    assert summary["quality_labeled"] == 3
    assert summary["quality_ok"] == 2


def test_propose_rule_candidates_filters_by_quality_and_samples():
    runs = [
        {
            "input": {"subject": "請求書", "attachment_ext": ".pdf", "has_attachment": True},
            "action_id": "ocr_process",
            "result": {"status": "success"},
            "quality": {"label": "ok"},
        },
        {
            "input": {"subject": "請求書", "attachment_ext": ".pdf", "has_attachment": True},
            "action_id": "ocr_process",
            "result": {"status": "success"},
            "quality": {"label": "ok"},
        },
        {
            "input": {"subject": "請求書", "attachment_ext": ".pdf", "has_attachment": True},
            "action_id": "ocr_process",
            "result": {"status": "success"},
            "quality": {"label": "ng"},
        },
        {
            "input": {"subject": "日報", "attachment_ext": ".xlsx", "has_attachment": True},
            "action_id": "save_only",
            "result": {"status": "success"},
            "quality": {"label": "ok"},
        },
    ]

    meta, candidates = app_logic.propose_rule_candidates(
        runs,
        min_samples=3,
        min_quality_rate=0.6,
    )

    assert len(meta) == 1
    assert meta[0]["subject_filter"] == "請求書"
    assert meta[0]["quality_rate"] >= 0.6

    assert len(candidates) == 1
    assert candidates[0]["action_id"] == "ocr_process"
    assert candidates[0]["target_ext"] == ".pdf"


def test_start_job_queues_and_starts_thread(monkeypatch):
    jobs = {}
    called = []
    done = threading.Event()

    monkeypatch.setattr(app_logic.time, "time", lambda: 1700000000)

    def run_fn(job_id):
        called.append(job_id)
        done.set()

    job_id = app_logic.start_job(jobs, run_fn)
    assert job_id == "job-1700000000"
    assert jobs[job_id]["status"] == "queued"
    assert done.wait(1.0)
    assert called == [job_id]


@pytest.mark.parametrize(
    ("raw_text", "expected_format", "expected_parsed"),
    [
        ("", "empty", None),
        ('{"a":1}', "json", {"a": 1}),
        ("hello", "text", None),
    ],
)
def test_parse_feedback_input_formats(raw_text, expected_format, expected_parsed):
    parsed = app_logic.parse_feedback_input(raw_text)
    assert parsed["format"] == expected_format
    assert parsed["parsed"] == expected_parsed


def test_parse_feedback_input_invalid_json_falls_back_to_text():
    parsed = app_logic.parse_feedback_input("{invalid")
    assert parsed["format"] == "text"
    assert parsed["parsed"] is None


def test_parse_rules_input_plain_text_returns_error():
    rules, error = app_logic.parse_rules_input("not-json")
    assert rules == []
    assert "not JSON" in error


def test_parse_rules_input_invalid_json_returns_error():
    rules, error = app_logic.parse_rules_input("[invalid")
    assert rules == []
    assert error.startswith("Invalid JSON:")


def test_parse_rules_input_dict_is_wrapped_into_list():
    rules, error = app_logic.parse_rules_input('{"subject_filter":"A","action_id":"x"}')
    assert error is None
    assert rules == [{"subject_filter": "A", "action_id": "x"}]


def test_parse_rules_input_skips_non_dict_entries():
    rules, error = app_logic.parse_rules_input('[{"subject_filter":"A","action_id":"x"}, 1, "x"]')
    assert len(rules) == 1
    assert "skipped" in error


def test_run_rule_check_detects_missing_keys_and_feedback_hints():
    feedback = {"raw": "slow and fail"}
    rules = [{"subject_filter": "invoice"}, {"action_id": "save_only"}]
    result = app_logic.run_rule_check(feedback, rules)
    assert result["status"] == "needs_fix"
    assert len(result["missing"]) == 2
    assert any("slowness" in hint for hint in result["hints"])
    assert any("errors/failures" in hint for hint in result["hints"])


def test_run_workflow_improvement_generates_prioritized_ideas():
    result = app_logic.run_workflow_improvement({"raw": "manual operation is slow and unclear"})
    assert result["status"] == "ok"
    assert [idea["priority"] for idea in result["ideas"]] == ["high", "high", "medium"]
    assert result["ideas"][0]["title"] == "Template Picker"


def test_run_workflow_improvement_returns_default_ideas_when_no_signal():
    result = app_logic.run_workflow_improvement({"raw": "looks fine"})
    assert [idea["title"] for idea in result["ideas"]] == ["One-Click Rerun", "Before/After Diff"]


def test_run_skill_suggestion_keyword_and_default_paths():
    keyword = app_logic.run_skill_suggestion({"raw": "Need OCR and classify docs"})
    assert [skill["name"] for skill in keyword["skills"]] == ["ocr_quality_audit", "document_classifier"]

    default = app_logic.run_skill_suggestion({"raw": "no special request"})
    assert default["skills"][0]["name"] == "quality_summary"


def test_run_quality_agents_aggregates_results_and_can_toggle_agents():
    result = app_logic.run_quality_agents(
        feedback_text="manual and slow",
        context_text='{"tenant":"x"}',
        rules_text='{"subject_filter":"A"}',
    )
    assert result["input"]["feedback"]["format"] == "text"
    assert result["input"]["context"]["format"] == "json"
    assert result["input"]["rules_count"] == 1
    assert result["rules_error"] is None
    assert [row["agent"] for row in result["results"]] == [
        "rule_check",
        "workflow_improvement",
        "skill_suggestion",
    ]

    only_skill = app_logic.run_quality_agents(
        feedback_text="classify",
        context_text="",
        rules_text="[]",
        run_rule=False,
        run_workflow=False,
        run_skill=True,
    )
    assert [row["agent"] for row in only_skill["results"]] == ["skill_suggestion"]


def test_generate_intent_spec_template_mode():
    spec, error, source = app_logic.generate_intent_spec(
        app_context="メール",
        goal="請求書メールをOCRして保存",
        scope="過去7日",
        success="抽出率95%",
        artifacts="サンプルなし",
        use_ai=False,
    )
    assert error is None
    assert source == "template"
    assert spec["spec_version"] == "1.0"
    assert spec["steps"][0]["action"] == "fetch_mails"


def test_generate_intent_spec_ai_mode_without_key_falls_back(monkeypatch):
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    spec, error, source = app_logic.generate_intent_spec(
        app_context="メール",
        goal="請求書の処理",
        scope="標準運用",
        success="欠損なし",
        artifacts="",
        use_ai=True,
    )
    assert source == "template"
    assert "OPENAI_API_KEY not found" in error
    assert spec["spec_version"] == "1.0"


def test_summarize_run_window_counts_status_and_error():
    runs = [
        {"workflow": "engine", "timestamp": "2026-02-06T11:10:01Z", "result": {"status": "success", "output_path": None}},
        {"workflow": "engine", "timestamp": "2026-02-06T11:10:02Z", "result": {"status": "error", "error": "boom"}},
        {"workflow": "mail_workflow", "timestamp": "2026-02-06T11:10:03Z", "result": {"status": "success", "output_path": "data/out/a.txt"}},
    ]
    summary = app_logic.summarize_run_window(runs, start_index=0)
    assert summary["total"] == 3
    assert summary["success"] == 2
    assert summary["error"] == 1
    assert summary["with_output"] == 1
    assert summary["latest_error"] == "boom"
    assert summary["latest_timestamp"] == "2026-02-06T11:10:03Z"
    assert summary["workflows"] == ["engine", "mail_workflow"]


def test_summarize_run_window_respects_start_index():
    runs = [
        {"workflow": "engine", "result": {"status": "success"}},
        {"workflow": "engine", "result": {"status": "error", "error": "x"}},
    ]
    summary = app_logic.summarize_run_window(runs, start_index=1)
    assert summary["total"] == 1
    assert summary["success"] == 0
    assert summary["error"] == 1


def test_summarize_run_detail_rows_extracts_expected_fields():
    runs = [
        {
            "timestamp": "2026-02-06T11:10:01Z",
            "workflow": "engine",
            "action_id": "save_only",
            "input": {"subject": "Invoice 123", "attachment_ext": ".pdf"},
            "result": {"status": "success", "error": None, "output_path": "data/out/a.txt"},
            "quality": {"label": "ok", "score": 0.91},
        }
    ]
    rows = app_logic.summarize_run_detail_rows(runs, start_index=0, limit=10)
    assert len(rows) == 1
    assert rows[0]["workflow"] == "engine"
    assert rows[0]["subject"] == "Invoice 123"
    assert rows[0]["quality_score"] == 0.91


def test_compute_job_duration_seconds():
    job = {
        "started_at": "2026-02-06T11:10:01+00:00",
        "ended_at": "2026-02-06T11:10:03.250000+00:00",
    }
    assert app_logic.compute_job_duration_seconds(job) == 2.25


def test_analyze_user_instruction_template_mode():
    analyzed, error, source = app_logic.analyze_user_instruction(
        user_instruction="請求書をOCRして保存したい",
        domain_hint="accounting_mail_invoice",
        use_ai=False,
    )
    assert error is None
    assert source == "template"
    assert analyzed["record_type"] == "instruction_intake"
    assert analyzed["tasks"]
    assert analyzed["follow_up_questions"]


def test_analyze_user_instruction_ai_without_key_falls_back(monkeypatch):
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    analyzed, error, source = app_logic.analyze_user_instruction(
        user_instruction="分類して処理したい",
        domain_hint="accounting_mail_invoice",
        use_ai=True,
    )
    assert source == "template"
    assert "OPENAI_API_KEY not found" in error
    assert analyzed["intent_summary"]


def test_analyze_and_log_user_instruction_writes_jsonl(tmp_path):
    log_path = tmp_path / "intent_intake.jsonl"
    record, error, source = app_logic.analyze_and_log_user_instruction(
        user_instruction="請求書をOCRして保存したい",
        log_path=log_path,
        domain_hint="accounting_mail_invoice",
        use_ai=False,
    )
    assert error is None
    assert source == "template"
    lines = log_path.read_text(encoding="utf-8").strip().splitlines()
    assert len(lines) == 1
    saved = json.loads(lines[0])
    assert saved["record_type"] == "instruction_intake"
    assert saved["source"] == "template"
    assert record["instruction_raw"] == "請求書をOCRして保存したい"


def test_run_bridge_compile_summary_success(tmp_path):
    system_path = tmp_path / "invoice_bot_v2.json"
    rules_path = tmp_path / "mail_business_rules.json"

    def build_fn():
        system_path.write_text("{}", encoding="utf-8")
        rules_path.write_text("[]", encoding="utf-8")

    summary = app_logic.run_bridge_compile_summary(build_fn, system_path, rules_path)
    assert summary["status"] == "done"
    assert summary["error"] is None
    assert summary["artifacts"][0]["exists"] is True
    assert summary["artifacts"][1]["exists"] is True


def test_run_bridge_compile_summary_error(tmp_path):
    system_path = tmp_path / "invoice_bot_v2.json"
    rules_path = tmp_path / "mail_business_rules.json"

    def build_fn():
        raise RuntimeError("compile failed")

    summary = app_logic.run_bridge_compile_summary(build_fn, system_path, rules_path)
    assert summary["status"] == "error"
    assert "compile failed" in summary["error"]
    assert summary["artifacts"][0]["exists"] is False
