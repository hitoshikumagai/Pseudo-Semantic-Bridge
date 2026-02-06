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
