import json
from pathlib import Path

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
