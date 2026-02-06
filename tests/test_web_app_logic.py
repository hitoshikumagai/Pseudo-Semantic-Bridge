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
