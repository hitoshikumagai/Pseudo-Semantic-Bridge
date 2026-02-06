import json
from pathlib import Path

import src.bridge.builder as builder


def test_compile_system_spec_writes_json(tmp_path, monkeypatch):
    class DummyConfig:
        def model_dump_json(self, indent=2):
            return json.dumps({"job_name": "test"}, indent=indent)

    def fake_parse_excel_spec(path):
        return DummyConfig()

    monkeypatch.setattr(builder, "parse_excel_spec", fake_parse_excel_spec)

    out_path = tmp_path / "config.json"
    builder._compile_system_spec("dummy.xlsx", str(out_path))

    data = json.loads(out_path.read_text(encoding="utf-8"))
    assert data["job_name"] == "test"


def test_compile_business_rules_writes_json(tmp_path, monkeypatch):
    rules = [{"subject_filter": "A", "task_name": "X"}]

    class DummyDf:
        def where(self, *args, **kwargs):
            return self

        def to_dict(self, orient="records"):
            assert orient == "records"
            return rules

    def fake_read_excel(path):
        return DummyDf()

    monkeypatch.setattr(builder.pd, "read_excel", fake_read_excel)

    excel_path = tmp_path / "rules.xlsx"
    excel_path.write_text("x", encoding="utf-8")

    out_path = tmp_path / "rules.json"
    builder._compile_business_rules(str(excel_path), str(out_path))

    data = json.loads(out_path.read_text(encoding="utf-8"))
    assert data == rules
