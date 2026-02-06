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


def test_compile_system_spec_uses_pydantic_v1_json_method(tmp_path, monkeypatch):
    class DummyConfigV1:
        def json(self, indent=2):
            return json.dumps({"version": 1}, indent=indent)

    monkeypatch.setattr(builder, "parse_excel_spec", lambda _path: DummyConfigV1())

    out_path = tmp_path / "config_v1.json"
    builder._compile_system_spec("dummy.xlsx", str(out_path))

    data = json.loads(out_path.read_text(encoding="utf-8"))
    assert data["version"] == 1


def test_compile_system_spec_falls_back_to_plain_dict(tmp_path, monkeypatch):
    monkeypatch.setattr(builder, "parse_excel_spec", lambda _path: {"job_name": "dict-fallback"})

    out_path = tmp_path / "config_dict.json"
    builder._compile_system_spec("dummy.xlsx", str(out_path))

    data = json.loads(out_path.read_text(encoding="utf-8"))
    assert data["job_name"] == "dict-fallback"


def test_compile_system_spec_skips_when_spec_is_empty(tmp_path, monkeypatch):
    monkeypatch.setattr(builder, "parse_excel_spec", lambda _path: None)

    out_path = tmp_path / "config_empty.json"
    builder._compile_system_spec("dummy.xlsx", str(out_path))

    assert not out_path.exists()


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


def test_compile_business_rules_skips_when_excel_missing(tmp_path, monkeypatch):
    called = {"read_excel": False}

    def fake_read_excel(_path):
        called["read_excel"] = True
        return None

    monkeypatch.setattr(builder.pd, "read_excel", fake_read_excel)

    out_path = tmp_path / "rules_missing.json"
    builder._compile_business_rules(str(tmp_path / "missing.xlsx"), str(out_path))

    assert called["read_excel"] is False
    assert not out_path.exists()


def test_compile_business_rules_swallows_compile_error(tmp_path, monkeypatch):
    excel_path = tmp_path / "rules.xlsx"
    excel_path.write_text("x", encoding="utf-8")

    def fake_read_excel(_path):
        raise RuntimeError("read failed")

    monkeypatch.setattr(builder.pd, "read_excel", fake_read_excel)

    out_path = tmp_path / "rules_error.json"
    builder._compile_business_rules(str(excel_path), str(out_path))

    assert not out_path.exists()


def test_build_all_configs_calls_internal_compilers(monkeypatch):
    calls = []

    monkeypatch.setattr(
        builder,
        "_compile_system_spec",
        lambda excel_path, json_out_path: calls.append(("system", excel_path, json_out_path)),
    )
    monkeypatch.setattr(
        builder,
        "_compile_business_rules",
        lambda excel_path, json_out_path: calls.append(("business", excel_path, json_out_path)),
    )

    builder.build_all_configs()

    assert calls == [
        ("system", "specs/accounting/invoice_bot_v2.xlsx", "configs/accounting/invoice_bot_v2.json"),
        ("business", "specs/accounting/mail_business_rules.xlsx", "configs/accounting/mail_business_rules.json"),
    ]
