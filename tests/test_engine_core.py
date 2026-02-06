import types

import src.engine.core as core
from src.engine.core import GenericEtlEngine
from src.schema.definitions import AttachmentRule, OutlookConfig


class DummyItem:
    def __init__(self, name: str, extension: str, children=None, child_error: bool = False):
        self._name = name
        self._extension = extension
        self._children = children or []
        self._child_error = child_error

    @property
    def name(self) -> str:
        return self._name

    @property
    def extension(self) -> str:
        return self._extension

    @property
    def is_container(self) -> bool:
        return bool(self._children) or self._child_error

    def get_children(self):
        if self._child_error:
            raise RuntimeError("children unavailable")
        return self._children


class DummyAdapter:
    def __init__(self, items):
        self._items = items
        self.keywords = []

    def fetch_items(self, keyword):
        self.keywords.append(keyword)
        return self._items


def _build_config(rules):
    return OutlookConfig(
        job_name="test",
        domain="test",
        search_keywords=["alpha"],
        destination_path="/tmp",
        rules=rules,
    )


def test_build_executor_falls_back_to_none_on_invalid_max_concurrency():
    config = _build_config(
        [
            AttachmentRule(
                extension=".msg",
                processor_id="save_only",
                parameters={"max_concurrency": "bad-number"},
            )
        ]
    )

    engine = GenericEtlEngine(config, DummyAdapter([]))
    assert engine._executor is None


def test_try_execute_rule_returns_false_when_no_matching_extension():
    config = _build_config(
        [AttachmentRule(extension=".pdf", processor_id="save_only", parameters={})]
    )
    engine = GenericEtlEngine(config, DummyAdapter([]))

    assert engine._try_execute_rule(DummyItem("mail-1", ".msg")) is False


def test_try_execute_rule_uses_value_attribute_for_processor_id(monkeypatch):
    records = []
    calls = []

    def fake_append_run(record):
        records.append(record)

    def fake_get_processor(processor_id):
        assert processor_id == "processor_from_value"

        def handler(item, output_dir, params):
            calls.append((item.name, output_dir, params))

        return handler

    monkeypatch.setattr(core, "append_run", fake_append_run)
    monkeypatch.setattr(core, "get_processor", fake_get_processor)

    config = _build_config(
        [AttachmentRule(extension=".msg", processor_id="save_only", parameters={})]
    )
    config.rules = [
        types.SimpleNamespace(
            extension=".msg",
            processor_id=types.SimpleNamespace(value="processor_from_value"),
            parameters={},
        )
    ]
    engine = GenericEtlEngine(config, DummyAdapter([]))

    item = DummyItem("mail-1", ".msg", child_error=True)
    assert engine._try_execute_rule(item) is True
    assert calls == [("mail-1", "/tmp", {})]
    assert records[0]["result"]["status"] == "success"
    assert records[0]["input"]["has_attachment"] is False


def test_try_execute_rule_logs_mail_workflow_error_and_returns_false(monkeypatch):
    records = []

    def fake_append_run(record):
        records.append(record)

    def fake_get_processor(_processor_id):
        def handler(item, output_dir, params):
            raise RuntimeError("mail workflow failed")

        return handler

    monkeypatch.setattr(core, "append_run", fake_append_run)
    monkeypatch.setattr(core, "get_processor", fake_get_processor)

    config = _build_config(
        [AttachmentRule(extension=".msg", processor_id="mail_workflow", parameters={})]
    )
    engine = GenericEtlEngine(config, DummyAdapter([]))

    ok = engine._try_execute_rule(DummyItem("mail-1", ".msg"))
    assert ok is False
    assert len(records) == 1
    assert records[0]["processor_id"] == "mail_workflow"
    assert records[0]["result"]["status"] == "error"
    assert "mail workflow failed" in records[0]["result"]["error"]


def test_try_execute_rule_invalid_max_concurrency_returns_false(monkeypatch):
    records = []
    handler_calls = []

    def fake_append_run(record):
        records.append(record)

    def fake_get_processor(_processor_id):
        def handler(item, output_dir, params):
            handler_calls.append((item.name, output_dir, params))

        return handler

    monkeypatch.setattr(core, "append_run", fake_append_run)
    monkeypatch.setattr(core, "get_processor", fake_get_processor)

    config = _build_config(
        [
            AttachmentRule(
                extension=".msg",
                processor_id="save_only",
                parameters={"max_concurrency": "invalid"},
            )
        ]
    )
    engine = GenericEtlEngine(config, DummyAdapter([]))

    ok = engine._try_execute_rule(DummyItem("mail-1", ".msg"))
    assert ok is False
    assert handler_calls == []
    assert records == []


def test_run_async_mail_workflow_path_skips_engine_append_run(monkeypatch):
    calls = []
    records = []

    def fake_get_processor(_processor_id):
        def handler(item, output_dir, params):
            calls.append((item.name, output_dir, params))

        return handler

    monkeypatch.setattr(core, "get_processor", fake_get_processor)
    monkeypatch.setattr(core, "append_run", lambda record: records.append(record))

    config = _build_config(
        [
            AttachmentRule(
                extension=".msg",
                processor_id="mail_workflow",
                parameters={"max_concurrency": 2},
            )
        ]
    )
    adapter = DummyAdapter([DummyItem("mail-1", ".msg")])
    engine = GenericEtlEngine(config, adapter)

    engine.run()

    assert adapter.keywords == ["alpha"]
    assert calls == [("mail-1", "/tmp", {"max_concurrency": 2})]
    assert records == []
