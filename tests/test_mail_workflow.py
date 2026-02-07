import json

from src.catalog.workflows.mail_router import mail_workflow


class DummyChild:
    def __init__(self, name: str, extension: str = ""):
        self.name = name
        self.extension = extension


class DummyItem:
    def __init__(self, name: str, children=None, fail_on_save: bool = False):
        self._name = name
        self._children = children or []
        self.saved_to = None
        self._fail_on_save = fail_on_save

    @property
    def name(self) -> str:
        return self._name

    def get_children(self):
        return self._children

    def save_to(self, directory: str):
        if self._fail_on_save:
            raise RuntimeError("save failed")
        self.saved_to = directory


def _write_rules(path, rules):
    path.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")


def test_mail_workflow_rule_match_and_saves_body(tmp_path):
    rules_path = tmp_path / "rules.json"
    _write_rules(
        rules_path,
        [
            {
                "subject_filter": "Report",
                "task_name": "REPORT",
                "action_id": "save_process",
                "require_attachment": False,
            }
        ],
    )

    item = DummyItem("Daily Report")
    mail_workflow(item, str(tmp_path), {"rule_file": str(rules_path)})

    assert item.saved_to == str(tmp_path / "REPORT")


def test_mail_workflow_resolves_relative_rule_path_from_cwd(tmp_path, monkeypatch):
    rules_path = tmp_path / "rules.json"
    _write_rules(
        rules_path,
        [
            {
                "subject_filter": "Report",
                "task_name": "REPORT",
                "action_id": "save_process",
                "require_attachment": False,
            }
        ],
    )

    monkeypatch.chdir(tmp_path)
    item = DummyItem("Daily Report")
    mail_workflow(item, str(tmp_path), {"rule_file": "rules.json"})

    assert item.saved_to == str(tmp_path / "REPORT")


def test_mail_workflow_requires_attachment_skips(tmp_path):
    rules_path = tmp_path / "rules.json"
    _write_rules(
        rules_path,
        [
            {
                "subject_filter": "Invoice",
                "task_name": "INVOICE",
                "action_id": "save_process",
                "require_attachment": True,
            }
        ],
    )

    item = DummyItem("Invoice 123")
    mail_workflow(item, str(tmp_path), {"rule_file": str(rules_path)})

    assert item.saved_to is None


def test_mail_workflow_with_attachments_calls_handler(tmp_path, monkeypatch):
    rules_path = tmp_path / "rules.json"
    _write_rules(
        rules_path,
        [
            {
                "subject_filter": "Invoice",
                "task_name": "INVOICE",
                "action_id": "save_process",
                "require_attachment": True,
            }
        ],
    )

    calls = []

    def fake_handler(item, output_dir, params):
        calls.append((item.name, output_dir))

    monkeypatch.setitem(
        mail_workflow.__globals__["PROCESSOR_MAP"],
        "save_process",
        fake_handler,
    )

    item = DummyItem("Invoice 123", children=[DummyChild("a.pdf")])
    mail_workflow(item, str(tmp_path), {"rule_file": str(rules_path)})

    assert calls == [("a.pdf", str(tmp_path / "INVOICE"))]


def test_mail_workflow_missing_rule_file_returns_early(monkeypatch):
    logs = []
    monkeypatch.setattr("builtins.print", lambda *args, **kwargs: logs.append(" ".join(map(str, args))))

    item = DummyItem("Any Subject")
    mail_workflow(item, "/tmp/out", {"rule_file": "/tmp/not-found.json"})

    assert item.saved_to is None
    assert any("Rule file not found" in line for line in logs)


def test_mail_workflow_invalid_rule_json_returns_early(tmp_path, monkeypatch):
    logs = []
    monkeypatch.setattr("builtins.print", lambda *args, **kwargs: logs.append(" ".join(map(str, args))))

    rules_path = tmp_path / "rules.json"
    rules_path.write_text("{invalid-json", encoding="utf-8")
    item = DummyItem("Any Subject")
    mail_workflow(item, str(tmp_path), {"rule_file": str(rules_path)})

    assert item.saved_to is None
    assert any("Rule load error" in line for line in logs)


def test_mail_workflow_no_matching_rule_does_nothing(tmp_path):
    rules_path = tmp_path / "rules.json"
    _write_rules(
        rules_path,
        [{"subject_filter": "Invoice", "task_name": "INVOICE", "action_id": "save_process"}],
    )

    item = DummyItem("Daily Report")
    mail_workflow(item, str(tmp_path), {"rule_file": str(rules_path)})

    assert item.saved_to is None


def test_mail_workflow_unknown_action_falls_back_to_save_only_and_logs(tmp_path, monkeypatch):
    rules_path = tmp_path / "rules.json"
    _write_rules(
        rules_path,
        [
            {
                "subject_filter": "Invoice",
                "task_name": "INVOICE",
                "action_id": "unknown_action",
                "require_attachment": True,
            }
        ],
    )

    calls = []
    records = []

    def fake_save_only(item, output_dir, params):
        calls.append((item.name, output_dir, params))

    monkeypatch.setitem(mail_workflow.__globals__, "save_only", fake_save_only)
    monkeypatch.setitem(mail_workflow.__globals__, "append_run", lambda record: records.append(record))

    item = DummyItem("Invoice 123", children=[DummyChild("a.PDF")])
    mail_workflow(item, str(tmp_path), {"rule_file": str(rules_path)})

    assert calls == [("a.PDF", str(tmp_path / "INVOICE"), {"rule_file": str(rules_path)})]
    assert len(records) == 1
    assert records[0]["result"]["status"] == "success"
    assert records[0]["input"]["attachment_ext"] == ".pdf"
    assert records[0]["result"]["action_executed"] == "unknown_action"


def test_mail_workflow_merges_rule_parameters(tmp_path, monkeypatch):
    rules_path = tmp_path / "rules.json"
    _write_rules(
        rules_path,
        [
            {
                "subject_filter": "Invoice",
                "task_name": "INVOICE",
                "action_id": "save_process",
                "require_attachment": True,
                "parameters": {"lang": "jpn"},
            }
        ],
    )

    calls = []

    def fake_handler(item, output_dir, params):
        calls.append(params)

    monkeypatch.setitem(mail_workflow.__globals__["PROCESSOR_MAP"], "save_process", fake_handler)

    item = DummyItem("Invoice 123", children=[DummyChild("a.pdf")])
    mail_workflow(item, str(tmp_path), {"rule_file": str(rules_path)})

    assert calls
    assert calls[0]["lang"] == "jpn"
    assert calls[0]["rule_file"] == str(rules_path)


def test_mail_workflow_attachment_error_appends_error_record(tmp_path, monkeypatch):
    rules_path = tmp_path / "rules.json"
    _write_rules(
        rules_path,
        [
            {
                "subject_filter": "Invoice",
                "task_name": "INVOICE",
                "action_id": "save_process",
                "require_attachment": "true",
            }
        ],
    )

    records = []

    def fake_handler(item, output_dir, params):
        raise RuntimeError("handler failed")

    monkeypatch.setitem(mail_workflow.__globals__["PROCESSOR_MAP"], "save_process", fake_handler)
    monkeypatch.setitem(mail_workflow.__globals__, "append_run", lambda record: records.append(record))

    item = DummyItem("Invoice 123", children=[DummyChild("a.pdf")])
    mail_workflow(item, str(tmp_path), {"rule_file": str(rules_path)})

    assert len(records) == 1
    assert records[0]["result"]["status"] == "error"
    assert "handler failed" in records[0]["result"]["error"]
    assert records[0]["input"]["attachment_ext"] == ".pdf"


def test_mail_workflow_body_save_error_appends_error_record(tmp_path, monkeypatch):
    rules_path = tmp_path / "rules.json"
    _write_rules(
        rules_path,
        [
            {
                "subject_filter": "Report",
                "task_name": "REPORT",
                "action_id": "save_process",
                "require_attachment": False,
            }
        ],
    )

    records = []
    monkeypatch.setitem(mail_workflow.__globals__, "append_run", lambda record: records.append(record))

    item = DummyItem("Daily Report", fail_on_save=True)
    mail_workflow(item, str(tmp_path), {"rule_file": str(rules_path)})

    assert len(records) == 1
    assert records[0]["result"]["status"] == "error"
    assert "save failed" in records[0]["result"]["error"]
    assert records[0]["input"]["attachment_ext"] == ""
