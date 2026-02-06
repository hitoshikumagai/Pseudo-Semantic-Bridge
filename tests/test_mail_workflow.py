import json

from src.catalog.workflows.mail_router import mail_workflow


class DummyChild:
    def __init__(self, name: str):
        self.name = name


class DummyItem:
    def __init__(self, name: str, children=None):
        self._name = name
        self._children = children or []
        self.saved_to = None

    @property
    def name(self) -> str:
        return self._name

    def get_children(self):
        return self._children

    def save_to(self, directory: str):
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
