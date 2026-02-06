import json
from pathlib import Path

import pytest

from src.schema.definitions import AttachmentRule, IntentSpecification


def test_attachment_rule_accepts_custom_processor_id():
    rule = AttachmentRule(
        extension=".msg",
        processor_id="agent_external_api",
        parameters={"x": 1},
    )
    assert rule.processor_id == "agent_external_api"


def test_intent_specification_accepts_accounting_sample():
    sample_path = Path("specs/accounting/invoice_intent_spec.sample.json")
    data = json.loads(sample_path.read_text(encoding="utf-8"))
    spec = IntentSpecification(**data)
    assert spec.domain == "accounting_mail_invoice"
    assert spec.steps[0].action == "fetch_mails"


def test_intent_specification_rejects_duplicate_step_ids():
    payload = {
        "spec_id": "spec-1",
        "spec_version": "1.0",
        "domain": "accounting_mail_invoice",
        "intent": "test",
        "steps": [
            {"id": "s1", "action": "fetch_mails"},
            {"id": "s1", "action": "ocr_process"},
        ],
        "verification": {"required_fields": [], "min_quality_score": 0.8},
        "fallback": {"on_failure": "route_manual_review"},
    }

    with pytest.raises(ValueError, match="step ids must be unique"):
        IntentSpecification(**payload)
