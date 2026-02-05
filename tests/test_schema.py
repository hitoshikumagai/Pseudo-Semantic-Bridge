from src.schema.definitions import AttachmentRule


def test_attachment_rule_accepts_custom_processor_id():
    rule = AttachmentRule(
        extension=".msg",
        processor_id="agent_external_api",
        parameters={"x": 1},
    )
    assert rule.processor_id == "agent_external_api"
