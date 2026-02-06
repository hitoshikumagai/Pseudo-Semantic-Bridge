import sys
import types

from src.catalog.agents.external_api_agent import external_api_agent


class DummyItem:
    name = "dummy.msg"
    extension = ".msg"
    body = "Hello from body"


def test_external_api_agent_openai_sdk(monkeypatch, tmp_path):
    class DummyOpenAI:
        last_input = None

        def __init__(self, api_key):
            self.api_key = api_key
            self.responses = self

        def create(self, model, input):
            DummyOpenAI.last_input = input
            return types.SimpleNamespace(output_text="ok")

    dummy_module = types.ModuleType("openai")
    dummy_module.OpenAI = DummyOpenAI
    sys.modules["openai"] = dummy_module

    monkeypatch.setenv("OPENAI_API_KEY", "test-key")

    params = {
        "model": "gpt-4.1",
        "prompt": "Summarize: {input_text}",
        "save_output": True,
        "input_mode": "body",
    }
    external_api_agent(DummyItem(), str(tmp_path), params)

    out_file = tmp_path / "dummy.msg.openai.txt"
    assert out_file.exists()
    assert out_file.read_text(encoding="utf-8") == "ok"
    assert DummyOpenAI.last_input == "Summarize: Hello from body"
