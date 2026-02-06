import sys
import types
import importlib

from src.catalog.agents.external_api_agent import external_api_agent

agent_module = importlib.import_module("src.catalog.agents.external_api_agent")


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


def test_resolve_api_key_prefers_token_env(monkeypatch):
    monkeypatch.setenv("TOKEN_A", "key-a")
    monkeypatch.setenv("OPENAI_API_KEY", "default-key")

    assert agent_module._resolve_api_key("TOKEN_A") == "key-a"
    assert agent_module._resolve_api_key("") == "default-key"


def test_extract_input_text_fallback_order():
    item_with_raw_body = types.SimpleNamespace(
        name="mail.msg",
        _raw_item=types.SimpleNamespace(Body="Raw body", HTMLBody="Raw html"),
    )
    item_with_empty_raw_body = types.SimpleNamespace(
        name="mail.msg",
        _raw_item=types.SimpleNamespace(Body="", HTMLBody="<p>HTML</p>"),
    )
    item_with_raw_html = types.SimpleNamespace(
        name="mail.msg",
        _raw_item=types.SimpleNamespace(HTMLBody="<p>HTML</p>"),
    )
    plain_item = types.SimpleNamespace(name="mail.msg", _raw_item=None)

    assert agent_module._extract_input_text(item_with_raw_body, prefer_body=True) == "Raw body"
    assert agent_module._extract_input_text(item_with_empty_raw_body, prefer_body=True) == ""
    assert agent_module._extract_input_text(item_with_raw_html, prefer_body=True) == "<p>HTML</p>"
    assert agent_module._extract_input_text(plain_item, prefer_body=True) == "mail.msg"
    assert agent_module._extract_input_text(item_with_raw_body, prefer_body=False) == "mail.msg"


def test_external_api_agent_handles_sdk_import_failure(monkeypatch, capsys):
    monkeypatch.delitem(sys.modules, "openai", raising=False)

    external_api_agent(DummyItem(), "/tmp", {})

    captured = capsys.readouterr()
    assert "OpenAI SDK import failed" in captured.out


def test_external_api_agent_returns_when_api_key_missing(monkeypatch, capsys):
    class DummyOpenAI:
        def __init__(self, api_key):
            raise AssertionError("OpenAI should not be initialized when api key is missing")

    dummy_module = types.ModuleType("openai")
    dummy_module.OpenAI = DummyOpenAI
    monkeypatch.setitem(sys.modules, "openai", dummy_module)
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)

    external_api_agent(DummyItem(), "/tmp", {"token_env": "MISSING_KEY"})

    captured = capsys.readouterr()
    assert "API key not found" in captured.out


def test_external_api_agent_save_output_false_and_input_mode_name(monkeypatch, tmp_path):
    class DummyOpenAI:
        last_input = None

        def __init__(self, api_key):
            self.responses = self

        def create(self, model, input):
            DummyOpenAI.last_input = input
            return types.SimpleNamespace(output_text="ok")

    dummy_module = types.ModuleType("openai")
    dummy_module.OpenAI = DummyOpenAI
    monkeypatch.setitem(sys.modules, "openai", dummy_module)
    monkeypatch.setenv("OPENAI_API_KEY", "test-key")

    external_api_agent(
        DummyItem(),
        str(tmp_path),
        {
            "prompt": "Name only: {input_text}",
            "save_output": False,
            "input_mode": "name",
        },
    )

    assert DummyOpenAI.last_input == "Name only: dummy.msg"
    assert not (tmp_path / "dummy.msg.openai.txt").exists()


def test_external_api_agent_uses_stringified_response_without_output_text(monkeypatch, tmp_path):
    class DummyResponse:
        def __str__(self):
            return "stringified-response"

    class DummyOpenAI:
        def __init__(self, api_key):
            self.responses = self

        def create(self, model, input):
            return DummyResponse()

    dummy_module = types.ModuleType("openai")
    dummy_module.OpenAI = DummyOpenAI
    monkeypatch.setitem(sys.modules, "openai", dummy_module)
    monkeypatch.setenv("OPENAI_API_KEY", "test-key")

    external_api_agent(DummyItem(), str(tmp_path), {"save_output": True})

    out_file = tmp_path / "dummy.msg.openai.txt"
    assert out_file.exists()
    assert out_file.read_text(encoding="utf-8") == "stringified-response"


def test_external_api_agent_handles_openai_call_failure(monkeypatch, capsys):
    class DummyOpenAI:
        def __init__(self, api_key):
            self.responses = self

        def create(self, model, input):
            raise RuntimeError("api failure")

    dummy_module = types.ModuleType("openai")
    dummy_module.OpenAI = DummyOpenAI
    monkeypatch.setitem(sys.modules, "openai", dummy_module)
    monkeypatch.setenv("OPENAI_API_KEY", "test-key")

    external_api_agent(DummyItem(), "/tmp", {})

    captured = capsys.readouterr()
    assert "OpenAI call failed" in captured.out
