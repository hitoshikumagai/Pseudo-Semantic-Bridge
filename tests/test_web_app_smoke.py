import importlib
import sys
import types


class _DummyContext:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


class _FakeStreamlit(types.ModuleType):
    def __init__(self, pressed_buttons=None):
        super().__init__("streamlit")
        self.session_state = {}
        self.pressed_buttons = set(pressed_buttons or [])
        self.calls = []

    def set_page_config(self, **kwargs):
        self.calls.append(("set_page_config", kwargs))

    def markdown(self, *args, **kwargs):
        self.calls.append(("markdown", args, kwargs))

    def write(self, *args, **kwargs):
        self.calls.append(("write", args, kwargs))

    def tabs(self, labels):
        self.calls.append(("tabs", list(labels)))
        return [_DummyContext() for _ in labels]

    def columns(self, specs):
        self.calls.append(("columns", list(specs)))
        return [_DummyContext() for _ in specs]

    def data_editor(self, data, key=None, **kwargs):
        self.calls.append(("data_editor", key))
        if key == "ai_candidate_editor" and isinstance(data, list) and data:
            edited = [dict(row) for row in data]
            edited[0]["select"] = True
            return edited
        return data

    def button(self, label, **kwargs):
        self.calls.append(("button", label))
        return label in self.pressed_buttons

    def number_input(self, label, **kwargs):
        self.calls.append(("number_input", label))
        return kwargs.get("value", 0)

    def text_area(self, label, **kwargs):
        self.calls.append(("text_area", label))
        return kwargs.get("value", "")

    def text_input(self, label, **kwargs):
        self.calls.append(("text_input", label))
        return kwargs.get("value", "")

    def radio(self, label, options, **kwargs):
        self.calls.append(("radio", label))
        return options[0]

    def caption(self, *args, **kwargs):
        self.calls.append(("caption", args, kwargs))

    def dataframe(self, *args, **kwargs):
        self.calls.append(("dataframe", args, kwargs))

    def success(self, *args, **kwargs):
        self.calls.append(("success", args, kwargs))

    def info(self, *args, **kwargs):
        self.calls.append(("info", args, kwargs))

    def warning(self, *args, **kwargs):
        self.calls.append(("warning", args, kwargs))

    def json(self, *args, **kwargs):
        self.calls.append(("json", args, kwargs))

    def experimental_rerun(self):
        self.calls.append(("experimental_rerun",))


def _import_web_app_with_fakes(monkeypatch, pressed_buttons=None, runs=None):
    fake_st = _FakeStreamlit(pressed_buttons=pressed_buttons)
    monkeypatch.setitem(sys.modules, "streamlit", fake_st)

    fake_builder = types.ModuleType("src.bridge.builder")
    fake_builder.build_all_configs = lambda: None
    monkeypatch.setitem(sys.modules, "src.bridge.builder", fake_builder)

    fake_engine = types.ModuleType("src.engine.core")

    class GenericEtlEngine:
        pass

    fake_engine.GenericEtlEngine = GenericEtlEngine
    monkeypatch.setitem(sys.modules, "src.engine.core", fake_engine)

    fake_adapter = types.ModuleType("src.adapter.outlook")

    class OutlookAdapter:
        pass

    fake_adapter.OutlookAdapter = OutlookAdapter
    monkeypatch.setitem(sys.modules, "src.adapter.outlook", fake_adapter)

    tracker = {"save_rules": [], "run_engine_job": [], "start_job": 0}
    fake_app_logic = types.ModuleType("src.web.app_logic")
    fake_app_logic.load_rules = lambda _path: []
    fake_app_logic.load_jsonl_runs = lambda _path: list(runs or [])
    fake_app_logic.propose_rule_candidates = (
        lambda _runs, min_samples, min_quality_rate: (
            [{"subject_filter": "Invoice", "quality_rate": 0.9}],
            [
                {
                    "subject_filter": "Invoice",
                    "task_name": "AUTO",
                    "require_attachment": True,
                    "target_ext": ".pdf",
                    "action_id": "ocr_process",
                    "parameters": {},
                }
            ],
        )
    )

    def save_rules(path, rules):
        tracker["save_rules"].append((path, list(rules)))

    def run_engine_job(jobs, job_id, build_fn, config_path, adapter_factory, engine_factory):
        tracker["run_engine_job"].append(
            (job_id, build_fn, config_path, adapter_factory, engine_factory)
        )
        jobs[job_id]["status"] = "done"

    def start_job(jobs, run_fn):
        tracker["start_job"] += 1
        job_id = "job-smoke"
        jobs[job_id] = {"status": "queued"}
        run_fn(job_id)
        return job_id

    fake_app_logic.save_rules = save_rules
    fake_app_logic.run_engine_job = run_engine_job
    fake_app_logic.start_job = start_job
    fake_app_logic.summarize_quality = (
        lambda _runs: {"total": 1, "success": 1, "quality_labeled": 1, "quality_ok": 1}
    )
    monkeypatch.setitem(sys.modules, "src.web.app_logic", fake_app_logic)

    monkeypatch.delitem(sys.modules, "web_app", raising=False)
    module = importlib.import_module("web_app")
    return module, fake_st, tracker


def test_web_app_import_smoke(monkeypatch):
    module, fake_st, tracker = _import_web_app_with_fakes(monkeypatch, pressed_buttons=set(), runs=[])

    assert module.APP_TITLE == "Pseudo Semantic Bridge"
    assert "rules" in fake_st.session_state
    assert fake_st.session_state["rules"][0]["action_id"] == "ocr_process"
    assert tracker["save_rules"] == []
    assert tracker["run_engine_job"] == []
    assert any(call[0] == "set_page_config" for call in fake_st.calls)


def test_web_app_major_button_flows_smoke(monkeypatch):
    pressed = {
        "Save Rules",
        "Generate Candidates",
        "Append Selected To Rules",
        "受付内容を保存",
        "AIで整理（準備）",
        "Run Pipeline",
    }
    runs = [{"timestamp": "2026-02-01T00:00:00", "result": {"status": "success"}}]
    module, fake_st, tracker = _import_web_app_with_fakes(monkeypatch, pressed_buttons=pressed, runs=runs)

    assert tracker["save_rules"]
    assert tracker["start_job"] == 1
    assert tracker["run_engine_job"]
    assert fake_st.session_state["last_job_id"] == "job-smoke"
    assert fake_st.session_state["jobs"]["job-smoke"]["status"] == "done"
    assert fake_st.session_state["quality_intake_ready"] is True
    assert fake_st.session_state["quality_intake"]["app_context"] == "メール"
    assert any(rule.get("task_name") == "AUTO" for rule in fake_st.session_state["rules"])
    assert module.LOGS_PATH.as_posix() == "data/logs/psb_run.jsonl"
