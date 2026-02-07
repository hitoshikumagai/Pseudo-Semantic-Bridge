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
        self.calls = []
        self.pressed_buttons = set(pressed_buttons or [])

    def set_page_config(self, **kwargs):
        self.calls.append(("set_page_config", kwargs))

    def markdown(self, *args, **kwargs):
        self.calls.append(("markdown",))

    def write(self, *args, **kwargs):
        self.calls.append(("write",))

    def tabs(self, labels):
        self.calls.append(("tabs", list(labels)))
        return [_DummyContext() for _ in labels]

    def columns(self, specs):
        if isinstance(specs, int):
            count = specs
        else:
            count = len(specs)
        self.calls.append(("columns", count))
        return [_DummyContext() for _ in range(count)]

    def data_editor(self, data, **kwargs):
        self.calls.append(("data_editor", kwargs.get("key")))
        return data

    def button(self, label, **kwargs):
        self.calls.append(("button", label))
        return label in self.pressed_buttons

    def number_input(self, label, **kwargs):
        return kwargs.get("value", 0)

    def text_area(self, label, **kwargs):
        return kwargs.get("value", "")

    def text_input(self, label, **kwargs):
        return kwargs.get("value", "")

    def checkbox(self, label, **kwargs):
        return kwargs.get("value", False)

    def radio(self, label, options, **kwargs):
        return options[0]

    def caption(self, *args, **kwargs):
        self.calls.append(("caption",))

    def metric(self, *args, **kwargs):
        self.calls.append(("metric",))

    def dataframe(self, *args, **kwargs):
        self.calls.append(("dataframe",))

    def success(self, *args, **kwargs):
        self.calls.append(("success",))

    def info(self, *args, **kwargs):
        self.calls.append(("info",))

    def warning(self, *args, **kwargs):
        self.calls.append(("warning",))

    def json(self, *args, **kwargs):
        self.calls.append(("json",))

    def rerun(self):
        self.calls.append(("rerun",))


def _import_web_app_with_fakes(monkeypatch, pressed_buttons=None):
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

    tracker = {"run_engine_job": 0, "start_job": 0}
    fake_app_logic = types.ModuleType("src.web.app_logic")
    fake_app_logic.load_rules = lambda _path: []
    fake_app_logic.load_jsonl_runs = lambda _path: []
    fake_app_logic.load_jsonl_runs_tail = lambda _path, max_lines=200: []
    fake_app_logic.propose_rule_candidates = lambda *_args, **_kwargs: ([], [])
    fake_app_logic.save_rules = lambda *_args, **_kwargs: None
    fake_app_logic.analyze_and_log_user_instruction = (
        lambda **_kwargs: (
            {
                "record_type": "instruction_intake",
                "source": "template",
                "instruction_raw": "smoke",
                "intent_summary": "smoke",
                "tasks": [],
                "constraints": [],
                "missing_info": [],
                "follow_up_questions": [],
            },
            None,
            "template",
        )
    )
    fake_app_logic.run_bridge_compile_summary = (
        lambda *_args, **_kwargs: {
            "status": "done",
            "error": None,
            "started_at": "2026-02-06T00:00:00+00:00",
            "ended_at": "2026-02-06T00:00:01+00:00",
            "artifacts": [],
        }
    )
    fake_app_logic.compute_job_duration_seconds = lambda _job: 0.1
    fake_app_logic.summarize_run_window = (
        lambda runs, start_index=0: {
            "total": max(len(runs) - start_index, 0),
            "success": 0,
            "error": 0,
            "with_output": 0,
            "latest_error": None,
            "latest_timestamp": None,
            "workflows": [],
        }
    )
    fake_app_logic.summarize_run_detail_rows = lambda *_args, **_kwargs: []
    fake_app_logic.generate_intent_spec = (
        lambda **_kwargs: (
            {
                "spec_id": "spec-smoke",
                "spec_version": "1.0",
                "domain": "accounting_mail_invoice",
                "intent": "smoke",
                "inputs": {},
                "steps": [{"id": "s1", "action": "fetch_mails", "params": {}}],
                "verification": {"required_fields": [], "min_quality_score": 0.8},
                "fallback": {"on_failure": "route_manual_review"},
            },
            None,
            "template",
        )
    )

    def run_engine_job(jobs, job_id, build_fn, config_path, adapter_factory, engine_factory):
        tracker["run_engine_job"] += 1
        jobs[job_id]["status"] = "done"

    def start_job(jobs, run_fn):
        tracker["start_job"] += 1
        job_id = "job-smoke"
        jobs[job_id] = {"status": "queued"}
        run_fn(job_id)
        return job_id

    fake_app_logic.run_engine_job = run_engine_job
    fake_app_logic.start_job = start_job
    fake_app_logic.summarize_quality = (
        lambda _runs: {"total": 0, "success": 0, "quality_labeled": 0, "quality_ok": 0}
    )
    monkeypatch.setitem(sys.modules, "src.web.app_logic", fake_app_logic)

    monkeypatch.delitem(sys.modules, "web_app", raising=False)
    module = importlib.import_module("web_app")
    return module, fake_st, tracker


def test_web_app_import_smoke(monkeypatch):
    module, fake_st, _tracker = _import_web_app_with_fakes(monkeypatch, pressed_buttons=set())
    assert module.APP_TITLE == "Pseudo Semantic Bridge"
    assert "rules" in fake_st.session_state
    assert any(call[0] == "tabs" for call in fake_st.calls)


def test_web_app_run_pipeline_smoke(monkeypatch):
    _module, fake_st, tracker = _import_web_app_with_fakes(
        monkeypatch,
        pressed_buttons={"Run Pipeline"},
    )
    assert tracker["start_job"] == 1
    assert tracker["run_engine_job"] == 1
    assert fake_st.session_state["last_job_id"] == "job-smoke"
    assert fake_st.session_state["jobs"]["job-smoke"]["status"] == "done"
