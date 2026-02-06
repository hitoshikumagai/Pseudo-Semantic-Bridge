from typing import List

from src.catalog import register_processor
from src.engine.core import GenericEtlEngine
from src.schema.definitions import AttachmentRule, OutlookConfig


class DummyItem:
    def __init__(self, name: str, extension: str, children=None):
        self._name = name
        self._extension = extension
        self._children = children or []

    @property
    def name(self) -> str:
        return self._name

    @property
    def extension(self) -> str:
        return self._extension

    @property
    def is_container(self) -> bool:
        return len(self._children) > 0

    def get_children(self):
        return self._children


class DummyAdapter:
    def __init__(self, items: List[DummyItem]):
        self._items = items

    def fetch_items(self, keyword: str) -> List[DummyItem]:
        return self._items


def test_engine_runs_with_max_concurrency():
    calls = []

    @register_processor("test_handler")
    def _handler(item, output_dir, params):
        calls.append(item.name)

    config = OutlookConfig(
        job_name="test",
        domain="test",
        search_keywords=["x"],
        destination_path="/tmp",
        rules=[
            AttachmentRule(
                extension=".msg",
                processor_id="test_handler",
                parameters={"max_concurrency": 2},
            )
        ],
    )

    engine = GenericEtlEngine(config, DummyAdapter([DummyItem("mail-1", ".msg")]))
    engine.run()

    assert calls == ["mail-1"]


def test_engine_recurses_into_container_when_no_rule_on_container():
    calls = []

    @register_processor("test_handler_container")
    def _handler(item, output_dir, params):
        calls.append(item.name)

    child = DummyItem("child-1", ".msg")
    container = DummyItem("container", ".none", children=[child])

    config = OutlookConfig(
        job_name="test",
        domain="test",
        search_keywords=["x"],
        destination_path="/tmp",
        rules=[
            AttachmentRule(
                extension=".msg",
                processor_id="test_handler_container",
                parameters={},
            )
        ],
    )

    engine = GenericEtlEngine(config, DummyAdapter([container]))
    engine.run()

    assert calls == ["child-1"]


def test_engine_async_handler_error_does_not_raise():
    @register_processor("test_handler_async_error")
    def _handler(item, output_dir, params):
        raise RuntimeError("boom")

    config = OutlookConfig(
        job_name="test",
        domain="test",
        search_keywords=["x"],
        destination_path="/tmp",
        rules=[
            AttachmentRule(
                extension=".msg",
                processor_id="test_handler_async_error",
                parameters={"max_concurrency": 2},
            )
        ],
    )

    engine = GenericEtlEngine(config, DummyAdapter([DummyItem("mail-1", ".msg")]))
    engine.run()


def test_engine_sync_handler_error_does_not_raise():
    @register_processor("test_handler_sync_error")
    def _handler(item, output_dir, params):
        raise RuntimeError("boom")

    config = OutlookConfig(
        job_name="test",
        domain="test",
        search_keywords=["x"],
        destination_path="/tmp",
        rules=[
            AttachmentRule(
                extension=".msg",
                processor_id="test_handler_sync_error",
                parameters={"max_concurrency": 1},
            )
        ],
    )

    engine = GenericEtlEngine(config, DummyAdapter([DummyItem("mail-1", ".msg")]))
    engine.run()
