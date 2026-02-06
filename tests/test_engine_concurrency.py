from typing import List

from src.catalog import register_processor
from src.engine.core import GenericEtlEngine
from src.schema.definitions import AttachmentRule, OutlookConfig


class DummyItem:
    def __init__(self, name: str, extension: str):
        self._name = name
        self._extension = extension

    @property
    def name(self) -> str:
        return self._name

    @property
    def extension(self) -> str:
        return self._extension

    @property
    def is_container(self) -> bool:
        return False

    def get_children(self):
        return []


class DummyAdapter:
    def fetch_items(self, keyword: str) -> List[DummyItem]:
        return [DummyItem("mail-1", ".msg")]


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

    engine = GenericEtlEngine(config, DummyAdapter())
    engine.run()

    assert calls == ["mail-1"]
