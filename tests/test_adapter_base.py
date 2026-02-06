import pytest

from src.adapter.base import BaseAdapter, UnifiedItem


class ConcreteItem(UnifiedItem):
    @property
    def name(self) -> str:
        return "item"

    @property
    def extension(self) -> str:
        return ".txt"

    @property
    def is_container(self) -> bool:
        return False

    def get_children(self):
        return []

    def save_to(self, directory: str) -> str:
        return f"{directory}/item.txt"


class ConcreteAdapter(BaseAdapter):
    def fetch_items(self, keyword: str):
        return [ConcreteItem({"keyword": keyword})]


def test_unified_item_is_abstract():
    with pytest.raises(TypeError):
        UnifiedItem(raw_item={})


def test_base_adapter_is_abstract():
    with pytest.raises(TypeError):
        BaseAdapter()


def test_concrete_item_and_adapter_behave_as_expected():
    adapter = ConcreteAdapter()
    items = adapter.fetch_items("invoice")

    assert len(items) == 1
    assert items[0].name == "item"
    assert items[0].extension == ".txt"
    assert items[0].is_container is False
    assert items[0].get_children() == []
    assert items[0].save_to("/tmp") == "/tmp/item.txt"
    assert items[0]._raw_item == {"keyword": "invoice"}
