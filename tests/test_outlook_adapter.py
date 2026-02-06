import importlib
import sys
import types
from pathlib import Path

import pytest


def _load_outlook_module(monkeypatch):
    client_mod = types.ModuleType("win32com.client")
    client_mod.Dispatch = lambda *_args, **_kwargs: None

    win32com_mod = types.ModuleType("win32com")
    win32com_mod.client = client_mod

    monkeypatch.setitem(sys.modules, "win32com", win32com_mod)
    monkeypatch.setitem(sys.modules, "win32com.client", client_mod)

    import src.adapter.outlook as outlook

    return importlib.reload(outlook)


def test_outlook_item_name_prefers_subject_then_filename_then_unknown(monkeypatch):
    outlook = _load_outlook_module(monkeypatch)

    assert outlook.OutlookItem(types.SimpleNamespace(Subject="Mail", FileName="a.txt")).name == "Mail"
    assert outlook.OutlookItem(types.SimpleNamespace(FileName="a.txt")).name == "a.txt"
    assert outlook.OutlookItem(types.SimpleNamespace()).name == "Unknown"


def test_outlook_item_extension_container_and_attachment(monkeypatch):
    outlook = _load_outlook_module(monkeypatch)

    container = outlook.OutlookItem(types.SimpleNamespace(Subject="Mail", Attachments=[]))
    attachment = outlook.OutlookItem(types.SimpleNamespace(FileName="A.PDF"))

    assert container.extension == ".msg"
    assert attachment.extension == ".pdf"


def test_outlook_item_is_container_requires_subject_and_attachments(monkeypatch):
    outlook = _load_outlook_module(monkeypatch)

    assert outlook.OutlookItem(types.SimpleNamespace(Subject="Mail", Attachments=[])).is_container is True
    assert outlook.OutlookItem(types.SimpleNamespace(Subject="Mail")).is_container is False
    assert outlook.OutlookItem(types.SimpleNamespace(Attachments=[])).is_container is False


def test_outlook_item_get_children_wraps_attachments(monkeypatch):
    outlook = _load_outlook_module(monkeypatch)

    raw = types.SimpleNamespace(
        Subject="Mail",
        Attachments=[types.SimpleNamespace(FileName="a.pdf"), types.SimpleNamespace(FileName="b.zip")],
    )
    item = outlook.OutlookItem(raw)

    children = item.get_children()
    assert len(children) == 2
    assert all(isinstance(child, outlook.OutlookItem) for child in children)


def test_outlook_item_get_children_handles_iteration_error(monkeypatch):
    outlook = _load_outlook_module(monkeypatch)

    class BrokenAttachments:
        def __iter__(self):
            raise RuntimeError("boom")

    raw = types.SimpleNamespace(Subject="Mail", Attachments=BrokenAttachments())
    item = outlook.OutlookItem(raw)

    assert item.get_children() == []


def test_outlook_item_save_to_container_uses_saveas_with_sanitized_name(monkeypatch, tmp_path):
    outlook = _load_outlook_module(monkeypatch)

    calls = []

    def save_as(path, fmt):
        calls.append((path, fmt))

    raw = types.SimpleNamespace(Subject='mail:report?.msg', Attachments=[], SaveAs=save_as)
    item = outlook.OutlookItem(raw)

    result = item.save_to(str(tmp_path))
    result_path = Path(result)

    assert result_path.name == "mail_report_.msg"
    assert calls == [(str(result_path), 3)]


def test_outlook_item_save_to_attachment_uses_save_as_wrapper(monkeypatch, tmp_path):
    outlook = _load_outlook_module(monkeypatch)

    temp_dir = tmp_path / "temp"
    temp_dir.mkdir()
    monkeypatch.setattr(outlook.tempfile, "gettempdir", lambda: str(temp_dir))

    def save_as(path):
        Path(path).write_text("content", encoding="utf-8")

    raw = types.SimpleNamespace(FileName="doc.txt", save_as=save_as)
    item = outlook.OutlookItem(raw)

    result = item.save_to(str(tmp_path / "out"))
    assert Path(result).read_text(encoding="utf-8") == "content"


def test_outlook_item_save_to_attachment_falls_back_to_saveasfile(monkeypatch, tmp_path):
    outlook = _load_outlook_module(monkeypatch)

    temp_dir = tmp_path / "temp"
    temp_dir.mkdir()
    monkeypatch.setattr(outlook.tempfile, "gettempdir", lambda: str(temp_dir))

    def save_as_file(path):
        Path(path).write_text("from-com", encoding="utf-8")

    raw = types.SimpleNamespace(FileName="doc2.txt", SaveAsFile=save_as_file)
    item = outlook.OutlookItem(raw)

    result = item.save_to(str(tmp_path / "out"))
    assert Path(result).read_text(encoding="utf-8") == "from-com"


def test_outlook_item_save_to_attachment_raises_when_no_save_method(monkeypatch, tmp_path):
    outlook = _load_outlook_module(monkeypatch)
    raw = types.SimpleNamespace(FileName="doc3.txt")
    item = outlook.OutlookItem(raw)

    with pytest.raises(Exception):
        item.save_to(str(tmp_path / "out"))


def test_outlook_adapter_fetch_items_success(monkeypatch):
    outlook = _load_outlook_module(monkeypatch)

    raw_items = [types.SimpleNamespace(Subject="Mail-1", Attachments=[])]

    class FakeItems:
        def Restrict(self, query):
            assert "subject" in query
            return raw_items

    class FakeFolder:
        Items = FakeItems()

    class FakeNamespace:
        def GetDefaultFolder(self, folder_id):
            assert folder_id == 6
            return FakeFolder()

    class FakeApplication:
        def GetNamespace(self, name):
            assert name == "MAPI"
            return FakeNamespace()

    monkeypatch.setattr(outlook.win32com.client, "Dispatch", lambda app: FakeApplication())

    adapter = outlook.OutlookAdapter()
    items = adapter.fetch_items("Invoice")

    assert len(items) == 1
    assert items[0].name == "Mail-1"


def test_outlook_adapter_init_failure_and_search_failure_paths(monkeypatch):
    outlook = _load_outlook_module(monkeypatch)

    monkeypatch.setattr(outlook.win32com.client, "Dispatch", lambda app: (_ for _ in ()).throw(RuntimeError("connect fail")))
    adapter = outlook.OutlookAdapter()
    assert adapter.fetch_items("x") == []

    class BrokenItems:
        def Restrict(self, _query):
            raise RuntimeError("search fail")

    class BrokenFolder:
        Items = BrokenItems()

    class FakeNamespace:
        def GetDefaultFolder(self, _folder_id):
            return BrokenFolder()

    adapter.outlook = FakeNamespace()
    assert adapter.fetch_items("x") == []
