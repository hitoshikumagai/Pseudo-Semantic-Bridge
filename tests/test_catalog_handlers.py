from src.catalog.handlers import archive, basic, document


class DummyItem:
    def __init__(self, save_path="/tmp/file.txt", fail=False):
        self._save_path = save_path
        self._fail = fail
        self.calls = []
        self.name = "dummy.pdf"

    def save_to(self, output_dir):
        self.calls.append(output_dir)
        if self._fail:
            raise RuntimeError("save failed")
        return self._save_path


def test_save_only_calls_save_to():
    item = DummyItem()

    basic.save_only(item, "/tmp/output")

    assert item.calls == ["/tmp/output"]


def test_save_only_handles_save_error_without_raising(monkeypatch):
    logs = []
    monkeypatch.setattr("builtins.print", lambda *args, **kwargs: logs.append(" ".join(map(str, args))))

    item = DummyItem(fail=True)
    basic.save_only(item, "/tmp/output")

    assert any("Save Error" in line for line in logs)


def test_unzip_file_reads_mode_from_kwargs_params():
    item = DummyItem(save_path="/tmp/output/a.zip")

    archive.unzip_file(item, "/tmp/output", params={"mode": "manual"})

    assert item.calls == ["/tmp/output"]


def test_unzip_file_handles_save_error_without_raising(monkeypatch):
    logs = []
    monkeypatch.setattr("builtins.print", lambda *args, **kwargs: logs.append(" ".join(map(str, args))))

    item = DummyItem(fail=True)
    archive.unzip_file(item, "/tmp/output", {"mode": "auto"})

    assert any("Zip Error" in line for line in logs)


def test_pdf_to_text_ocr_uses_args_params_and_calls_save_to():
    item = DummyItem(save_path="/tmp/output/a.pdf")

    document.pdf_to_text_ocr(item, "/tmp/output", {"lang": "jpn"})

    assert item.calls == ["/tmp/output"]


def test_pdf_to_text_ocr_handles_save_error_without_raising(monkeypatch):
    logs = []
    monkeypatch.setattr("builtins.print", lambda *args, **kwargs: logs.append(" ".join(map(str, args))))

    item = DummyItem(fail=True)
    document.pdf_to_text_ocr(item, "/tmp/output", params={"lang": "eng"})

    assert any("OCR Error" in line for line in logs)
