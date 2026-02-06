import pandas as pd
import pytest

from src.bridge import excel_parser


def test_parse_excel_spec_raises_when_file_missing(tmp_path):
    missing_path = tmp_path / "missing.xlsx"

    with pytest.raises(FileNotFoundError):
        excel_parser.parse_excel_spec(str(missing_path))


def test_parse_excel_spec_wraps_settings_sheet_error(monkeypatch):
    monkeypatch.setattr(excel_parser.os.path, "exists", lambda _: True)

    def fake_read_excel(*args, **kwargs):
        if kwargs.get("sheet_name") == "Settings":
            raise RuntimeError("settings broken")
        raise AssertionError("Rules sheet should not be read after Settings failure")

    monkeypatch.setattr(excel_parser.pd, "read_excel", fake_read_excel)

    with pytest.raises(ValueError) as err:
        excel_parser.parse_excel_spec("dummy.xlsx")

    assert "Settingsシート" in str(err.value)


def test_parse_excel_spec_wraps_rules_sheet_error(monkeypatch):
    monkeypatch.setattr(excel_parser.os.path, "exists", lambda _: True)

    settings_df = pd.DataFrame(
        [
            ["Job Name", "Invoice Bot"],
            ["Keywords", "請求書, Invoice"],
        ]
    )

    def fake_read_excel(*args, **kwargs):
        sheet_name = kwargs.get("sheet_name")
        if sheet_name == "Settings":
            return settings_df
        if sheet_name == "Rules":
            raise RuntimeError("rules broken")
        raise AssertionError(f"unexpected sheet_name: {sheet_name}")

    monkeypatch.setattr(excel_parser.pd, "read_excel", fake_read_excel)

    with pytest.raises(ValueError) as err:
        excel_parser.parse_excel_spec("dummy.xlsx")

    assert "Rulesシート" in str(err.value)


def test_parse_excel_spec_parses_rules_and_sanitizes_parameters(monkeypatch):
    monkeypatch.setattr(excel_parser.os.path, "exists", lambda _: True)

    settings_df = pd.DataFrame(
        [
            ["Job Name", "Invoice Bot"],
            ["Keywords", "請求書, Invoice , ,"],
        ]
    )
    rules_df = pd.DataFrame(
        [
            {
                "Extension": ".pdf",
                "Processor ID": "pdf_to_text_ocr",
                "Parameters": '{"lang":"jpn"}',
            },
            {
                "Extension": ".msg",
                "Processor ID": "mail_workflow",
                "Parameters": "",
            },
            {
                "Extension": ".zip",
                "Processor ID": "unzip_file",
                "Parameters": '["not-a-dict"]',
            },
            {
                "Extension": ".xlsx",
                "Processor ID": "save_only",
                "Parameters": "{broken",
            },
            {
                "Extension": None,
                "Processor ID": "save_only",
                "Parameters": "{}",
            },
        ]
    )

    def fake_read_excel(*args, **kwargs):
        sheet_name = kwargs.get("sheet_name")
        if sheet_name == "Settings":
            return settings_df
        if sheet_name == "Rules":
            return rules_df
        raise AssertionError(f"unexpected sheet_name: {sheet_name}")

    monkeypatch.setattr(excel_parser.pd, "read_excel", fake_read_excel)

    config = excel_parser.parse_excel_spec("dummy.xlsx")

    assert config.job_name == "Invoice Bot"
    assert config.domain == "common"
    assert config.destination_path == "./data/output"
    assert config.search_keywords == ["請求書", "Invoice"]
    assert [rule.extension for rule in config.rules] == [".pdf", ".msg", ".zip", ".xlsx"]
    assert config.rules[0].parameters == {"lang": "jpn"}
    assert config.rules[1].parameters == {}
    assert config.rules[2].parameters == {}
    assert config.rules[3].parameters == {}
