import json
import os

import pytest

import config


def test_config_overrides(tmp_path, monkeypatch):
    # create a temporary config file
    data = {
        "APP_TITLE": "Test Title",
        "CURRENT_THEME": "light",
        "CLIENT_EXCEL_PATH": "custom.xlsx",
    }
    cfg_file = tmp_path / "config.json"
    cfg_file.write_text(json.dumps(data))

    # load the test config file using the helper
    config.load_config(str(cfg_file))

    assert config.APP_TITLE == "Test Title"
    assert config.CURRENT_THEME == "light"
    assert "custom.xlsx" in config.CLIENT_EXCEL_PATH


def test_config_ignores_unknown_keys(tmp_path):
    data = {
        "APP_TITLE": "Allowed",
        "UNSAFE_KEY": "should_not_be_applied",
    }
    cfg_file = tmp_path / "config.json"
    cfg_file.write_text(json.dumps(data))

    config.load_config(str(cfg_file))

    assert config.APP_TITLE == "Allowed"
    assert not hasattr(config, "UNSAFE_KEY")


def test_pdf_footer_text_has_no_placeholder():
    assert "XXXX" not in config.PDF_FOOTER_TEXT
