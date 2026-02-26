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
