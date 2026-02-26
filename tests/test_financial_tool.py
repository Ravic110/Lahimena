# verify that one of the core modules is importable once the project is installed
import importlib
import os
import subprocess
import sys

import pytest


def test_core_module_importable():
    spec = importlib.util.find_spec("gui")
    assert spec is not None, "gui module should be discoverable after installation"


def test_tsarakonta_package_importable():
    # ensure the package is importable
    import importlib

    spec = importlib.util.find_spec("finances.tsarakonta")
    assert spec is not None, "finances.tsarakonta should be discoverable as a package"


def test_tsarakonta_help_output(tmp_path, monkeypatch):
    # pandas is required by the financial module; skip the test if not available
    try:
        import pandas  # type: ignore
    except ImportError:
        pytest.skip("pandas not installed, skipping financial tool invocation")

    # run the module with --help and ensure it exits successfully
    cmd = [sys.executable, "-m", "finances.tsarakonta", "--help"]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    assert proc.returncode == 0
    assert "Lance TsaraKonta" in proc.stdout or "usage" in proc.stdout.lower()
