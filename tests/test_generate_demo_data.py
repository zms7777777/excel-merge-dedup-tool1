# -*- coding: utf-8 -*-
"""Unit tests for generate_demo_data.py (runs as a script, writes to ./data)."""
import os
import subprocess
import sys

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def test_generate_creates_three_demo_tables(tmp_path):
    r = subprocess.run(
        [sys.executable, os.path.join(ROOT, "generate_demo_data.py")],
        capture_output=True, text=True, cwd=str(tmp_path),
    )
    assert r.returncode == 0, r.stderr
    data = tmp_path / "data"
    files = sorted(os.path.basename(str(p)) for p in data.glob("*.xlsx"))
    assert files == ["门店A_客户名单.xlsx", "门店B_客户名单.xlsx", "门店C_客户名单.xlsx"]


def test_generated_data_has_expected_shape(tmp_path):
    subprocess.run(
        [sys.executable, os.path.join(ROOT, "generate_demo_data.py")],
        capture_output=True, text=True, cwd=str(tmp_path),
    )
    df = pd.read_excel(tmp_path / "data" / "门店A_客户名单.xlsx", dtype=str)
    assert list(df.columns) == ["姓名", "手机号", "城市"]
    assert len(df) == 5


def test_generated_data_is_idempotent(tmp_path):
    for _ in range(2):
        r = subprocess.run(
            [sys.executable, os.path.join(ROOT, "generate_demo_data.py")],
            capture_output=True, text=True, cwd=str(tmp_path),
        )
        assert r.returncode == 0, r.stderr
    assert len(list((tmp_path / "data").glob("*.xlsx"))) == 3
