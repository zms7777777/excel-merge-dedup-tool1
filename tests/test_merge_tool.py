# -*- coding: utf-8 -*-
"""Unit tests for merge_tool.py."""
import os
import sys

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import merge_tool  # noqa: E402


# --------------------------------------------------------------------------- #
# find_tables
# --------------------------------------------------------------------------- #
def test_find_tables_collects_supported_extensions(tmp_path):
    for name in ("a.xlsx", "b.xls", "c.csv", "d.txt", "e.json"):
        (tmp_path / name).write_text("x")
    found = merge_tool.find_tables(str(tmp_path))
    basenames = [os.path.basename(f) for f in found]
    assert basenames == ["a.xlsx", "b.xls", "c.csv"]


def test_find_tables_excludes_excel_temp_files(tmp_path):
    (tmp_path / "real.xlsx").write_text("x")
    (tmp_path / "~$real.xlsx").write_text("x")
    found = merge_tool.find_tables(str(tmp_path))
    basenames = [os.path.basename(f) for f in found]
    assert basenames == ["real.xlsx"]


def test_find_tables_returns_sorted(tmp_path):
    for name in ("z.csv", "a.csv", "m.csv"):
        (tmp_path / name).write_text("x")
    found = merge_tool.find_tables(str(tmp_path))
    basenames = [os.path.basename(f) for f in found]
    assert basenames == sorted(basenames)


def test_find_tables_empty_dir(tmp_path):
    assert merge_tool.find_tables(str(tmp_path)) == []


# --------------------------------------------------------------------------- #
# read_one
# --------------------------------------------------------------------------- #
def test_read_one_xlsx_keeps_strings(tmp_path):
    path = tmp_path / "t.xlsx"
    pd.DataFrame({"手机号": ["13800000001"], "名": ["张伟"]}).to_excel(path, index=False)
    df = merge_tool.read_one(str(path))
    # numbers must stay as strings (no scientific notation)
    assert df["手机号"].iloc[0] == "13800000001"
    assert df["手机号"].dtype == object


def test_read_one_csv_utf8(tmp_path):
    path = tmp_path / "t.csv"
    path.write_text("姓名,手机号\n张伟,13800000001\n", encoding="utf-8-sig")
    df = merge_tool.read_one(str(path))
    assert list(df.columns) == ["姓名", "手机号"]
    assert df["手机号"].iloc[0] == "13800000001"


def test_read_one_csv_gbk(tmp_path):
    path = tmp_path / "t_gbk.csv"
    path.write_bytes("姓名,城市\n李娜,上海\n".encode("gbk"))
    df = merge_tool.read_one(str(path))
    assert df["城市"].iloc[0] == "上海"


def test_read_one_csv_keeps_empty_as_string(tmp_path):
    path = tmp_path / "t.csv"
    path.write_text("a,b\n,x\n", encoding="utf-8")
    df = merge_tool.read_one(str(path))
    # keep_default_na=False -> empty stays "" not NaN
    assert df["a"].iloc[0] == ""


# --------------------------------------------------------------------------- #
# clean_frame
# --------------------------------------------------------------------------- #
def test_clean_frame_strips_cells_and_columns():
    df = pd.DataFrame({" 姓名 ": [" 张伟 ", "李娜"], "城市": ["北京 ", " 上海"]})
    out = merge_tool.clean_frame(df)
    assert list(out.columns) == ["姓名", "城市"]
    assert out["姓名"].tolist() == ["张伟", "李娜"]
    assert out["城市"].tolist() == ["北京", "上海"]


def test_clean_frame_drops_fully_empty_rows():
    df = pd.DataFrame({"a": ["x", "", "  "], "b": ["y", "", ""]})
    out = merge_tool.clean_frame(df)
    assert len(out) == 1
    assert out["a"].tolist() == ["x"]


def test_clean_frame_drops_nan_rows():
    df = pd.DataFrame({"a": ["x", np.nan], "b": ["y", np.nan]})
    out = merge_tool.clean_frame(df)
    assert len(out) == 1


def test_clean_frame_keeps_partially_filled_rows():
    df = pd.DataFrame({"a": ["x", ""], "b": ["", "y"]})
    out = merge_tool.clean_frame(df)
    assert len(out) == 2


# --------------------------------------------------------------------------- #
# main (in-process, driven via sys.argv so coverage tracks main())
# --------------------------------------------------------------------------- #
def _make_demo(dirpath):
    # A: 4 rows incl. one fully-empty row that clean_frame should drop.
    pd.DataFrame({
        "姓名": ["张伟 ", "李娜", "王芳", ""],
        "手机号": ["13800000001", "13800000002", "13800000003", ""],
    }).to_excel(os.path.join(dirpath, "A.xlsx"), index=False)
    pd.DataFrame({
        "姓名": ["王芳", "赵强"],  # 王芳/13800000003 duplicates A
        "手机号": ["13800000003", "13800000005"],
    }).to_excel(os.path.join(dirpath, "B.xlsx"), index=False)


def _run_main(monkeypatch, args):
    monkeypatch.setattr(sys, "argv", ["merge_tool.py", *args])
    merge_tool.main()


@pytest.fixture
def demo_dir(tmp_path):
    data = tmp_path / "data"
    data.mkdir()
    _make_demo(str(data))
    return data


def test_main_dedup_by_column(monkeypatch, tmp_path, demo_dir):
    out = tmp_path / "out.xlsx"
    _run_main(monkeypatch, ["-i", str(demo_dir), "-o", str(out), "--dedup-cols", "手机号"])
    merged = pd.read_excel(out, sheet_name="合并结果", dtype=str)
    # A cleaned to 3 rows + B 2 rows = 5, minus 1 phone dup -> 4
    assert len(merged) == 4
    assert "_来源文件" in merged.columns


def test_main_dedup_all(monkeypatch, tmp_path, demo_dir):
    out = tmp_path / "out.xlsx"
    _run_main(monkeypatch, ["-i", str(demo_dir), "-o", str(out), "--dedup-all"])
    merged = pd.read_excel(out, sheet_name="合并结果", dtype=str)
    assert len(merged) == 4


def test_main_no_dedup_keeps_all(monkeypatch, tmp_path, demo_dir):
    out = tmp_path / "out.xlsx"
    _run_main(monkeypatch, ["-i", str(demo_dir), "-o", str(out)])
    merged = pd.read_excel(out, sheet_name="合并结果", dtype=str)
    assert len(merged) == 5  # only the empty row removed, no dedup


def test_main_missing_dedup_column_warns(monkeypatch, tmp_path, demo_dir, capsys):
    out = tmp_path / "out.xlsx"
    _run_main(monkeypatch, ["-i", str(demo_dir), "-o", str(out), "--dedup-cols", "不存在的列"])
    assert "警告" in capsys.readouterr().out


def test_main_report_sheet_written(monkeypatch, tmp_path, demo_dir):
    out = tmp_path / "out.xlsx"
    _run_main(monkeypatch, ["-i", str(demo_dir), "-o", str(out), "--dedup-cols", "手机号"])
    report = pd.read_excel(out, sheet_name="处理报告", dtype=str)
    assert report.iloc[0, 0] == "合并前总行数"


def test_main_keep_last(monkeypatch, tmp_path, demo_dir):
    out = tmp_path / "out.xlsx"
    _run_main(monkeypatch, ["-i", str(demo_dir), "-o", str(out),
                            "--dedup-cols", "手机号", "--keep", "last"])
    merged = pd.read_excel(out, sheet_name="合并结果", dtype=str)
    # 王芳 kept from B (last occurrence)
    wangfang = merged[merged["姓名"] == "王芳"]
    assert wangfang["_来源文件"].iloc[0] == "B.xlsx"


def test_main_reads_csv_files(monkeypatch, tmp_path):
    data = tmp_path / "data"
    data.mkdir()
    (data / "a.csv").write_text("姓名,手机号\n张伟,13800000001\n", encoding="utf-8-sig")
    out = tmp_path / "out.xlsx"
    _run_main(monkeypatch, ["-i", str(data), "-o", str(out)])
    merged = pd.read_excel(out, sheet_name="合并结果", dtype=str)
    assert merged["手机号"].iloc[0] == "13800000001"


def test_main_continues_when_one_file_unreadable(monkeypatch, tmp_path, capsys):
    data = tmp_path / "data"
    data.mkdir()
    pd.DataFrame({"姓名": ["张伟"], "手机号": ["13800000001"]}).to_excel(
        data / "good.xlsx", index=False)
    (data / "bad.xlsx").write_text("not a real xlsx")
    out = tmp_path / "out.xlsx"
    _run_main(monkeypatch, ["-i", str(data), "-o", str(out)])
    assert "读取失败" in capsys.readouterr().out
    merged = pd.read_excel(out, sheet_name="合并结果", dtype=str)
    assert len(merged) == 1


def test_main_errors_on_missing_dir(monkeypatch, tmp_path, capsys):
    with pytest.raises(SystemExit) as exc:
        _run_main(monkeypatch, ["-i", str(tmp_path / "nope"), "-o", str(tmp_path / "o.xlsx")])
    assert exc.value.code == 1
    assert "目录不存在" in capsys.readouterr().out


def test_main_errors_on_empty_dir(monkeypatch, tmp_path, capsys):
    data = tmp_path / "empty"
    data.mkdir()
    with pytest.raises(SystemExit) as exc:
        _run_main(monkeypatch, ["-i", str(data), "-o", str(tmp_path / "o.xlsx")])
    assert exc.value.code == 1
    assert "没找到" in capsys.readouterr().out


def test_main_errors_when_all_files_fail(monkeypatch, tmp_path, capsys):
    data = tmp_path / "data"
    data.mkdir()
    (data / "bad.xlsx").write_text("not a real xlsx")
    with pytest.raises(SystemExit) as exc:
        _run_main(monkeypatch, ["-i", str(data), "-o", str(tmp_path / "o.xlsx")])
    assert exc.value.code == 1
    assert "没有任何文件成功读取" in capsys.readouterr().out


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v"]))
