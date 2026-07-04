#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
共享工具模块（shared Excel/CSV utilities）
------------------------------------------------------------------
把 merge_tool.py 与 generate_demo_data.py 中重复的表格读写、清洗逻辑
抽取到这里，供各脚本统一调用，避免复制粘贴。
"""

import glob
import os

import pandas as pd

# 支持的表格后缀
TABLE_PATTERNS = ("*.xlsx", "*.xls", "*.csv")
# Excel 打开时生成的临时文件前缀
TEMP_FILE_PREFIX = "~$"
# 读 CSV 时依次尝试的常见中文编码
CSV_ENCODINGS = ("utf-8-sig", "gbk", "utf-8")
# 合并时标记每行来源文件的列名
SOURCE_COLUMN = "_来源文件"


def find_tables(input_dir):
    """收集目录下所有表格文件路径（已排除临时文件）。"""
    files = []
    for pattern in TABLE_PATTERNS:
        files.extend(glob.glob(os.path.join(input_dir, pattern)))
    files = [f for f in files if not os.path.basename(f).startswith(TEMP_FILE_PREFIX)]
    return sorted(files)


def read_table(path):
    """读单个文件为 DataFrame，全部按字符串读，避免手机号被转成科学计数法。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        for enc in CSV_ENCODINGS:
            try:
                return pd.read_csv(path, dtype=str, encoding=enc, keep_default_na=False)
            except (UnicodeDecodeError, Exception):
                continue
        raise ValueError(f"无法识别编码：{path}")
    return pd.read_excel(path, dtype=str, keep_default_na=False)


def clean_frame(df):
    """基础清洗：列名和单元格去首尾空格，删除完全空白的行。"""
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    df = df[~(df.apply(lambda r: all(v == "" or v == "nan" for v in r), axis=1))]
    return df


def save_frame(df, path, sheet_name="Sheet1", index=False):
    """把单个 DataFrame 写成一张 Excel 表。"""
    df.to_excel(path, sheet_name=sheet_name, index=index)


def save_sheets(path, sheets, engine="openpyxl"):
    """把多个 DataFrame 写进同一个 Excel 文件。

    sheets: 可迭代对象，每项形如 (df, sheet_name, startrow)，
    startrow 可省略（默认 0），用于在同一 sheet 内堆叠多张表。
    """
    with pd.ExcelWriter(path, engine=engine) as writer:
        for item in sheets:
            df, sheet_name = item[0], item[1]
            startrow = item[2] if len(item) > 2 else 0
            df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=startrow)
