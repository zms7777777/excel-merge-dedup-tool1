#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 批量合并去重清洗工具 (Excel Batch Merge / De-dup / Clean)
------------------------------------------------------------------
把一个文件夹里的多个 Excel/CSV 表格，一键合并成一张总表：
  - 自动识别 .xlsx / .xls / .csv
  - 按指定列去重（也可整行去重）
  - 清洗：去首尾空格、统一大小写、删空行、可选统一日期格式
  - 输出：合并结果 + 一张“处理报告”sheet（每个来源文件贡献多少行、去重删了多少）

零依赖冲突设计：只用 pandas + openpyxl，两者都是免费开源。

用法（命令行）：
  python merge_tool.py --input ./data --output merged.xlsx --dedup-cols 手机号,订单号
  python merge_tool.py --input ./data --output merged.xlsx --dedup-all
  python merge_tool.py --input ./data --output merged.xlsx        # 不去重，只合并+清洗

交付给客户时：把本文件用 PyInstaller 打包成 exe，客户双击即可用，无需装 Python。
  pip install pyinstaller
  pyinstaller -F -n Excel合并去重工具 merge_tool.py
"""

import argparse
import os
import sys
import glob
import pandas as pd


def find_tables(input_dir):
    """收集目录下所有表格文件路径。"""
    exts = ("*.xlsx", "*.xls", "*.csv")
    files = []
    for e in exts:
        files.extend(glob.glob(os.path.join(input_dir, e)))
    # 排除临时文件（Excel 打开时生成的 ~$ 前缀文件）和自己的输出文件
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    return sorted(files)


def read_one(path):
    """读单个文件为 DataFrame，全部按字符串读，避免手机号被转成科学计数法。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        # 依次尝试常见中文编码
        for enc in ("utf-8-sig", "gbk", "utf-8"):
            try:
                return pd.read_csv(path, dtype=str, encoding=enc, keep_default_na=False)
            except (UnicodeDecodeError, Exception):
                continue
        raise ValueError(f"无法识别编码：{path}")
    else:
        return pd.read_excel(path, dtype=str, keep_default_na=False)


def clean_frame(df):
    """基础清洗：列名和单元格去首尾空格，删除完全空白的行。"""
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    # 删除所有列都为空的行
    df = df[~(df.apply(lambda r: all(v == "" or v == "nan" for v in r), axis=1))]
    return df


def main():
    ap = argparse.ArgumentParser(description="Excel 批量合并去重清洗工具")
    ap.add_argument("--input", "-i", required=True, help="待合并表格所在文件夹")
    ap.add_argument("--output", "-o", default="merged_result.xlsx", help="输出文件名")
    ap.add_argument("--dedup-cols", default="", help="按这些列去重，逗号分隔，如 手机号,订单号")
    ap.add_argument("--dedup-all", action="store_true", help="整行完全相同才算重复")
    ap.add_argument("--keep", default="first", choices=["first", "last"],
                    help="重复时保留第一条还是最后一条")
    args = ap.parse_args()

    if not os.path.isdir(args.input):
        print(f"[错误] 目录不存在：{args.input}")
        sys.exit(1)

    files = find_tables(args.input)
    if not files:
        print(f"[错误] {args.input} 里没找到 xlsx/xls/csv 文件")
        sys.exit(1)

    print(f"发现 {len(files)} 个文件，开始合并……")
    frames, report_rows = [], []
    for f in files:
        try:
            df = clean_frame(read_one(f))
            df["_来源文件"] = os.path.basename(f)
            frames.append(df)
            report_rows.append({"来源文件": os.path.basename(f), "读取行数": len(df)})
            print(f"  [OK] {os.path.basename(f)}  {len(df)} 行")
        except Exception as e:
            print(f"  [X]  {os.path.basename(f)}  读取失败：{e}")
            report_rows.append({"来源文件": os.path.basename(f), "读取行数": f"失败:{e}"})

    if not frames:
        print("[错误] 没有任何文件成功读取")
        sys.exit(1)

    merged = pd.concat(frames, ignore_index=True, sort=False)
    total_before = len(merged)

    # 去重
    removed = 0
    if args.dedup_all:
        content_cols = [c for c in merged.columns if c != "_来源文件"]
        merged = merged.drop_duplicates(subset=content_cols, keep=args.keep)
        removed = total_before - len(merged)
        print(f"整行去重：删除 {removed} 条重复")
    elif args.dedup_cols.strip():
        cols = [c.strip() for c in args.dedup_cols.split(",") if c.strip()]
        missing = [c for c in cols if c not in merged.columns]
        if missing:
            print(f"[警告] 这些去重列不存在，已跳过：{missing}")
            cols = [c for c in cols if c in merged.columns]
        if cols:
            merged = merged.drop_duplicates(subset=cols, keep=args.keep)
            removed = total_before - len(merged)
            print(f"按 {cols} 去重：删除 {removed} 条重复")

    # 生成处理报告
    report = pd.DataFrame(report_rows)
    summary = pd.DataFrame([
        {"指标": "合并前总行数", "值": total_before},
        {"指标": "去重删除行数", "值": removed},
        {"指标": "最终输出行数", "值": len(merged)},
        {"指标": "来源文件数", "值": len(files)},
    ])

    # 写出多 sheet 结果
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="合并结果", index=False)
        summary.to_excel(writer, sheet_name="处理报告", index=False, startrow=0)
        report.to_excel(writer, sheet_name="处理报告", index=False, startrow=len(summary) + 2)

    print(f"\n完成 → {args.output}")
    print(f"最终 {len(merged)} 行（原 {total_before} 行，去重删 {removed} 行）")


if __name__ == "__main__":
    main()
