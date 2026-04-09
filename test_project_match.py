# -*- coding: utf-8 -*-
"""
项目匹配逻辑离线测试：读取 Excel「项目」列，统计匹配方式与成功率，并输出结果表。
用法：
  python test_project_match.py                    # 弹窗选择 Excel
  python test_project_match.py 你的文件.xlsx      # 指定文件
输出：控制台汇总 + 项目匹配测试结果.xlsx（与原表同目录）
"""

import os
import sys

import pandas as pd

# 与录单脚本同目录，复用其匹配逻辑
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fx_ludan import get_project_keyword_with_meta, setup_logging

try:
    from tkinter import Tk, filedialog
    TK_AVAILABLE = True
except ImportError:
    TK_AVAILABLE = False

PROJECT_COL = "项目"
OUTPUT_SUFFIX = "项目匹配测试结果"


def _preprocess_project(raw):
    """与 fx_ludan 流程一致：取第一个括号前的内容"""
    if not raw or not str(raw).strip():
        return ""
    s = str(raw).strip()
    if ")" in s:
        s = s.split(")")[0].strip()
    return s


def run_test(excel_path):
    if not os.path.isfile(excel_path):
        print(f"文件不存在: {excel_path}")
        return
    try:
        df = pd.read_excel(excel_path, header=0)
    except Exception as e:
        print(f"读取 Excel 失败: {e}")
        return

    if PROJECT_COL not in df.columns:
        print(f"未找到列「{PROJECT_COL}」，当前列: {list(df.columns)}")
        return

    rows = []
    for idx, row in df.iterrows():
        raw = row.get(PROJECT_COL, "")
        text = _preprocess_project(raw)
        keyword, match_type, fuzzy_score = get_project_keyword_with_meta(text)
        rows.append({
            "行号": idx + 1,
            "原文": raw,
            "预处理后": text,
            "匹配结果": keyword,
            "匹配方式": match_type,
            "模糊得分": fuzzy_score if fuzzy_score is not None else "",
        })

    result_df = pd.DataFrame(rows)

    # 统计
    total = len(result_df)
    by_type = result_df["匹配方式"].value_counts()
    n_default = by_type.get("default", 0)
    n_fuzzy = by_type.get("fuzzy_partial", 0) + by_type.get("fuzzy_token", 0)
    n_exact = by_type.get("exact_contains", 0)
    n_synonym = by_type.get("synonym", 0)
    default_rate = (n_default / total * 100) if total else 0
    fuzzy_scores = result_df.loc[result_df["模糊得分"] != "", "模糊得分"]
    low_fuzzy = (fuzzy_scores < 85).sum() if len(fuzzy_scores) else 0

    # 控制台报告
    print()
    print("=" * 50)
    print("项目匹配测试报告")
    print("=" * 50)
    print(f"数据源: {excel_path}")
    print(f"总行数: {total}")
    print()
    print("匹配方式分布:")
    for k, v in by_type.items():
        pct = v / total * 100
        print(f"  {k}: {v} ({pct:.1f}%)")
    print()
    print("关键指标:")
    print(f"  默认兜底: {n_default} 条 ({default_rate:.1f}%)  ← 越低越好")
    print(f"  模糊匹配: {n_fuzzy} 条")
    print(f"  模糊得分<85: {low_fuzzy} 条  ← 可考虑补同义词或标准词")
    print("=" * 50)
    print()
    print("匹配结果明细:")
    print("-" * 50)
    for _, r in result_df.iterrows():
        score_str = f" 得分={r['模糊得分']}" if r["模糊得分"] != "" else ""
        print(f"  {r['行号']:4d} | {str(r['原文'])[:36]:36s} -> {r['匹配结果']:12s} [{r['匹配方式']}]{score_str}")
    print("-" * 50)

    # 保存结果表
    out_dir = os.path.dirname(os.path.abspath(excel_path))
    out_name = f"{OUTPUT_SUFFIX}.xlsx"
    out_path = os.path.join(out_dir, out_name)
    try:
        result_df.to_excel(out_path, index=False)
        print(f"结果已保存: {out_path}")
    except Exception as e:
        print(f"保存结果表失败: {e}")
    print()


def main():
    setup_logging()
    if len(sys.argv) >= 2:
        path = sys.argv[1]
    elif TK_AVAILABLE:
        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="选择包含「项目」列的 Excel",
            filetypes=[("Excel", "*.xlsx"), ("所有文件", "*.*")]
        )
        root.destroy()
        if not path:
            print("未选择文件，退出")
            return
    else:
        print("用法: python test_project_match.py [Excel路径]")
        print("或安装 tkinter 后无参数运行以弹窗选择文件")
        return
    run_test(path)


if __name__ == "__main__":
    main()
