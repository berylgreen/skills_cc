# -*- coding: utf-8 -*-
"""
批量机考试卷自动评分脚本。
功能：按 exam_config.json 提取答卷、评分并写入 Excel。
"""
import argparse
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

try:
    from exam_engine import grade_roster, load_exam_config, write_excel
except ImportError:
    from scripts.exam_engine import grade_roster, load_exam_config, write_excel


def parse_args():
    parser = argparse.ArgumentParser(description="批量机考评分")
    parser.add_argument("--config", default="exam_config.json")
    return parser.parse_args()


def cleanup_intermediate_files(config):
    files_cfg = config.get("files", {})
    output_path = os.path.abspath(files_cfg.get("output_xlsx", "成绩表.xlsx"))
    candidates = [
        files_cfg.get("llm_requests_jsonl", "llm_requests.jsonl"),
        files_cfg.get("llm_grades_jsonl", "llm_grades.jsonl"),
        files_cfg.get("llm_cache_jsonl", "llm_cache.jsonl"),
    ]
    removed = []
    failures = []
    seen = set()
    for path in candidates:
        if not path:
            continue
        normalized_path = os.path.abspath(path)
        if normalized_path == output_path or normalized_path in seen:
            continue
        seen.add(normalized_path)
        if not os.path.exists(path):
            continue
        try:
            os.remove(path)
            removed.append(path)
        except OSError as exc:
            failures.append(f"{path}: {exc}")
    return removed, failures


def main():
    args = parse_args()
    config = load_exam_config(args.config)
    try:
        result = grade_roster(config)
    except ValueError as exc:
        message = str(exc)
        if "缺少主观题模型评分" in message or "主观题模型评分" in message:
            raise SystemExit(
                message
                + "。请先执行 llm_grade.py 的标准流程（prepare -> 模型/agent 评分 -> merge），"
                + "不要通过关闭 llm.enabled 或把 llm.require_for_subjective 改成 false 来绕过主观题评分。"
            )
        raise
    output_path = write_excel(config, result)
    removed, failures = cleanup_intermediate_files(config)
    print(f"已保存: {output_path} ({len(result['roster'])}名学生)")
    if removed:
        print(f"已清理临时文件: {'、'.join(removed)}")
    if failures:
        raise SystemExit("临时文件清理失败: " + "; ".join(failures))


if __name__ == "__main__":
    main()
