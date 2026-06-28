# -*- coding: utf-8 -*-
"""
批量机考试卷自动评分脚本。
功能：按 exam_config.json 提取答卷、评分并写入 Excel。
"""
import argparse
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


def main():
    args = parse_args()
    config = load_exam_config(args.config)
    result = grade_roster(config)
    output_path = write_excel(config, result)
    print(f"已保存: {output_path} ({len(result['roster'])}名学生)")


if __name__ == "__main__":
    main()
