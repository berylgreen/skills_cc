# -*- coding: utf-8 -*-
"""
Codex agent 主观题评分执行器骨架（二版）。

职责：
- 读取 llm_requests.jsonl
- 校验请求字段
- 支持按 request_hash 断点续跑
- 预留 Codex 真实评分入口
- 产出与 llm_grades.jsonl 兼容的结果文件

说明：
- 当前版本仍不直接调用真实 Codex 平台能力。
- `run_codex_grading()` 是未来接真实执行器的位置。
- 默认将每条请求写成 needs_human_review=true 的占位结果，便于后续接入真实执行器。
- Codex runner 只负责生成兼容的 `llm_grades.jsonl`，评分结束后的临时文件清理由 `grade_exam.py` 统一处理。
"""

import argparse
import json
import os
from typing import Any, Dict, Iterable, List, Set

REQUIRED_REQUEST_FIELDS = [
    'request_hash',
    'student_id',
    'filename',
    'question_id',
    'question_type',
    'max_score',
    'prompt',
    'reference_answer',
    'rubric',
    'student_answer',
]


def read_jsonl(path: str) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    if not os.path.exists(path):
        raise FileNotFoundError(f'未找到请求文件: {path}')
    with open(path, encoding='utf-8-sig') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            rows.append(json.loads(line))
    return rows


def write_jsonl(path: str, rows: Iterable[Dict[str, Any]]) -> int:
    count = 0
    with open(path, 'w', encoding='utf-8') as f:
        for row in rows:
            f.write(json.dumps(row, ensure_ascii=False) + '\\n')
            count += 1
    return count


def load_completed_request_hashes(path: str) -> Set[str]:
    if not path or not os.path.exists(path):
        return set()
    hashes: Set[str] = set()
    with open(path, encoding='utf-8-sig') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            row = json.loads(line)
            request_hash = row.get('request_hash')
            if request_hash:
                hashes.add(str(request_hash))
    return hashes


def load_existing_rows(path: str) -> List[Dict[str, Any]]:
    if not path or not os.path.exists(path):
        return []
    return read_jsonl(path)


def should_skip_request(row: Dict[str, Any], completed_hashes: Set[str]) -> bool:
    return str(row.get('request_hash', '')) in completed_hashes


def validate_request(row: Dict[str, Any]) -> None:
    missing = [key for key in REQUIRED_REQUEST_FIELDS if key not in row]
    if missing:
        raise ValueError(f'主观题请求缺少字段: {missing}')


def build_codex_task(row: Dict[str, Any]) -> Dict[str, Any]:
    return {
        'request_hash': row['request_hash'],
        'student_id': row['student_id'],
        'filename': row['filename'],
        'question_id': row['question_id'],
        'question_type': row['question_type'],
        'max_score': row['max_score'],
        'grading_prompt': {
            'prompt': row['prompt'],
            'reference_answer': row['reference_answer'],
            'rubric': row['rubric'],
            'student_answer': row['student_answer'],
        }
    }


def run_codex_placeholder(task: Dict[str, Any]) -> Dict[str, Any]:
    return {
        'ok': True,
        'score': 0,
        'confidence': 'low',
        'needs_human_review': True,
        'deductions': ['codex_agent_runner 尚未接入真实 Codex 执行器'],
        'evidence': [],
        'raw_response': None,
        'status': 'placeholder',
    }


def run_codex_grading(task: Dict[str, Any]) -> Dict[str, Any]:
    """未来接入真实 Codex 执行器时，替换这里。"""
    return run_codex_placeholder(task)


def normalize_codex_result(task: Dict[str, Any], raw_result: Dict[str, Any]) -> Dict[str, Any]:
    if not raw_result.get('ok'):
        return {
            'request_hash': task['request_hash'],
            'student_id': task['student_id'],
            'filename': task['filename'],
            'question_id': task['question_id'],
            'question_type': task['question_type'],
            'max_score': task['max_score'],
            'score': 0,
            'confidence': 'low',
            'needs_human_review': True,
            'deductions': [f"Codex 调用失败: {raw_result.get('error', 'unknown_error')}"] ,
            'evidence': [],
            'grader_mode': 'agent_runner',
            'agent_backend': 'codex',
            'status': 'error',
        }
    score = float(raw_result.get('score', 0))
    max_score = float(task['max_score'])
    score = max(0.0, min(score, max_score))
    if score.is_integer():
        score = int(score)
    return {
        'request_hash': task['request_hash'],
        'student_id': task['student_id'],
        'filename': task['filename'],
        'question_id': task['question_id'],
        'question_type': task['question_type'],
        'max_score': task['max_score'],
        'score': score,
        'confidence': raw_result.get('confidence', 'low'),
        'needs_human_review': bool(raw_result.get('needs_human_review', True)),
        'deductions': list(raw_result.get('deductions', [])),
        'evidence': list(raw_result.get('evidence', [])),
        'grader_mode': 'agent_runner',
        'agent_backend': 'codex',
        'status': raw_result.get('status', 'success'),
    }


def parse_args():
    parser = argparse.ArgumentParser(description='Codex agent 主观题评分执行器骨架')
    parser.add_argument('--requests', default='llm_requests.jsonl')
    parser.add_argument('--output', default='llm_grades.jsonl')
    return parser.parse_args()


def main():
    args = parse_args()
    requests = read_jsonl(args.requests)
    existing_rows = load_existing_rows(args.output)
    completed_hashes = load_completed_request_hashes(args.output)
    new_rows: List[Dict[str, Any]] = []
    for row in requests:
        validate_request(row)
        if should_skip_request(row, completed_hashes):
            continue
        task = build_codex_task(row)
        raw_result = run_codex_grading(task)
        new_rows.append(normalize_codex_result(task, raw_result))
    all_rows = existing_rows + new_rows
    count = write_jsonl(args.output, all_rows)
    print(f'已生成 Codex 评分结果: {args.output} ({count}条，其中新增{len(new_rows)}条)')


if __name__ == '__main__':
    main()
