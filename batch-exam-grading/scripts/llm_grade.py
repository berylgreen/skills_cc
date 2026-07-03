# -*- coding: utf-8 -*-
"""
程序分析题/编程题大模型评分脚本。
"""
import argparse
import hashlib
import json
import os
import sys
from typing import Any, Dict, List

sys.stdout.reconfigure(encoding="utf-8")

try:
    from exam_engine import build_llm_requests, load_exam_config
except ImportError:
    from scripts.exam_engine import build_llm_requests, load_exam_config


DEFAULT_OUTPUT = "llm_grades.jsonl"
DEFAULT_REQUESTS = "llm_requests.jsonl"

GRADING_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "score": {"type": "number"},
        "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
        "needs_human_review": {"type": "boolean"},
        "deductions": {"type": "array", "items": {"type": "string"}},
        "evidence": {"type": "array", "items": {"type": "string"}}
    },
    "required": ["score", "confidence", "needs_human_review", "deductions", "evidence"]
}


def anonymous_id(student_id: str) -> str:
    return hashlib.sha256(student_id.encode("utf-8")).hexdigest()[:12]


def build_prompt(request: Dict[str, Any]) -> str:
    payload = {
        "anonymous_student_id": anonymous_id(request["student_id"]),
        "question_id": request["question_id"],
        "question_type": request["question_type"],
        "max_score": request["max_score"],
        "question_prompt": request["prompt"],
        "reference_answer": request["reference_answer"],
        "rubric": request["rubric"],
        "student_answer": request["student_answer"]
    }
    return (
        "你是高校机考阅卷员。只能依据给定评分标准评分，不得自行增加或改变得分点。"
        "允许等价表达和等价代码结构；若答案缺失、无法判断或疑似提取错误，标记 needs_human_review=true。"
        "总分必须在 0 到 max_score 之间。请返回符合 JSON Schema 的结果。\n\n"
        + json.dumps(payload, ensure_ascii=False, indent=2)
    )


def write_jsonl(path: str, rows: List[Dict[str, Any]]) -> int:
    with open(path, "w", encoding="utf-8") as f:
        for row in rows:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")
    return len(rows)


def load_cached_results(*paths: str) -> Dict[str, Dict[str, Any]]:
    cached: Dict[str, Dict[str, Any]] = {}
    for path in paths:
        if not path or not os.path.exists(path):
            continue
        with open(path, encoding="utf-8-sig") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                item = json.loads(line)
                request_hash = item.get("request_hash")
                if request_hash:
                    cached[request_hash] = item
    return cached


def extract_output_json(response: Any) -> Dict[str, Any]:
    text = getattr(response, "output_text", None)
    if text:
        return json.loads(text)
    for item in getattr(response, "output", []) or []:
        for content in getattr(item, "content", []) or []:
            value = getattr(content, "text", None)
            if value:
                return json.loads(value)
    raise ValueError("无法从模型响应中提取 JSON 文本")


def call_openai(prompt: str, model: str) -> Dict[str, Any]:
    try:
        from openai import OpenAI
    except ImportError as exc:
        raise RuntimeError("缺少 openai 包，请先执行 pip install openai") from exc

    client = OpenAI()
    response = client.responses.create(
        model=model,
        input=[
            {"role": "developer", "content": "你是严格、稳定、可审计的考试评分器。必须输出结构化 JSON。"},
            {"role": "user", "content": prompt}
        ],
        temperature=0,
        text={
            "format": {
                "type": "json_schema",
                "name": "exam_grade",
                "schema": GRADING_SCHEMA,
                "strict": True
            }
        }
    )
    return extract_output_json(response)


def apply_review_policy(result: Dict[str, Any], review_policy: Dict[str, Any], max_score: float) -> Dict[str, Any]:
    score = max(0.0, min(float(result.get("score", 0)), max_score))
    result["score"] = int(score) if score.is_integer() else score
    if score in review_policy.get("force_review_scores", []):
        result["needs_human_review"] = True
    if result.get("confidence", "low") in review_policy.get("force_review_confidence", ["low"]):
        result["needs_human_review"] = True
    if review_policy.get("force_review_full_score", True) and score == max_score:
        result["needs_human_review"] = True
    return result


def normalize_llm_config(config: Dict[str, Any]) -> Dict[str, Any]:
    llm_cfg = config.setdefault("llm", {})
    mode = llm_cfg.get("mode", "llm_api")
    if mode == "claude_agent":
        llm_cfg["mode"] = "agent_runner"
        llm_cfg.setdefault("agent_backend", "claude")
    elif llm_cfg.get("mode") == "agent_runner":
        llm_cfg.setdefault("agent_backend", "claude")
    return config


def llm_mode(config: Dict[str, Any]) -> str:
    return config.get("llm", {}).get("mode", "llm_api")


def llm_agent_backend(config: Dict[str, Any]) -> str:
    return config.get("llm", {}).get("agent_backend", "")


def resolve_requests_output(config: Dict[str, Any], cli_value: str) -> str:
    if cli_value != DEFAULT_REQUESTS:
        return cli_value
    return config.get("files", {}).get("llm_requests_jsonl", DEFAULT_REQUESTS)


def call_claude_agent_placeholder(request: Dict[str, Any]) -> Dict[str, Any]:
    raise RuntimeError(
        "当前配置为 claude_agent 模式。请先根据 llm_requests.jsonl 使用 Claude Code / agent 完成评分，"
        "并生成与 llm_grades.jsonl 兼容的结果文件后再执行合并。"
    )


def grade_with_openai(
    requests: List[Dict[str, Any]],
    model: str,
    review_policy: Dict[str, Any],
    cached_results: Dict[str, Dict[str, Any]] | None = None,
) -> List[Dict[str, Any]]:
    rows = []
    cached_results = cached_results or {}
    for idx, request in enumerate(requests, 1):
        cached = cached_results.get(request["request_hash"])
        if cached:
            rows.append(cached)
            continue
        print(f"评分 {idx}/{len(requests)}: {request['student_id']} Q{request['question_id']}")
        try:
            result = call_openai(build_prompt(request), model)
            result = apply_review_policy(result, review_policy, float(request["max_score"]))
        except Exception as exc:
            result = {
                "score": 0,
                "confidence": "low",
                "needs_human_review": True,
                "deductions": [f"模型评分失败: {exc}"],
                "evidence": []
            }
        rows.append({
            "student_id": request["student_id"],
            "anonymous_student_id": anonymous_id(request["student_id"]),
            "filename": request["filename"],
            "question_id": request["question_id"],
            "question_type": request["question_type"],
            "max_score": request["max_score"],
            "request_hash": request["request_hash"],
            **result
        })
    return rows


def parse_args():
    parser = argparse.ArgumentParser(description="程序分析题/编程题大模型评分")
    parser.add_argument("--config", default="exam_config.json")
    parser.add_argument("--mode", choices=["prepare", "openai", "agent_runner", "claude_agent"], default="prepare")
    parser.add_argument("--requests-output", default=DEFAULT_REQUESTS)
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    parser.add_argument("--model", default=os.environ.get("OPENAI_MODEL", ""))
    return parser.parse_args()


def main():
    args = parse_args()
    config = normalize_llm_config(load_exam_config(args.config))
    configured_mode = llm_mode(config)
    agent_backend = llm_agent_backend(config)
    cli_mode = "agent_runner" if args.mode == "claude_agent" else args.mode
    model = args.model or config.get("llm", {}).get("model", "")
    if cli_mode == "openai" and not model:
        raise SystemExit("openai 模式需要 --model 或 exam_config.llm.model 或 OPENAI_MODEL")
    if cli_mode == "openai" and configured_mode == "agent_runner":
        raise SystemExit("当前 exam_config.llm.mode=agent_runner，请改用 --mode agent_runner 或切回 llm_api")
    if cli_mode == "agent_runner" and configured_mode != "agent_runner":
        raise SystemExit("当前 exam_config.llm.mode 不是 agent_runner，请先修改配置或使用 --mode openai")
    if model:
        config.setdefault("llm", {})["model"] = model
    requests_output = resolve_requests_output(config, args.requests_output)
    requests = build_llm_requests(config)
    if cli_mode == "prepare":
        count = write_jsonl(requests_output, requests)
        print(f"已生成待评请求: {requests_output} ({count}条)")
        return
    if cli_mode == "agent_runner":
        count = write_jsonl(requests_output, requests)
        if agent_backend == "codex":
            platform_hint = "Codex agent"
        else:
            platform_hint = "Claude Code / agent"
        message = (
            f"已生成 Agent 待评请求: {requests_output} ({count}条)。"
            f"当前 agent_backend={agent_backend or 'claude'}，请使用 {platform_hint} 对这些请求评分，并产出与 llm_grades.jsonl 兼容的结果文件。"
        )
        print(message)
        if not os.environ.get("BATCH_EXAM_GRADING_TEST"):
            raise SystemExit(message)
        return
    review_policy = config.get("llm", {}).get("review_policy", {})
    cache_path = config.get("files", {}).get("llm_cache_jsonl", "llm_cache.jsonl")
    cached_results = load_cached_results(cache_path, args.output)
    rows = grade_with_openai(requests, model, review_policy, cached_results=cached_results)
    count = write_jsonl(args.output, rows)
    merged_cache = {item["request_hash"]: item for item in cached_results.values()}
    for row in rows:
        merged_cache[row["request_hash"]] = row
    write_jsonl(cache_path, list(merged_cache.values()))
    print(f"已生成模型评分: {args.output} ({count}条)")


if __name__ == "__main__":
    main()
