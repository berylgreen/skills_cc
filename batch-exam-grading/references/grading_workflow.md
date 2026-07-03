# 批量机考评分流程参考

## 一、准备文件

| 文件 | 说明 |
|------|------|
| `exam_config.json` | 试卷结构、评分规则和提取规则 |
| `roster.xlsx` | 学生名单（默认示例；也兼容 csv） |
| `answers/*.docx` | 学生答卷 |
| `grade_exam.py` | 成绩合并与 Excel 导出 |
| `llm_grade.py` | 主观题请求生成与模型评分 |
| `exam_engine.py` | 共享核心逻辑 |
| `llm_requests.jsonl` | `files.llm_requests_jsonl` 未配置时的默认主观题请求导出文件名，主要用于 agent-runner / Codex 流程，最终合并成功后自动删除 |
| `llm_cache.jsonl` | `files.llm_cache_jsonl` 未配置时的默认主观题评分缓存文件名，仅在本次评分流程中临时复用；最终合并成功后自动删除 |

## 二、先配 `exam_config.json`

> 自然语言需求解释规则：当用户说“根据当前文件夹中的文件进行评分”“新建一个 gpt_1 保存成绩”“先导出成绩表”这类话时，默认只表示输入来源与输出命名/目录需求，不表示允许关闭主观题模型评分、启用 fallback，或跳过 `llm_grade.py` 的标准流程。只要试卷中存在 `analysis` / `code` 题，就必须先完成 `prepare -> 模型/agent 评分 -> merge` 的完整流程。

重点字段：

- `files`：目录与文件名
- `roster`：名单列名映射
- `parsing.student_id_pattern`：从文件名提取学号
- `parsing.sections`：选择题、分析题等区段标题
- `questions`：每题的提取来源和评分方式

## 三、先做提取检查

```powershell
python llm_grade.py --config exam_config.json --mode prepare
```

检查 `files.llm_requests_jsonl` 指向的请求文件（若未配置则默认是 `llm_requests.jsonl`）：

- 题号是否完整
- `student_answer` 是否串题
- 主观题是否只来自配置里启用 `llm.enabled=true` 的题
- 同一条请求是否带 `request_hash`

请求文件对应的配置项是 `files.llm_requests_jsonl`。若未显式配置，则默认输出到 `llm_requests.jsonl`。它只用于本轮评分准备，主要服务于 agent-runner / Codex 流程；最终成绩合并成功后会被自动删除。

## 四、评分

### 4.1 主观题模型评分

无论使用 `llm_api` 还是 `agent_runner`，只要存在主观题，都必须先完成本节对应的评分产物生成，不能直接跳到合并成绩。若 `llm.mode = agent_runner`，则此处改为执行本地 agent 评分流程，但仍以生成与 `files.llm_grades_jsonl` 兼容的结果文件为完成标志。

```powershell
python llm_grade.py --config exam_config.json --mode openai --model <model>
```

相同 `request_hash` 会直接复用本轮流程中的缓存，不再重新调用模型。OpenAI 模式会把评分结果写入 CLI `--output` 指定的路径（默认文件名为 `llm_grades.jsonl`），并读取 `files.llm_cache_jsonl` 指向的缓存文件（若未配置则默认是 `llm_cache.jsonl`）。最终成绩合并成功后，`files.llm_grades_jsonl` 与 `files.llm_cache_jsonl` 指向的文件会被自动删除；若未配置，则默认文件名分别为 `llm_grades.jsonl`、`llm_cache.jsonl`。注意：`openai` 模式不会回读 `files.llm_requests_jsonl` 指向的请求导出文件作为评分输入，而是按当前配置重新构建请求。

### 4.2 合并成绩

```powershell
python grade_exam.py --config exam_config.json
```

该步骤在成功导出最终 Excel 后，会自动清理 `files.llm_requests_jsonl`、`files.llm_grades_jsonl`、`files.llm_cache_jsonl` 指向的中间文件；若未配置，则默认分别为 `llm_requests.jsonl`、`llm_grades.jsonl`、`llm_cache.jsonl`。无论这些结果文件来自 OpenAI 还是 Codex runner，评分结束后工作目录默认只保留最终成绩文件；但这只表示中间文件会被清理，不表示可以省略模型评分步骤。只要存在主观题，最终 Excel 仍应以完整的主观题评分链路为前提；若有模型评分结果，还应保留“模型评分审计”sheet 作为审计痕迹。

## 五、换卷时优先改什么

1. 改 `questions`
2. 改 `parsing.sections`
3. 改 `roster`
4. 改 `llm.model` 时会影响请求哈希和缓存命中
5. 只有提取来源完全变了，才改 Python

## 六、常见问题

| 问题 | 原因 | 处理 |
|------|------|------|
| 主观题请求为空 | 文件名学号正则不匹配 | 改 `parsing.student_id_pattern` |
| 某题一直 0 分 | `extract` 或 `grading` 配错 | 检查该题配置 |
| 主观题无法合并 | 缺少 `files.llm_grades_jsonl` 指向的评分结果文件（默认文件名为 `llm_grades.jsonl`）或 request_hash 不匹配 | 先重新跑 `llm_grade.py --mode openai` 生成本轮结果，并确保输出路径与 `files.llm_grades_jsonl` 保持一致，再执行合并。除非用户明确要求放弃主观题模型评分，否则不要通过关闭 `llm.require_for_subjective` 或改走 `fallback_grading` 来绕过此问题 |
| 名单成绩对不上 | 名单字段映射错误 | 改 `roster.student_id_field` 等字段 |
