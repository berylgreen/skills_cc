---
name: batch-exam-grading
description: >-
  批量机考试卷自动评分工具。从 Word 答题纸(docx)提取学生答案，按 `exam_config.json` 中定义的试卷结构、题型、答案与评分标准完成评分，
  客观题走确定性规则，程序分析题与编程题可按配置调用大模型 rubric 评分，并输出 Excel 成绩表与模型审计记录。适用于高校机考批量阅卷、
  换卷后重建评分配置、从答题纸提取答案、主观题模型评分、人工复核追踪。关键词：机考评分、批量阅卷、答题纸提取、自动评分、大模型评分、
  rubric、exam_config、Excel成绩表。
---

# 批量机考试卷评分

## 核心策略

采用“配置优先，主观题默认强制 LLM，完成后自动清理临时文件”：

> 自然语言需求解释规则：当用户说“根据当前文件夹中的文件进行评分”“新建一个 gpt_1 保存成绩”“先导出成绩表”这类话时，默认只表示输入来源与输出命名/目录需求，不表示允许关闭主观题模型评分、启用 fallback，或跳过 `llm_grade.py` 的标准流程。只要试卷中存在 `analysis` / `code` 题，就必须先完成 `prepare -> LLM/agent 评分 -> merge` 的完整流程。

- 试卷结构、名单字段、题型、答案和评分规则都写在 `exam_config.json`。
- `grade_exam.py` 只按配置提取答案、评分并写 Excel。
- `llm_grade.py` 只按配置为主观题生成请求或调用模型评分。
- 主观题评分模式支持两类：
  - `llm_api`：通过 `llm_grade.py` 调外部模型 API（当前内置 `openai` 实现）
  - `agent_runner`：保留同样的请求/结果格式，但评分结果由本地 agent 流程产出后再合并；具体平台由 `llm.agent_backend` 决定，当前支持 `claude`，并为 `codex` 预留配置位
- `analysis` 和 `code` 题默认必须先有模型评分结果；缺失时直接报错，不能仅因为用户想“先出成绩”或“新建一个输出文件名”就关闭 `llm.enabled`、把 `llm.require_for_subjective` 改成 `false`，或直接跳过主观题评分流程。
- `llm_grade.py` 会按 `request_hash` 复用 `files.llm_cache_jsonl` 指向的缓存文件中的历史结果；若未配置，则默认使用 `llm_cache.jsonl`。同一请求不会重复调用模型，但该缓存只在本次评分流程中临时保留。
- `grade_exam.py` 在成功导出最终成绩文件后，会自动删除 `files.llm_requests_jsonl`、`files.llm_grades_jsonl`、`files.llm_cache_jsonl` 指向的评分中间文件；若未配置，则分别使用默认文件名 `llm_requests.jsonl`、`llm_grades.jsonl`、`llm_cache.jsonl`。最终交付默认只保留 Excel，但主观题必须先真实完成模型评分流程；中间文件被清理，不等于可以省略模型评分步骤。若存在主观题评分结果，最终 Excel 还应保留“模型评分审计”sheet 作为审计痕迹。

优先修改配置，不要先改 Python 脚本。只有答题纸版式或提取来源真的变了，才调整脚本。

## 文件结构

```text
工作目录/
├── exam_config.json                    # 从 references/exam_config.template.json 复制后填写
├── roster.xlsx                         # 学生名单（默认示例；也兼容 csv）
├── answers/                            # 学生答卷 docx
├── grade_exam.py
├── llm_grade.py
├── llm_requests.jsonl                # `files.llm_requests_jsonl` 未配置时的默认待评请求文件名，合并成功后自动删除
├── llm_grades.jsonl                  # `files.llm_grades_jsonl` 未配置时的默认评分结果文件名，合并成功后自动删除
├── llm_cache.jsonl                   # `files.llm_cache_jsonl` 未配置时的默认评分缓存文件名，合并成功后自动删除
└── scores.xlsx                       # 最终保留的成绩文件
```

## 工作流程

### 1. 复制脚本和模板

复制：

- `scripts/grade_exam.py`
- `scripts/llm_grade.py`
- `scripts/exam_engine.py`
- `references/exam_config.template.json`，复制后改名为 `exam_config.json`

### 2. 配置试卷

在 `exam_config.json` 中填写：

- `files`：答卷目录、名单文件、输出文件，以及 `llm_requests_jsonl` / `llm_grades_jsonl` / `llm_cache_jsonl` 等中间文件路径
- `roster`：名单字段名映射
- `parsing`：学号正则、各题区段关键词、代码表格识别关键字
- `questions`：每一道题的题号、题型、分值、提取来源、评分规则
- `llm`：模型名、主观题是否强制模型评分、人工复核策略

如果你面对的是“同一门课但试卷一直在变”的场景，建议按下面顺序阅读：

1. `references/exam-config-quick-checklist.md`：先做 5 分钟快速检查，判断这次换卷主要是“改题目”还是“改版式”。
2. `references/adapting-exam-config.md`：需要系统调整配置时，查看复用字段、必改字段、版式变化字段和脚本修改触发条件。
3. `references/exam_config.template.json`：按模板复制出新的 `exam_config.json` 并开始填写。
4. `references/grading_workflow.md`：配置完成后，按完整流程执行提取检查、主观题评分和成绩合并。
5. `references/codex-integration-strategy.md`：如果要把主观题执行器迁移到 Codex，先看这份接入策略说明。

### 3. 生成主观题待评请求

```powershell
python llm_grade.py --config exam_config.json --mode prepare
```

检查 `files.llm_requests_jsonl` 指向的请求文件（若未配置则默认是 `llm_requests.jsonl`）中的 `student_answer` 是否提取正常。每条请求都带有 `request_hash`，主要用于 agent-runner / Codex 流程。

### 4. 调用模型评分

```powershell
python llm_grade.py --config exam_config.json --mode openai --model <当前可用模型>
```

输出到 `--output` 指定的评分结果文件（默认文件名为 `llm_grades.jsonl`），并同步更新 `files.llm_cache_jsonl` 指向的缓存文件（若未配置则默认是 `llm_cache.jsonl`）。

如果 `exam_config.json` 中设置的是 `llm.mode = "agent_runner"`，则 `llm_grade.py` 不直接调用外部 API。此时主观题评分仍然是必需的，只是评分执行端由本地 agent 流程承担。标准必经流程是：

1. 先执行 `python llm_grade.py --config exam_config.json --mode prepare` 生成 `files.llm_requests_jsonl` 指向的请求文件（默认文件名为 `llm_requests.jsonl`）
2. 再由本地 agent 流程逐条或批量消费这些请求；具体平台由 `llm.agent_backend` 决定。
   - `claude`：当前可按 Claude Code / agent 流程处理
   - `codex`：可使用 `scripts/codex_agent_runner.py` 作为执行器骨架，读取 `files.llm_requests_jsonl` 指向的请求文件（默认 `llm_requests.jsonl`），并通过 CLI `--output` 产出与 `files.llm_grades_jsonl` 对应标准结构兼容的结果文件；若未显式指定输出路径，则默认文件名为 `llm_grades.jsonl`。当前已支持断点续跑与统一结果标准化接口
3. 生成与 `files.llm_grades_jsonl` 对应标准结构兼容的结果文件后，再执行 `grade_exam.py` 合并成绩

不要把 `agent_runner` 理解为 fallback 或“本地规则评分”。它只是主观题模型评分的执行后端不同；除非用户明确授权降级，否则不得因为想先产生成绩表或只调整输出文件名，就关闭 `llm.enabled`、修改 `llm.require_for_subjective`，或跳过上述流程。

### 5. 合并成绩表

```powershell
python grade_exam.py --config exam_config.json
```

输出：

- `成绩` sheet：每题分、总分、大题分
- `模型评分审计` sheet：主观题评分理由、扣分点、置信度、复核标记

## 题目配置约定

`questions` 数组中的**每一项就是一个小题**，不是一个大题。导出的 Excel 会为每个题目生成一列 `Q{id}`，例如 `Q1`、`Q21`、`Q25`，用于保存逐小题得分。

每题至少定义：

- `id`
- `type`：`choice`、`fill`、`analysis`、`code`
- `score`
- `section`
- `extract`

### 小题与大题的关系

- `id` 决定逐题分数列，例如 `id=21` 会输出到 `Q21`。
- `section` 是**大题分组键**，不只是展示标题。
- 同一大题下的多个小题，必须使用**完全一致**的 `section` 文本；Excel 会把这些题的分数自动汇总成一个大题总分列。
- `总分` 是所有 `Q{id}` 分数之和。

例如：

- `Q1` 到 `Q15` 都写 `section: "一、选择题"`，会自动汇总成“一、选择题”列。
- `Q21`、`Q22` 都写 `section: "三、程序分析题"`，会自动汇总成“三、程序分析题”列。
- `Q23`、`Q24`、`Q25` 都写 `section: "四、编程题"`，会自动汇总成“四、编程题”列。

如果同一大题下的 `section` 文本有细微差异（例如空格不同、少字、编号不同），系统会把它们当作不同大题分别汇总，所以配置时要保持完全一致。

可用的提取来源：


- `choice_inline`：从客观题区域的 `1、A 2、B` 这类行提取
- `regex`：按正则抓取答案
- `section_question`：从主观题区段按题号提取
- `code_table`：从代码表格提取

可用的确定性评分模式：

- `exact`
- `contains_any`
- `regex_any`
- `numeric_map`
- `keyword_points`

## 人工复核

优先复核：

- `needs_human_review=true`
- `confidence=low`
- 0 分
- 满分
- 提取为空
- 题目结构和配置不一致

## 受限例外：何时才允许 fallback

- 默认禁止通过关闭 `llm.require_for_subjective` 来完成主观题评分。
- 只有用户**明确要求放弃主观题模型评分**，并接受主观题改走 `fallback_grading` 带来的评分质量下降与审计能力下降时，才允许将其改为 `false`。
- 像“导出成绩”“新建一个输出目录/文件名”“按当前文件夹直接评分”“先出一个成绩表”这类表述，**都不构成降级依据**。
- 同理，除非用户明确授权，否则不要把题级 `llm.enabled` 从 `true` 改成 `false` 来绕过主观题评分流程。

## 降级方案

- 新卷面标题不同：先改 `parsing.sections`
- 名单字段不同：改 `roster`
- 题目数量变化：只改 `questions`
