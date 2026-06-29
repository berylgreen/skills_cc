# Codex 接入策略说明

## 一、目标

本文档说明如何在不破坏当前评分体系的前提下，把 `batch-exam-grading` 的主观题评分能力迁移或扩展到 Codex。

目标不是重写整套评分系统，而是复用现有结构：

- `exam_config.json`
- `llm_requests.jsonl`
- `llm_grades.jsonl`
- `exam_engine.py`
- `grade_exam.py`
- Excel 成绩输出链路

换句话说，迁移到 Codex 时，优先替换的是“主观题执行器层”，而不是“评分引擎层”。

---

## 二、为什么优先选择 CLI，而不是直接 SDK

对于当前这套 skill，Codex 接入推荐优先走 CLI 子进程，而不是一开始就直接接 SDK。

### 推荐原因

1. **对现有系统侵入最小**
   - 只需要继续完善 `scripts/codex_agent_runner.py`
   - 不需要改动客观题评分、逐题汇总、大题小计和 Excel 输出

2. **与现有 `agent_runner` 语义一致**
   - 当前主观题执行模式已经抽象为：
     - `llm.mode = "llm_api"`
     - `llm.mode = "agent_runner"`
   - Codex 更适合作为 `agent_runner` 的一个具体执行后端，而不是重新发明一套 API 模式

3. **便于保留文件协议**
   - CLI 方案天然适合消费 `llm_requests.jsonl`
   - 也天然适合产出 `llm_grades.jsonl`

4. **后续可以再升级**
   - 如果未来 Codex SDK 更稳定、并发要求更高、日志要求更完整，再从 CLI 升级到 SDK

### 不优先选择 SDK 的原因

- 会让主观题执行器重新向 `llm_api` 方向耦合
- 会增加环境依赖和迁移成本
- 对当前阶段的收益不如 CLI 明显

---

## 三、当前哪些层可以直接复用

迁移到 Codex 时，这些层原则上不需要重写。

### 1. 配置层

- `exam_config.json`
- `questions`
- `section`
- `grading`
- `llm.reference_answer`
- `llm.rubric`

### 2. 评分引擎层

- `scripts/exam_engine.py`
- `scripts/grade_exam.py`

它们负责：

- 客观题评分
- 主观题结果读取
- `Q{id}` 逐题分数
- section 汇总
- 总分统计
- Excel 输出

这些逻辑不应感知当前主观题到底来自 Claude、Codex 还是外部 API。

### 3. 中间协议层

- `llm_requests.jsonl`
- `llm_grades.jsonl`

这是最值得保持稳定的一层，也是跨平台迁移时最重要的资产。

---

## 四、Codex 执行器推荐职责

当前推荐由：

- `scripts/codex_agent_runner.py`

来承担 Codex 的主观题执行逻辑。

建议它长期保留以下职责边界：

1. 读取 `llm_requests.jsonl`
2. 校验请求字段
3. 支持按 `request_hash` 断点续跑
4. 调用 Codex 执行评分
5. 把结果标准化为兼容的 `llm_grades.jsonl`

它**不应**负责：

- 客观题评分
- 大题汇总
- 最终 Excel 合并

这些仍应交给 `exam_engine.py` / `grade_exam.py`。

---

## 五、建议保留的核心函数

建议 `scripts/codex_agent_runner.py` 长期保留下列函数接口，即使内部实现后续变化，也尽量不改职责边界。

### 1. `build_codex_task(row)`

作用：
- 把通用请求行转换成 Codex 更容易消费的任务结构

### 2. `run_codex_grading(task)`

作用：
- 真实调用 Codex CLI 或未来 SDK 的执行入口

### 3. `normalize_codex_result(task, raw_result)`

作用：
- 把 Codex 的原始返回翻译成统一的评分结果格式
- 这是最关键的稳定层

### 4. `load_completed_request_hashes(path)`

作用：
- 读取已有结果
- 支持断点续跑和跳过已完成请求

### 5. `should_skip_request(row, completed_hashes)`

作用：
- 避免重复评分同一条请求

---

## 六、推荐的请求格式

建议继续使用统一的 `llm_requests.jsonl`，每行至少包含：

```json
{
  "request_hash": "abc123",
  "student_id": "2025001",
  "filename": "2025001_sample.docx",
  "question_id": 21,
  "question_type": "analysis",
  "max_score": 5,
  "prompt": "第21题题干",
  "reference_answer": "标准答案",
  "rubric": [
    {"point": "结论正确", "score": 5}
  ],
  "student_answer": "学生答案"
}
```

这套请求结构应继续保持平台无关，不要因为迁移到 Codex 就改字段名。

---

## 七、推荐的结果格式

建议 `llm_grades.jsonl` 继续保持平台无关的标准结构，每行至少包含：

```json
{
  "request_hash": "abc123",
  "student_id": "2025001",
  "filename": "2025001_sample.docx",
  "question_id": 21,
  "question_type": "analysis",
  "max_score": 5,
  "score": 4,
  "confidence": "medium",
  "needs_human_review": false,
  "deductions": ["解释不完整"],
  "evidence": ["结论正确，但缺少关键原因说明"],
  "grader_mode": "agent_runner",
  "agent_backend": "codex",
  "status": "success"
}
```

### 建议约束

- `score` 始终为数值
- `deductions` / `evidence` 始终为数组
- `needs_human_review` 始终为布尔值
- `status` 建议只用：
  - `success`
  - `error`
  - `placeholder`

---

## 八、错误结果建议格式

如果某条请求评分失败，不要跳过，建议仍然写一条错误记录：

```json
{
  "request_hash": "abc123",
  "student_id": "2025001",
  "filename": "2025001_sample.docx",
  "question_id": 21,
  "question_type": "analysis",
  "max_score": 5,
  "score": 0,
  "confidence": "low",
  "needs_human_review": true,
  "deductions": ["Codex 调用失败: rate_limit"],
  "evidence": [],
  "grader_mode": "agent_runner",
  "agent_backend": "codex",
  "status": "error"
}
```

这样更利于：

- 保持整批任务可汇总
- 后续人工复核
- 追踪失败原因

---

## 九、建议的接入顺序

推荐分三步推进。

### 第一步：骨架阶段

当前已完成的方向：

- 读取请求
- 写出兼容结果
- 保留标准化接口
- 支持断点续跑

### 第二步：CLI 接入阶段

把：

- `run_codex_grading(task)`

替换成真实的 Codex CLI 调用逻辑。

### 第三步：SDK 升级阶段（可选）

仅当出现以下情况时，再考虑从 CLI 升级到 SDK：

- 批量评分规模很大
- 需要更强的并发控制
- 需要更强的日志和监控
- Codex SDK 已经稳定

---

## 十、接入检查清单

在把 Codex 真正接入 runner 前，建议先逐项确认：

- [ ] `exam_config.json` 已使用 `llm.mode = "agent_runner"`
- [ ] `llm.agent_backend = "codex"`
- [ ] `llm_requests.jsonl` 已能稳定生成
- [ ] 每条请求都带 `request_hash`
- [ ] `run_codex_grading(task)` 能拿到结构化结果或明确错误
- [ ] `normalize_codex_result(...)` 产出的字段与 `llm_grades.jsonl` 兼容
- [ ] 错误记录也能正常落盘
- [ ] 再次运行时，已完成请求会被正确跳过
- [ ] `grade_exam.py` 能继续正常合并成绩

---

## 十一、一句话策略

如果以后要把这套评分 skill 迁到 Codex，最好的方式不是重写评分系统，而是：

**保留配置层、评分引擎层和 JSONL 协议层不动，只替换主观题执行器层。**
