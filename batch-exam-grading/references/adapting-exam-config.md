# 考试配置适配说明

## 一、为什么优先改 `exam_config.json`

同一门课的机考试卷经常会出现这些变化：

- 题号数量变化
- 题目分值变化
- 标准答案变化
- 程序分析题与编程题的评分 rubric 变化
- 大题下的小题数量变化

这些变化大多属于“试卷内容变化”，而不是“答题纸提取逻辑变化”。对这类变化，优先修改 `exam_config.json`，通常就能完成适配，不必先改 Python 脚本。

只有当答题纸版式、文件名格式、代码存放位置、题号格式等提取方式发生明显变化时，才需要继续调整 `parsing` / `extract`，甚至修改脚本。

---

## 二、哪些字段通常可以复用

如果课程类型不变、答题纸结构相近，下列配置通常可以直接复用或只做少量修改。

### 1. `files`

通常只需要改文件名或目录，不需要改字段结构：

- `answer_folder`
- `roster_file`（名单文件字段，默认推荐 `roster.xlsx`；旧字段 `roster_csv` 继续兼容，若两者同时存在则优先使用 `roster_file`）
- `output_xlsx`
- `llm_requests_jsonl`（LLM 请求导出文件路径，默认 `llm_requests.jsonl`，主要用于 agent-runner / Codex 流程）
- `llm_grades_jsonl`（主流程合并主观题评分结果的文件路径，默认 `llm_grades.jsonl`；若改了这个路径，应同步让 OpenAI / Codex runner 的输出路径与之保持一致）
- `llm_cache_jsonl`（LLM 评分缓存文件路径，默认 `llm_cache.jsonl`，仅在本轮评分流程中临时复用）

### 2. `roster`

如果名单表字段长期稳定，这部分通常可以复用：

- `student_id_field`
- `name_field`
- `seq_field`

### 3. `llm`

主观题评分后端通常可以长期复用，但这不表示可以放弃主观题模型评分本身：

- `mode`
- `agent_backend`
- `provider`
- `require_for_subjective`
- `review_policy`

其中：

- `mode` / `provider` / `agent_backend` 表示主观题模型评分的执行后端选择，不表示可以跳过模型评分。
- `require_for_subjective` 虽然属于 `llm` 配置的一部分，但它不应被当作普通调参项。对包含 `analysis` / `code` 的试卷，默认保持强制模型评分；不要因为用户只提到“导出成绩”“新建一个输出目录/文件名”“按当前文件夹评分”而临时关闭。

推荐的跨平台语义是：

- `mode = "llm_api"`：直接走模型 API
- `mode = "agent_runner"`：走本地代理流程
- `agent_backend = "claude" | "codex"`：当 `mode = "agent_runner"` 时决定具体平台

推荐的跨平台主观题配置组合示例：

- Claude：`mode = "agent_runner"` + `agent_backend = "claude"`
- Codex：`mode = "agent_runner"` + `agent_backend = "codex"`
- API：`mode = "llm_api"` + `provider = "openai"`（或其它 provider）

上述三种组合的区别仅在主观题评分的执行后端，不在于是否进行模型评分；它们都要求产出兼容的主观题评分结果后再合并成绩。

推荐长期保留的复核策略：

- 0 分强制人工复核
- 满分强制人工复核
- `confidence=low` 强制人工复核

### 4. 常见提取方式

如果答题纸版式没有大变，这些 `extract.source` 往往仍然适用：

- `choice_inline`
- `regex`
- `section_question`
- `section_regex`
- `code_table`

---

## 三、哪些字段每次换卷基本都要改

这些字段直接绑定某一套具体试卷，换卷后应优先检查。

### 1. `questions`

这是每次换卷时最核心的配置块，通常都要改。至少检查：

- `id`
- `type`
- `score`
- `section`
- `extract`
- `grading`
- `llm.reference_answer`
- `llm.rubric`

建议把 `questions` 理解为“逐小题评分清单”：

- 一道小题对应一条配置
- Excel 会输出 `Q{id}` 逐题分数
- 同一大题的多个小题通过相同 `section` 自动汇总

### 2. 客观题答案

每次考试几乎都要更新：

- 选择题标准答案
- 填空题标准答案
- 数值映射
- 关键词匹配列表

例如：

- `grading.mode = "exact"`
- `grading.answer = "B"`
- `grading.answers = ["JVM", "虚拟机"]`

### 3. 主观题评分 rubric

程序分析题和编程题最容易变化的是这部分：

- 参考答案
- 得分点
- 各得分点分值
- 可接受的等价表达
- 人工复核策略

如果本次考抽象类、下次考接口、多态或异常处理，rubric 应当同步重写，而不是沿用旧卷。

### 4. `section`

`section` 决定大题汇总分组，换卷时必须检查是否仍然正确。

例如：

- `Q1..Q15` 是否仍属于“一、选择题”
- `Q16..Q20` 是否仍属于“二、填空题”
- `Q21..Q22` 是否仍属于“三、程序分析题”
- `Q23..Q25` 是否仍属于“四、编程题”

同一大题下的多个小题必须使用**完全一致**的 `section` 文本，否则系统会把它们当作不同大题分别汇总。

---

## 四、哪些情况说明答题纸版式变了

以下情况通常意味着不能只改答案和 rubric，还要继续检查 `parsing` / `extract` 是否需要调整。

### 1. 区段标题变化

例如：

- “程序分析题”改成“简答题”
- “编程题”改成“程序设计题”

需要检查：

- `parsing.sections.choice.start_keywords`
- `parsing.sections.choice.end_keywords`
- `parsing.sections.analysis.start_keywords`
- `parsing.sections.analysis.end_keywords`
- 其他自定义 section 的关键词

### 2. 学号文件名变化

例如原来是：

- `124232025001_张三_答题纸.docx`

后来变成：

- `张三-124232025001.docx`

需要修改：

- `parsing.student_id_pattern`

### 3. 代码不再放在表格里

如果原先依赖 `code_table`，但学生代码改为正文粘贴或多个分散表格，需要调整：

- `extract.source`
- `table_index`
- 甚至新增更适合的新提取方式

### 4. 题号格式变化

例如：

- `1、` 改成 `第1题`
- `1.（5分）` 改成 `（1）`

需要检查：

- `extract.pattern`
- `question_number`
- `section_regex` / `regex` 的正则表达式

---

## 五、哪些情况才需要改 Python 脚本

原则上，优先尝试通过配置解决。只有在下面这些情况时，才建议修改 Python 脚本。

### 1. 现有提取来源不够用

如果以下方式都无法覆盖新卷：

- `choice_inline`
- `regex`
- `section_question`
- `section_regex`
- `code_table`

才考虑新增提取器。

### 2. 代码识别逻辑失效

例如：

- 代码不在表格里
- 表格结构高度不稳定
- 多道代码题的存放方式和旧卷完全不同

### 3. Excel 输出结构需要增强

例如以后希望增加：

- 每题是否满分
- 每大题得分率
- 大模型证据摘要
- 复核优先级
- 多 sheet 的分层输出

这种属于能力增强，才值得修改引擎。

---

## 六、推荐工作流

建议不要追求“一份永远通用的 `exam_config.json`”，而是采用“模板 + 每次复制”的方式。

### 1. 保留母版

保留一份稳定模板：

- `references/exam_config.template.json`

模板里放：

- 常见 `files`
- 常见 `roster`
- 常见 `llm`
- 常见 `parsing`
- `questions` 的写法范式

### 2. 每次新卷复制一份具体配置

例如：

- `exam_config.oop-final-a.json`
- `exam_config.oop-final-b.json`
- `exam_config.oop-2026-06.json`

不要直接覆盖母版。

### 3. 每次优先按这 4 类检查

按顺序处理通常最省时间：

1. `questions`
2. `section`
3. 客观题答案
4. 主观题 rubric

补充规则：如果用户提出“新建一个 xxx 保存成绩”“按当前文件夹直接评分”“先导出一个成绩表”等需求，默认仅解释为输入来源和输出命名需求，不自动推断为允许更改 `llm` 主观题评分策略，也不构成关闭 `llm.enabled` 或把 `require_for_subjective` 改成 `false` 的依据。

### 4. 再检查提取是否仍成立

重点看：

- 学号能否提取
- 选择题能否提取
- 程序分析题能否切段
- 代码题能否正确提取

### 5. 最后才考虑改脚本

如果前面这些都调不通，再评估是否需要改 Python。

---

## 七、换卷必看检查清单

每次新卷上线前，至少核对以下项目：

- [ ] 学号文件名格式变没变
- [ ] 选择题区段标题变没变
- [ ] 填空题区段标题变没变
- [ ] 程序分析题区段标题变没变
- [ ] 编程题区段标题变没变
- [ ] 题号总数变没变
- [ ] 每题分值变没变
- [ ] 大题归组变没变
- [ ] 客观题标准答案变没变
- [ ] 主观题 rubric 变没变
- [ ] 代码是否还在表格里
- [ ] 输出表是否仍需要 `Q1..Qn + 大题小计 + 总分 + 备注`

---

## 八、与当前 skill 的关系

当前 `batch-exam-grading` 已支持：

- 逐小题评分（`Q{id}`）
- 按 `section` 自动统计大题总分
- 主观题默认强制 LLM；fallback 仅为显式授权的例外路径
- Excel 输出每题分、总分、大题分

因此，换卷时优先把它当成“改配置”的工作，而不是“改脚本”的工作。只有当答案提取方式真的变了，才去调整 Python。
