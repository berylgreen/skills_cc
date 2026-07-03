# 5 分钟换卷速查表

## 1. 每次换卷先改什么

优先按这个顺序检查：

1. `questions`
2. `section`
3. 客观题答案
4. 主观题 rubric

如果只是题目内容变化，通常改完这四类就够了，不要先改脚本。

---

## 2. `questions` 里重点看什么

每次至少核对：

- `id`
- `type`
- `score`
- `section`
- `extract`
- `grading`
- `llm.reference_answer`
- `llm.rubric`

记住：

- 一条 `question` = 一道小题
- Excel 会输出 `Q{id}`
- 同一大题的多个小题，用相同 `section` 自动汇总

---

## 3. 哪些字段通常可以复用

通常先不动：

- `files`（含名单文件字段 `roster_file`）
- `roster`
- `llm.review_policy`
- 常见 `extract.source`（如 `choice_inline`、`section_question`、`code_table`）

只有文件路径、名单字段、模型策略真的变了，再改这些。

---

## 4. 什么时候说明版式变了

出现下面任一情况，就去看 `parsing` / `extract`：

- 区段标题变了（如“程序分析题”改名）
- 学号文件名格式变了
- 代码不在表格里了
- 题号格式变了（如 `1、` 变成 `第1题`）

优先改：

- `parsing.student_id_pattern`
- `parsing.sections.*`
- `extract.pattern`
- `extract.question_number`
- `extract.table_index`

---

## 5. 什么时候才改脚本

只有这几类情况再考虑改 Python：

- 现有提取方式完全不够用
- 代码题存放方式彻底变了
- 你想增强 Excel 输出结构

否则优先改配置。

---

## 6. 上线前 1 分钟检查

- [ ] 主观题模式是否正确：`llm.mode = llm_api` 或 `agent_runner`
- [ ] 如果使用 `agent_runner`，`llm.agent_backend` 是否已设置为 `claude` 或 `codex`
- [ ] 学号能提取
- [ ] 小题数量对得上
- [ ] 客观题答案已更新
- [ ] 主观题 rubric 已更新
- [ ] 同一大题的 `section` 文本完全一致
- [ ] 主观题请求不串题
- [ ] Excel 能输出 `Q1..Qn`
- [ ] Excel 能输出大题小计
- [ ] 四个大题小计之和等于总分

---

## 7. 不确定时看哪里

- 想看完整说明：`references/adapting-exam-config.md`
- 想看评分流程：`references/grading_workflow.md`
- 想看配置范式：`references/exam_config.template.json`
