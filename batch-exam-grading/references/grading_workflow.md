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
| `llm_cache.jsonl` | 主观题评分缓存，同请求直接复用 |

## 二、先配 `exam_config.json`

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

检查 `llm_requests.jsonl`：

- 题号是否完整
- `student_answer` 是否串题
- 主观题是否只来自配置里启用 `llm.enabled=true` 的题
- 同一条请求是否带 `request_hash`

## 四、评分

### 4.1 主观题模型评分

```powershell
python llm_grade.py --config exam_config.json --mode openai --model <model>
```

相同 `request_hash` 会直接复用缓存，不再重新调用模型。

### 4.2 合并成绩

```powershell
python grade_exam.py --config exam_config.json
```

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
| 主观题无法合并 | 缺少 `llm_grades.jsonl` 或 request_hash 不匹配 | 先跑 `llm_grade.py --mode openai`，必要时清理旧缓存 |
| 名单成绩对不上 | 名单字段映射错误 | 改 `roster.student_id_field` 等字段 |
