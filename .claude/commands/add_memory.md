## 写入规范 `add_memory`

### 1. 何时写
| 类型 | 立即写入 |
|------|----------|
| **Preference** 偏好 | 用户明确长期喜好或禁止事项 |
| **Procedure** 流程 | 用户要求的固定步骤 / SOP |
| **新事实**        | 实体之间出现新/变动关系（用/换/依赖/影响…） |

### 2. 如何写
- **一句事 + 一个原因**：若文本较长，**分段** 写入；每段只表达一件事。  
- **显式关键词**：若是更新/替换请包含 `update` / `replace(s)` 等字样，并在正文提到 **旧值**。
- **原因/影响**：用 `why:` 或 `impact:` 明示，可提高 LLM 抽取质量。  
- **参考时间**：尽量传 `reference_time="2025-06-16T09:30:00Z"`；缺省则取当前时间，并且具体的的时间要写在name中。  
- **大型附件**：只写「结论 + 指针」；例如 `doc_url`, `commit_sha`, `file_path` 字段指向原文件。

#### 文本示例（短）
```
[Preference] update: replace MPU-6050 with BMI270
why: 30 % lower current, smaller package
```

#### JSON 示例（结构化）
```python
add_memory(
  name="Sensor swap 2025-06-16T09:30:00Z",
  episode_body="{\"type\":\"Preference\",\"old_sensor\":\"MPU-6050\",\"new_sensor\":\"BMI270\",\"reason\":\"lower current\"}",
  source="json",
  source_description="design review #42",
  reference_time="2025-06-16T09:30:00Z",
  group_id="Radon_Monitor"
)
```