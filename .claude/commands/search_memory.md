
## 检索流程（先找再答）

> **黄金法则：先 `nodes` → 再 `facts` → 组合回答**  

### 0️⃣ 先协商 `group_id`（首要前置）
```
“为了精确检索，请告诉我当前的 group_id（项目/对话命名空间）？”
```
- 在得到明确回应前 **禁止任何检索调用**。  
- 调用工具时 **始终以字符串数组传递**：`["Radon_Monitor"]`。

### 1️⃣ 默认先查节点 `search_memory_nodes`
```json
{
  "query": "<用户关键词>",
  "group_ids": ["<group_id>"],
  "max_nodes": 15          // 可视任务大小增减，默认 10
}
```
- ✅ **可选**：若只关心 **偏好** 或 **流程**，可直接加  
  `entity: "Preference"` / `"Procedure"` 进行精准过滤。  
- ⚠️ 其他实体类型如 `Requirement`、`Component` 仅在 **已在服务器注册自定义模型** 后才可用。

### 2️⃣ 评估节点结果
| 场景 | 行动 |
|------|------|
| A. 找到且 `summary/attributes` 已能回答「是什么」| **直接作答**，流程结束 |
| B. 问题涉及「关系/交互/因果」，节点不够 | 记录相关节点 `uuid` → 使用 `search_memory_facts` 补充 |
| C. 结果为空 | 直接 `search_memory_facts` |

### 3️⃣ 按需查事实 `search_memory_facts`
- **精准版**（来自场景 B：已知中心节点）
  ```json
  {
    "query": "<用户关键词>",
    "group_ids": ["<group_id>"],
    "center_node_uuid": "<记录的uuid>",
    "max_facts": 20
  }
  ```
- **广泛版**（场景 C：节点为空）
  ```json
  {
    "query": "<用户关键词>",
    "group_ids": ["<group_id>"]
  }
  ```

### 4️⃣ 综合并回答
- 把 `nodes` 与 `facts` 中的 `summary` / `fact` 句子融合成完整答复。  
- 如果引用事实，请明确「截至 YYYY-MM-DD」或「在 as_of 时间点」说明时效。