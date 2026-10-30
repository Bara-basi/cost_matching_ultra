# 报关 PDF 文件命名接口

## 1. 用途

财务电脑上传报关 PDF 后，服务器负责解析合同号、报关单号和日期，并按财务规范返回每个 PDF 的目标文件名：

```text
YYYYMM_报关单上的合同号_报关单上的报关单号后四位.pdf
```

财务电脑不需要解析 PDF。收到映射后，应在本地另存一份使用规范名称的 PDF。

月份优先取报关单的 `出口日期`，缺失时回退 `申报日期`。合同号优先取解析结果中的原始报关合同号 `合同号_1`。

Windows 文件名不允许出现的字符（`\ / : * ? " < > |`）会替换为 `_`；其余合同号内容保持不变。

## 2. 上传接口

```http
POST /api/customs/tasks/{taskId}/upload
Authorization: Bearer <CUSTOMS_AGENT_TOKEN>
X-Agent-Name: finance-pc-01
Content-Type: multipart/form-data
```

Multipart 字段：

| 字段 | 必填 | 说明 |
|---|---|---|
| `files` | 是 | PDF，可重复出现；发送顺序就是本地文件保存顺序 |
| `agent_name` | 建议 | 与 `X-Agent-Name` 一致 |

服务器按 multipart `files` 字段的实际接收顺序处理，并把顺序写入任务目录下的 `upload_manifest.json`。因此不依赖文件名排序或文件系统枚举顺序。

接口会等待 PDF 解析阶段完成再返回命名表。扫描解析等后续工作仍在服务器后台继续，所以成功状态码仍为：

```text
202 Accepted
```

### `source=test` 解析测试任务

任务创建时如果 `source` 为 `test`，该任务只执行 PDF 解析和文件命名。上传接口返回响应体后任务直接进入 `completed`，不会继续拆单、成本匹配或写回飞书。

测试任务响应中的状态示例：

```json
{
  "taskId": "ct_0123456789abcdef",
  "state": "completed",
  "fileCount": 1,
  "documentNamesReady": true,
  "documentNames": [
    "202610_26MT-01A100_8477.pdf"
  ],
  "nameMap": {
    "a.pdf": "202610_26MT-01A100_8477.pdf"
  }
}
```

任务摘要中的 `mode` 为：

```json
{
  "mode": "test-parse-only",
  "aiRecords": 0
}
```

`source` 不是 `test` 的任务仍按完整链路执行。

### 固定主动上传 taskid

财务电脑也可以不执行 `claim`，直接使用固定的测试 taskid 调用上传接口：

```text
ct_49e8dc49d0ce48c1
```

服务端 `.env` 可通过以下配置修改：

```dotenv
CUSTOMS_FIXED_TEST_TASK_ID=ct_49e8dc49d0ce48c1
```

固定 taskid 的特点：

- 初始任务来源固定为 `test`；
- 不参与普通 Agent 的 `/claim` 队列轮询；
- 仍必须携带 Bearer `CUSTOMS_AGENT_TOKEN`；
- 可直接调用上传接口，无需先领取；
- 每次上传只执行 PDF 解析和规范命名；
- 返回响应体后任务进入 `completed`，不拆单、不匹配成本、不写飞书；
- 上一次任务进入 `completed` 或 `failed` 后，可以向同一个 taskid 再次上传；
- 同一 taskid 正在 `processing` 时，新的上传返回 HTTP `409`。

固定 taskid 示例：

```http
POST /api/customs/tasks/ct_49e8dc49d0ce48c1/upload
Authorization: Bearer <CUSTOMS_AGENT_TOKEN>
X-Agent-Name: finance-pc-01
Content-Type: multipart/form-data
```

该请求不需要先调用 `/api/customs/tasks/claim`。

## 3. 成功响应

```json
{
  "taskId": "ct_0123456789abcdef",
  "state": "processing",
  "fileCount": 2,
  "documentNamesReady": true,
  "statusUrl": "/api/customs/tasks/ct_0123456789abcdef",
  "documentNamesUrl": "/api/customs/tasks/ct_0123456789abcdef/document-names",
  "orderBasis": "multipart-files-field-received-order",
  "documentNames": [
    "202610_26MT-01A100_8477.pdf",
    "202610_26MT-02B200_5999.pdf"
  ],
  "documents": [
    {
      "index": 1,
      "originalFileName": "a.pdf",
      "storedFileName": "a.pdf",
      "newFileName": "202610_26MT-01A100_8477.pdf",
      "monthDateField": "出口日期",
      "error": null
    },
    {
      "index": 2,
      "originalFileName": "b.pdf",
      "storedFileName": "b.pdf",
      "newFileName": "202610_26MT-02B200_5999.pdf",
      "monthDateField": "申报日期",
      "error": null
    }
  ],
  "nameMap": {
    "a.pdf": "202610_26MT-01A100_8477.pdf",
    "b.pdf": "202610_26MT-02B200_5999.pdf"
  }
}
```

字段说明：

| 字段 | 说明 |
|---|---|
| `documentNames` | 按上传顺序排列的文件名数组；无法命名时对应位置为 `null` |
| `documents` | 逐文件详细信息，包含接收序号、原名、服务器保存名、新名和错误 |
| `nameMap` | `storedFileName -> newFileName` 映射，适合财务电脑逐项重命名 |
| `originalFileName` | 财务电脑上传时使用的原文件名 |
| `storedFileName` | 服务器安全保存后的文件名；仅在服务器生成同名文件时可能与原名不同 |
| `monthDateField` | 命名时实际使用的是 `出口日期` 还是 `申报日期` |
| `error` | 无法命名的原因；非空时不要用空名覆盖本地文件 |

`documentNames` 与请求中的 `files` 顺序严格一一对应。财务电脑处理重名、重传或本地已存在文件时应采用自己的受控冲突策略，但不得改变服务器返回的规范名称本身。

## 4. 命名失败项

单个 PDF 无法命名时，服务仍返回其余文件的名称，并在 `documents[].error` 中给出原因。常见原因包括：

- PDF 无法解析；
- 未识别出合同号或报关单号；
- 出口日期和申报日期都缺失；
- 一个 PDF 解析出多个合同号或报关单号，无法唯一命名。

财务电脑应保留命名失败项的原文件及错误信息，不要把它当作空文件名或改名成功。

## 5. 等待超时后补取

上传接口默认最多等待 300 秒。服务器 `.env` 可调整：

```dotenv
CUSTOMS_DOCUMENT_NAME_WAIT_SECONDS=300
```

若超时，仍返回 HTTP `202`，但响应中为：

```json
{
  "taskId": "ct_0123456789abcdef",
  "state": "processing",
  "fileCount": 20,
  "documentNamesReady": false,
  "statusUrl": "/api/customs/tasks/ct_0123456789abcdef",
  "documentNamesUrl": "/api/customs/tasks/ct_0123456789abcdef/document-names"
}
```

解析完成后可补取：

```http
GET /api/customs/tasks/{taskId}/document-names
Authorization: Bearer <CUSTOMS_AGENT_TOKEN>
```

补取响应：

```json
{
  "taskId": "ct_0123456789abcdef",
  "state": "processing",
  "documentNamesReady": true,
  "orderBasis": "multipart-files-field-received-order",
  "documentNames": [
    "202610_26MT-01A100_8477.pdf"
  ],
  "documents": [
    {
      "index": 1,
      "originalFileName": "a.pdf",
      "storedFileName": "a.pdf",
      "newFileName": "202610_26MT-01A100_8477.pdf",
      "monthDateField": "出口日期",
      "error": null
    }
  ],
  "nameMap": {
    "a.pdf": "202610_26MT-01A100_8477.pdf"
  }
}
```

接口尚未生成命名表时返回：

```json
{
  "taskId": "ct_0123456789abcdef",
  "state": "processing",
  "documentNamesReady": false
}
```

客户端应继续轮询，不要重新上传 PDF。

## 6. 服务端持久化文件

任务目录：

```text
.runtime/customs_tasks/<taskId>/
```

新增保留文件：

| 文件 | 内容 |
|---|---|
| `upload_manifest.json` | 上传接收顺序、原名、服务器保存名、大小和 SHA-256 |
| `document_names.json` | 解析后的完整命名映射 |

任务结束清理时只删除 `uploads/` 和 `work/`，上述两个文件与 `task.json`、`result.json` 一起保留。

## 7. 错误处理

| HTTP | 含义 | 客户端处理 |
|---|---|---|
| 202 | PDF 已接收；命名表可能已返回或仍需补取 | 检查 `documentNamesReady` |
| 400 | PDF 无效、数量/大小不符合要求 | 不重试同一文件，修正后重传 |
| 403 | Agent Token 错误 | 检查财务电脑本地密钥 |
| 404 | 任务不存在 | 检查 `taskId` |
| 409 | 任务状态冲突、已处理或已被其他电脑领取 | 查询任务状态，不盲目重传 |
| 413 | 文件超过限制 | 重新拆分业务任务 |

