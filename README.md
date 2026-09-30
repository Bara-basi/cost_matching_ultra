# 成本匹配工作台（cost_matching_ultra）

面向财务的成本核算自动化：把**报关单**拆成一条条可对账的记录，再从**睿贝 ERP**
找到对应的采购成本，分摊回每条记录，供财务核对与回填多维表格。

一句话链路：

```
报关单 PDF / 飞书记录 / 手工输入
        │
        ▼  解析 → 拆单 → 成本匹配
  按 采购订单 / 供应商 / 产品类型 拆成记录，采购金额按入库单实发分摊
        │
        ▼
  JSON 接口返回逐条结果（正常 / 需要复核）+ 可下载 Excel +（可选）写回飞书
```

> **界面状态**：Vue 财务工作台位于 `web/`，构建后由 FastAPI 在 `/` 提供页面。
> 工作台提供白天/夜间主题、轻量背景动效、异常聚焦、逐条核验、飞书原件区间处理、模板导入及写回预览。

**代码链路、模块职责、数据落盘位置、接口对接点**：见 [docs/代码结构与链路.md](docs/代码结构与链路.md)。

拆单与成本口径见 `docs/拆单说明.md`、`docs/产品类型与报关重量口径.md`、`docs/当前状态与异常台账.md`。

---

## 一、目录结构

```
app/                       后端服务（FastAPI + 业务服务）
  main.py                  服务入口：/api/*（暂无前端页面）
  services/                业务服务层
    parser/                报关单 PDF 解析
    pipeline.py            解析 → 拆单 → 成本匹配 的流水线（含逐单核验的拒核规则）
    contract_shipments.py  合同号 → 出运单定位
    shipment_index.py      出运产品行统一索引（拆单的数据底座）
    knapsack_split.py      子集和（背包）：把出运产品行分配到报关单
    cost_match.py          用入库单实发金额给拆单记录定成本
    grn_extract.py / grn_select.py  入库单附件解析与有效凭证挑选
    review.py              产出「财务看得懂」的逐条记录（异常类型 + 异常明细）
    workspace.py           任务编排（任务目录 / 状态 / 结果 / 下载）
    erp_sync.py            睿贝同步启动器（启动时按需同步）
    feishu_compare.py      飞书人工结果核对
    feishu_sync.py         飞书范围读取 + 写回
scripts/                   命令行脚本（解析、拆单、成本匹配、同步等）
docs/                      业务口径与说明文档
outputs/                   各阶段输出（Excel / 统计 / 每次运行的产物）
.cache/erp/                睿贝数据缓存（列表、明细、附件、索引、同步状态）
.runtime/jobs/             任务目录（状态 / 逐条明细 / 下载文件）
data/                      原始报关单 PDF、参考数据
```

---

## 二、环境准备

| 组件 | 版本 | 说明 |
| --- | --- | --- |
| Windows | 10/11 | 当前部署环境 |
| Python | 3.12 | 项目自带 `.venv` 虚拟环境 |

### 1. 安装后端依赖

```powershell
cd "E:\财务\cost matching\cost_matching_ultra"
& '.\.venv\Scripts\python.exe' -m pip install -r requirements.txt
```

> 新机器重建虚拟环境：
> ```powershell
> python -m venv .venv
> & '.\.venv\Scripts\python.exe' -m pip install -r requirements.txt
> ```
> 睿贝直连抓数（列表 / 订单明细）依赖 Playwright 浏览器，首次使用需执行一次：
> ```powershell
> & '.\.venv\Scripts\python.exe' -m playwright install chromium
> ```

### 2. 配置 `.env`

在项目根目录创建 `.env`（UTF-8），按需填写：

```ini
# 睿贝 ERP
ERP_API_KEY=              # MCP 服务令牌（AgentServerSimple）
ERP_USERNAME=             # 睿贝网页账号（列表 / 订单明细直连用）
ERP_PASSWORD=

# 飞书
LARK_APP_ID=
LARK_APP_SECRET=
MT_FINANCE_AI_COPY_APP_TOKEN=     # 目标表所在 base
MT_FINANCE_AI_COPY_TABLE_ID=      # 目标表（2026年报关记录 AI，可读可写）

# 远端写回由可信内网网关注入此密钥；不配置则工作台只允许预览，不允许写回
WORKBENCH_WRITE_TOKEN=
WORKBENCH_EVIDENCE_TOKEN=  # 可选：允许通过可信网关查看睿贝缓存凭证和附件对照结论

# 可选：AI 兜底（解析/核对）
DEEPSEEK_API_KEY=
```

> `.env` 已在 `.gitignore` 中，不会提交。

---

## 三、构建界面并启动服务

首次运行或修改 `web/` 后，先构建前端：

```powershell
cd 'E:\财务\cost matching\cost_matching_ultra\web'
npm install
npm run build
```

构建结果在 `web/dist/`，已加入 `.gitignore`。构建后重启 FastAPI，再访问根路径。
工作台没有用户登录；若启用飞书写回，部署方必须在受控内网通过可信网关为写回请求注入
`X-Workbench-Write-Token`，其值与服务端 `WORKBENCH_WRITE_TOKEN` 一致。不要把密钥放入前端代码。
同理，凭证详情和附件对照接口需要网关注入 `X-Workbench-Evidence-Token`。
未配置对应服务端密钥时，这些接口返回 403，界面显示“暂无法比较”。
若启用睿贝缓存凭证详情，网关还需向该受保护接口注入 `X-Workbench-Evidence-Token`。
操作员称呼只用于留痕，不代表已验证身份。

```powershell
cd "E:\财务\cost matching\cost_matching_ultra"

# 方式 A：相对路径不含空格，可以不写引号
.\.venv\Scripts\python.exe -m uvicorn app.main:app --host 127.0.0.1 --port 8765

# 方式 B：路径带引号时，PowerShell 必须加调用运算符 &，否则会报「意外的标记 -m」
& '.\.venv\Scripts\python.exe' -m uvicorn app.main:app --host 127.0.0.1 --port 8765
```

启动后访问 <http://127.0.0.1:8765/> 打开界面，<http://127.0.0.1:8765/api/health> 查看服务状态。
按 `Ctrl + C` 停止服务。

> **定时全量同步**：服务启动时按需同步一次（数据还新就跳过），之后后台调度线程**每 12 小时**再跑一轮：
> ① 刷新三类单据列表（按主键合并，只增不丢）→ ② 补抓缺失/过期的出运明细（MCP）
> → ③ 补抓采购附件里的入库单（成本匹配的唯一凭证，已下载的跳过）→ ④ 重建出运产品行索引。
> 全程在子进程里跑，不阻塞页面；进度见 `.cache/erp/sync_state.json`，日志见 `.cache/erp/reports/erp_sync.log`。
>
> 可用环境变量调整：`ERP_SYNC_ENABLED`（默认 1）、`ERP_SYNC_INTERVAL_HOURS`（默认 12）、
> `ERP_SYNC_REFRESH_DAYS`（出运明细重抓天数，默认 7）、`ERP_SYNC_ATTACHMENTS`（默认 1）。
> 需要把新入库单解析成金额时，手工执行 `scripts/erp_sync.py --with-grn-parse`。

---

## 四、接口一览

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| GET | `/api/health` | 服务状态、睿贝/飞书配置、上次同步状态 |
| GET | `/api/erp/sync` | 睿贝同步状态 |
| POST | `/api/erp/sync` | 触发一次睿贝同步（后台执行） |
| GET | `/api/sync/target` | 飞书目标表信息 |
| POST | `/api/lookup/parse` | 粘贴文本 → 识别成报关商品行（Form: `text`、`use_ai`） |
| POST | `/api/lookup` | 逐单核验（Form: `rows`=JSON 数组、`use_ai`），返回任务 id |
| POST | `/api/jobs` | 上传报关单文件建任务（multipart: `files`、`use_ai`） |
| GET | `/api/jobs/{id}` | 任务状态：`state` / `progress` / `message` / `summary` / `downloads` |
| GET | `/api/jobs/{id}/rows` | 逐条结果（供界面渲染） |
| POST | `/api/jobs/{id}/cancel` | 取消任务 |
| GET | `/api/jobs/{id}/download/{results\|exceptions}` | 下载 Excel |
| POST/GET | `/api/jobs/{id}/feishu-comparison` | 飞书人工结果核对（启动 / 读取） |
| POST | `/api/sync/scan` | 按范围读飞书记录并建任务（Form: `start_date` `end_date` `declarations` `contracts` `limit` `use_ai`） |
| GET | `/api/sync/{id}/plan` | 写回计划（将更新几条、几条需人工确认） |
| POST | `/api/sync/{id}/push` | 一键写回飞书（只写「正常」记录） |

新增接口：`GET /api/jobs` 查看任务；`GET /api/template` 下载模板；
`POST /api/template/check` 校验模板；`POST /api/lookup/with-evidence` 提交最多 50 条商品行与附件；
`POST /api/sync/preview` 只读预览飞书范围；`GET /api/jobs/{id}/input` 读取可修正输入；
`POST /api/lookup/parse` 优先解析飞书复制的制表符商品行，并从唯一原始 PDF 行与睿贝出运关系补齐缺失字段；
`POST /api/lookup/enrich` 可对手工输入再次补全，来源不唯一时只提示复核；
`POST /api/jobs/{id}/rerun` 重算；`POST /api/jobs/{id}/override` 记录人工调整；
`GET /api/jobs/{id}/audit` 查看操作记录。飞书区间使用副本表内未核算商品行直接计算，
按条件在飞书端筛选并按单批上限停止分页读取，不下载或解析 PDF；日期筛选会检索同合同号的其它商品行补齐计算上下文。
同报关单号、合同号和品名的历史拆分行先合并金额与重量，再重新核算；已有采购金额的历史组和补齐上下文行暂停自动写回。
写回时按报关单号、合同号、供应商简称和报关品名定位已有商品行，唯一匹配则更新，没有匹配则新增；重复或遗留旧行无法覆盖时暂停整组。
写回前校验报关金额守恒，并再次核对原商品行金额，发现变化即暂停。
无报关单号的内销行按合同号单独展示，本批出口核算暂不处理，不归类为报关单号格式异常。
若同一报关单的商品行有多个可能来源，可在写回预览中由财务明确指定来源行；
逐条核验上传的 PDF 附件会与输入金额、合同号、商品行数进行对照，附件解析失败不会中断成本匹配。
`/api/template` 下载的是从目标副本表读取的 112 列同名模板；可直接上传从该飞书副本表导出的 XLSX。
长报关单号须保留为文本，文件校验会拒绝被 Excel 当作数字改写的单号。

**约定**：`/api/sync/scan` 在范围内没有数据时不会报错，返回
`{"id": null, "selection": {...}, "message": "这段时间内没有可处理的数据（读取了 N 条记录）。"}`，
界面据此提示用户即可。

---

## 五、逐条结果字段说明

每条记录一行，只给**一个异常类型**和**一句异常明细**（给财务看的，不写技术细节）：

| 字段 | 说明 |
| --- | --- |
| 报关单号 / 合同号_1 | 与报关单、飞书保持同一写法 |
| 合同号（应收表格） | 采购订单去掉工厂简称（如 `26MT-03T094Y`） |
| 外销订单号 | 报关合同号 |
| 产品类型 | 粗分类：法兰 / 管件 / 无缝管 / 镍基无缝管 / 焊管 / 焊材 / 板棒 / 盘管 / 三角丝 / 其他 |
| 采购金额 | 入库单实发金额按本行出运采购金额(RMB)占比摊得 |
| 报关金额 | 报关金额（含客户费用分摊） |
| 报关重量 | 该报关行的重量分到本行（单条直接回填；多条按入库单附件重量，取不到则按报关金额，末行吸收尾差） |
| 异常类型 / 异常明细 | 正常 / 缺报关单 / 缺入库单 / 睿贝未定价 / 缺重量 / 产品类型缺失 / 出运单有额外费用 |
| 采购订单号 / 合同号_1 / 飞书记录 | 睿贝订单源与飞书原记录入口 |

**正常/异常判定口径**

- 只有「金额算漏」才算异常；同一采购组内分摊位置不同不算异常。
- 与飞书核对时：采购金额相同即正常；只有组内分配不同记「正常，但额外金额分配不一致」；
  报关金额、采购金额允许行级不同，合计一致即正常。
- 逐单核验时，若某出运单的报关单不全**且**客户费用对不上，直接拒核并提示
  「该出运单有额外费用，需要出运单对应的报关单才能确认分配规则」；
  其它入口遇到同类数据只把该条记录标为异常，不影响整批。

---

## 六、命令行全流程（不使用服务时）

```powershell
# 1) 同步睿贝数据（列表 + 出运明细补漏 + 重建索引）
& '.\.venv\Scripts\python.exe' scripts\erp_sync.py

# 2) 解析报关单 PDF → outputs\customs_parse\报关单解析结果_出口退税联.xlsx
& '.\.venv\Scripts\python.exe' scripts\parse_customs.py

# 3) 按产品编码补商品资料（产品类型回填用，断点续跑）
& '.\.venv\Scripts\python.exe' scripts\fetch_product_categories.py

# 4) 拆单 → outputs\shipments_split\
& '.\.venv\Scripts\python.exe' scripts\run_split_knapsack.py
& '.\.venv\Scripts\python.exe' scripts\run_split_knapsack.py --only 25MT-07F521Y,26MT-01T109

# 5) 成本匹配（入库单实发口径）→ outputs\cost_match\
& '.\.venv\Scripts\python.exe' scripts\build_cost_match.py

# 6) 本地核对（只按本地事实判定正常/异常，不与飞书比对）→ outputs\local_review\
& '.\.venv\Scripts\python.exe' scripts\build_local_review.py
```

---

## 七、输出说明

| 目录 | 内容 |
| --- | --- |
| `outputs/customs_parse/` | 报关单解析结果（出口退税联参与成本匹配，预录单只留档） |
| `outputs/shipments_split/` | 拆单明细、正常/异常/已核实、未认领产品行、本地异常 |
| `outputs/cost_match/` | 成本匹配全部/正常/异常/待核实/已核实 |
| `outputs/local_review/` | 本地核对（全部/正常/异常/汇总） |
| `outputs/runs/<任务号>/` | 每次运行的临时产物（拆单明细等） |
| `.runtime/jobs/<任务号>/` | 任务状态、逐条明细、可下载 Excel |
| `.cache/erp/` | 睿贝数据缓存（列表 / 明细 / 附件 / 索引 / 同步状态） |

---

## 八、常见问题

**Q：打开根路径提示前端未构建**
在 `web/` 运行 `npm install`、`npm run build`，然后重启服务。

**Q：启动时同步跑很久**
首次同步或数据过期时会全量刷新三类单据列表，属正常；之后只做增量补漏。
不想在启动时同步，可临时把 `app/services/erp_sync.py` 的启动调用注释掉，改为手工执行 `scripts/erp_sync.py`。

**Q：睿贝抓不到数据**
检查 `.env` 的 `ERP_API_KEY`（MCP）与 `ERP_USERNAME/ERP_PASSWORD`（网页直连），
并确认已执行 `playwright install chromium`。日志见 `.cache/erp/reports/erp_sync.log`。

**Q：飞书核对 / 一键同步报「未配置」或「没有权限」**
检查 `.env` 的 `LARK_APP_ID`、`LARK_APP_SECRET` 与目标表 `MT_FINANCE_AI_COPY_APP_TOKEN`、
`MT_FINANCE_AI_COPY_TABLE_ID`，并确认应用有该 base 的读写权限（其它 base 可能返回 403）。

**Q：某条记录显示「缺报关单」**
说明这张采购单还有报关单没解析到（多为 PDF 未提供）。补齐对应报关单后重跑即可。
