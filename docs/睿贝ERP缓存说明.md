# 睿贝 ERP 数据缓存说明

目标：一次性把睿贝 ERP 的单据贪婪抓下来，之后成本匹配环节**不再重复联网**，
同时避开「同一账号只能一人登录」的限制。

## 取数方式

不直接爬 DOM，而是**复用睿贝自己的内部接口**：
用 Playwright 登录一次拿到会话，然后在页面上下文里重放它的 XHR 列表接口
（Cookie 与登录态自动带上）。这样比点页面稳，也绕开了「按采购单/按出运单」按钮的异步 bug。

已确认的接口（均在 `app/services/erp_lists.py`）：

| 用途 | 接口 | 关键参数 |
| --- | --- | --- |
| 外销订单列表 | `POST /saleOrder_list` | `condition=all`，`contractDate_start/end` 可选 |
| 出运单列表 | `POST /shipment_select` | `condition=all`，`userDefaultTableName=biz_shipment` |
| 采购单列表 | `POST /purchase_selectPur` | `condition=all`，`purchaseIsOutSale=all`（全量 1951） |
| 采购单列表（仅含入库附件） | `POST /purchase_selectPur` | 追加 `grnChoose=Y`（134 条） |
| 某订单的出运明细 | `POST /shipmentItem_selectShipFollow` | `order_id`、`userDefaultTableName=salefollow_shipment` |
| 某订单的采购明细 | `POST /purchaseItems_listPlaceOrder` | `order_id` |

分页：每页固定 16 条，参数 `p`。

### 返回结构（列式表格）

```json
{"total": 1951, "root": [[ {"beanName":"purchase_code,purchase_id,tickId",
                           "columnValues":["26MT-05Q461-GYL-供应链","3003",""]}, ... ]]}
```

`beanName` 是**与 `columnValues` 一一对应的字段名列表**（不一定是每字段一个值），
解析逻辑在 `app/services/erp_replay.py::columns_to_rows`。
另外 `shipment_select` 返回的是**非标准 JSON**（key 无引号），由 `parse_lenient_json` 兼容。

## 缓存目录（面向代码读取）

```
.cache/erp/
  meta.json                       抓取时间、条数、参数
  sale_orders/
    orders.jsonl                  外销订单，一行一条 JSON（3668 条）
    index.json                    orderCode -> 订单摘要
    samples/*.json                少量缩进样例，供人工查看
  shipments/
    shipments.jsonl               出运单（1199 条）
    index.json                    invoiceCode/purchaseCode -> 出运摘要
  purchases/
    purchases.jsonl               采购单（1951 条）
    index.json                    purchase_code -> 采购摘要
    with_grn.jsonl                仅含入库单附件的那批（134 条，用于附件下载）
  details/orders/<订单号>.json    每个订单的产品行明细（出运行 + 采购行）
  attachments/
    lists/<采购单号>.json         该采购单的完整附件清单（经 MCP purchase.find）
    grn/<采购单号>/*.xlsx         下载下来的「入库单」文件
    grn_manifest.json             已下载入库单的登记表
    downloads/<采购单号>.zip      早期用网页「打包下载」得到的 ZIP（已弃用）
  index/order_index.json          统一关联索引（单号 -> 订单/出运/采购）
  links/order_graph.json          关联图（便于「一查二」）
  reports/                        各类核验报告（互查覆盖度、附件清单等）
```

2026 年数据量：外销订单 515、出运单 441、采购单 799。

## 附件取数（重要：走 MCP，不要走网页打包下载）

**踩过的坑**：一开始用网页上的「打包下载」按钮取 ZIP，并且用
`grnChoose=Y` 过滤采购单——结果只覆盖 134/1951 个采购单，且 ZIP 里
几乎没有入库单。两点都不对：

1. `grnChoose=Y` 只筛出「待处理入库」的采购单（134 条），
   而实际**发生过入库的采购单有 1807 条**（判据：`grnQuantity > 0` 或 `lastGrnDate` 非空）。
2. 网页「打包下载」拿到的 ZIP 不含全部附件；改用 MCP 后可直接拿到每个附件的下载地址。

**正确做法**（`app/services/erp_attachment_mcp.py` + `scripts/erp_fetch_grn_mcp.py`）：

```python
# MCP purchase.find 返回每个附件的名字与下载地址
{"attachmentList": [
  {"attachmentName": "26MT-03T271-A 鸿迪 入库单（回传）.xlsx",
   "downloadUrl": "https://erp.mtholdinggroup.com/tempFile/552719B9522A81891391C52C01E5C92A"},
  ...]}
```

带浏览器会话直接 GET 该 `downloadUrl` 即可下载（实测 200，内容正确），
只保留文件名含「入库单」的附件。清单一律缓存到 `attachments/lists/`，
下次运行自动跳过，支持断点续跑。

```powershell
& '.\.venv\Scripts\python.exe' scripts\erp_fetch_grn_mcp.py            # 全量
& '.\.venv\Scripts\python.exe' scripts\erp_fetch_grn_mcp.py --limit 5  # 试跑
& '.\.venv\Scripts\python.exe' scripts\erp_fetch_grn_mcp.py --force    # 重下
```

实测覆盖率（1807 个发生过入库的采购单）：**含入库单的约 98%**，与网页上人工看到的一致。

## 产品行明细（成本匹配真正要用的数据）

按订单抓取，落在 `details/orders/<订单号>.json`：

```json
{"_orderId": "6165",
 "shipment_items": [{"shipmentId": "...", "invoiceCode": "PI-26MT-03T428Y",
                     "itemList.code": "P26081866", "itemList.quantity": "1",
                     "itemList.price": "600.0", "itemList.amount": "600.0",
                     "itemList.sizeEn": "...", "itemList.countQuatity": "1.0"}],
 "purchase_items": [{"purchaseCode": "26MT-03T428Y-希蒙", "comName": "浙江希蒙雷斯钢管有限公司",
                     "itemList.code": "P26081866", "itemList.price": "1786.36",
                     "itemList.amount": "1786.36", "itemList.grnQuantity": "1.0"}]}
```

2026 年 515 个订单全部已抓（断点续跑：`scripts/erp_fetch_details.py --year 2026`）。

## 三单互查（只用睿贝原生关联字段）

```python
from app.services.erp_crossref import build
cross = build()
cross.lookup("PI-26MT-06C309")            # 订单 -> 出运单 + 采购单
cross.orders_of_shipment("559")           # 出运单 -> 它包含的所有订单
cross.purchases_of_shipment("559")        # 出运单 -> 采购单
cross.shipments_of_purchase("691")        # 采购单 -> 出运单
```

**能否做到三单互查：可以，而且用的是睿贝自己存的外键**，不做单号前缀猜测：

| 关系 | 原生证据 |
| --- | --- |
| 订单 ↔ 出运 | 明细接口 `/shipmentItem_selectShipFollow?order_id=<orderId>` 返回行里的 `shipmentId` |
| 订单 ↔ 采购 | 明细接口 `/purchaseItems_listPlaceOrder?order_id=<orderId>` 返回行里的 `purchaseId` |
| 出运 ↔ 采购 | 出运单头 `purchaseCode`（存的是它包含的采购单号，用 `,` 连接） |
| 出运 ↔ 订单（多单出运） | 出运单头 `orderCode`（同一订单号的两种写法只做大小写/`PI-` 前缀归一） |

明细接口的返回是「一张单据 + 若干产品行」：首行带 `shipmentId`/`purchaseId`，
同单据后续产品行的该字段为空，按空值向下继承即可，不涉及任何字符串推断。

实测覆盖（3540 个订单 / 1198 张出运单 / 1830 个采购单的明细已全量抓取）：

| 指标 | 结果 |
| --- | --- |
| 订单能定位到出运单 | 1126 |
| 订单能定位到采购单 | 1263 |
| 订单同时有出运与采购 | 1126 |
| 出运单能反查到订单 | 1183 / 1198 |
| 出运单能反查到采购单 | 1183 / 1198 |
| 采购单能反查到出运单 | 1645 / 1830 |
| 多单出运（一张出运单含多个订单） | 118 |
| 半包含订单（跨多张出运单） | 155 |

**半包含订单**：`shipments_of_order()` 返回该订单关联的全部出运单（原生 `shipmentId`），
例如 `PI-26MT-03R360` 跨 15 张、`PI-26MT-02N182` 跨 10 张、`PI-25MT-07F477` 跨 8 张，
配合订单行数量与出运行数量即可算出「还差多少」。

## 一查二：按任意单号找另外两张单

```python
from app.services.erp_cache_index import lookup

lookup("PI-26MT-03T428Y")
# {
#   "orderCode": "PI-26MT-03T428Y",
#   "order":     {"orderId": "6165", "comName": "Special Piping Materials Ltd", ...},
#   "shipments": [{"shipmentId": "1722", "invoiceCode": "PI-26MT-03T428Y",
#                  "purchaseCode": "26MT-03T428Y-希蒙", "shipDate": "2026-09-04", ...}],
#   "purchases": [{"purchase_id": "2893", "purchase_code": "26MT-03T428Y-希蒙",
#                  "supplierName": "浙江希蒙雷斯钢管有限公司", "amount": "1786.36"}]
# }
```

索引规模：6314 个节点，其中 3541 有订单、3276 有采购、2660 有出运，277 个三者齐全。

## 常用命令

```powershell
& '.\.venv\Scripts\python.exe' scripts\erp_fetch_all.py list --year 2026   # 抓三类列表（已有则跳过）
& '.\.venv\Scripts\python.exe' scripts\erp_fetch_all.py list --force       # 强制重抓
& '.\.venv\Scripts\python.exe' scripts\erp_build_index.py                  # 重建关联索引
& '.\.venv\Scripts\python.exe' scripts\erp_cache_check.py                  # 检查缓存完整性
```

## 后续

- 抓取脚本幂等：`--force` 才会重抓，未加时优先复用已有 `jsonl`。
- 入库单解析（取「实发数据」结算金额）是成本核算的下一步。
- 附件抓取可放后台跑：
  ```powershell
  Start-Process -FilePath ".\.venv\Scripts\python.exe" `
    -ArgumentList @("scripts\erp_fetch_grn_mcp.py") -WorkingDirectory (Get-Location) `
    -RedirectStandardOutput ".cache\erp\reports\grn_fetch.out.log" `
    -RedirectStandardError  ".cache\erp\reports\grn_fetch.err.log" `
    -PassThru -WindowStyle Hidden
  ```
  日志在 `.cache/erp/reports/grn_fetch.out.log`，进度看 `attachments/lists/` 的文件数。
