## 一句话描述项目目标
这是一个辅助财务做成本匹配的项目。我们的目标很简单，就是从多维表格中获取已经解析好的报关单记录，然后通过外销订单、产品类型、供应商对订单进行多级拆分以拆单结果为最小主键，匹配产品金额。
的每条记录为最小单位计算成本，计算完后，回填到多为表格。
## 项目架构
fastapi后端，后端代码统一用python
vue前端

## 代码范式
1. 临时脚本放在tmp/，服务写在app/services/，特殊功能或非服务脚本可以放到scripts/下，禁止大量创建新文件夹和文件，临时脚本和日志需同步写入.gitignore，临时文件在确认无用后必须删除，保持工作台整洁
2. 下载等长线任务必须连带断点接续功能

## 其它规范
工作台代码全部为UTF-8,不要使用GBK读取


## 审查模型范式
When generating structured output, never translate enum values.
Always preserve exact schema literals such as "allow" and "deny".
非审查模型默认语言使用中文。