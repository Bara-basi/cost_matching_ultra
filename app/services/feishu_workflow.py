"""从飞书副本商品行筛选未核算输入，并保守地规划拆单写回。"""
from __future__ import annotations

import hashlib
import itertools
import json
import re
import threading
from datetime import datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation

from app.services.feishu_client import FeishuClient, workbench_table
from app.services import feishu_compare, workspace
from app.services.scope import out_of_scope_reason

META_SOURCE = "成本匹配源记录ID"
META_KEY = "成本匹配拆分键"
META_ORIGINAL = "成本匹配原始报关金额"
_PUSH_LOCK = threading.Lock()
CHINA_TZ = timezone(timedelta(hours=8))


def _target():
    app, table = workbench_table(required=False)
    if not app or not table:
        raise ValueError("飞书目标表未配置")
    return FeishuClient(), app, table


def _text(value):
    return feishu_compare.text(value)


def _date(value):
    if isinstance(value, (int, float)):
        return datetime.fromtimestamp(value / 1000, tz=CHINA_TZ).strftime("%Y-%m-%d")
    raw = _text(value)
    if re.fullmatch(r"\d{13}", raw):
        return datetime.fromtimestamp(int(raw) / 1000, tz=CHINA_TZ).strftime("%Y-%m-%d")
    match = re.match(r"(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})", raw)
    return f"{int(match[1]):04d}-{int(match[2]):02d}-{int(match[3]):02d}" if match else ""


def _tokens(value):
    return {v.strip().upper() for v in re.split(r"[,&，;；\s]+", value or "") if v.strip()}


def _condition(name, operator, values):
    return {"field_name": name, "operator": operator, "value": values}


def _search_many(client, app, table, name, values, operator="is"):
    tokens = sorted({str(value).strip() for value in values if str(value).strip()})
    if not tokens:
        return []
    result = {}
    for start in range(0, len(tokens), 30):
        group = tokens[start:start + 30]
        filter = {"conjunction": "or", "conditions": [_condition(name, operator, [value]) for value in group]}
        for record in client.search_records(app, table, filter):
            result[record["record_id"]] = record
    return list(result.values())


def _date_filter(name, raw, start):
    day = datetime.strptime(raw, "%Y-%m-%d").replace(tzinfo=CHINA_TZ)
    moment = day if start else day + timedelta(days=1)
    millis = int(moment.timestamp() * 1000)
    return _condition(name, "isGreater" if start else "isLess",
                      ["ExactDate", str(millis - 1 if start else millis)])


def _record_input(record):
    fields = record.get("fields") or {}
    names = ("报关单号", "合同号_1", "商品序号", "报关品名", "海关编码", "报关金额",
             "报关重量", "申报数量", "申报单位", "币种", "产品类型", "采购金额", "供应商简称")
    item = {name: _text(fields.get(name)) for name in names}
    item.update(record_id=record.get("record_id") or "",
                pdfTokens=[str(pdf.get("file_token")) for pdf in fields.get("pdf原件") or []
                           if isinstance(pdf, dict) and pdf.get("file_token")])
    return item


def _input_reason(item):
    if not item["报关单号"]:
        return "内销记录，无需报关单号；本批出口核算暂不处理"
    if not re.fullmatch(r"\d{18}", item["报关单号"]):
        return "出口记录的报关单号需为 18 位数字"
    if not item["合同号_1"] or not item["报关品名"]:
        return "合同号或报关品名为空"
    # 明显不合合同号规范的历史遗留单（SP-/CY-/ZY-/xxSM-/24 年及更早）直接跳过，
    # 与 pipeline.write_declarations 的拦截口径保持一致（否则会「选中 N 条、实际只算 M 条」）。
    scope = out_of_scope_reason(item["合同号_1"])
    if scope:
        return f"{scope}，本项目不处理"
    amount = _amount(item["报关金额"])
    if amount is None or amount <= 0:
        return "报关金额缺失或不是正数"
    return ""


def _item_key(item):
    return (item["报关单号"], item["合同号_1"], item["报关品名"])


def select_records(*, start_date="", end_date="", declarations="", contracts="", limit=50):
    for value in (start_date, end_date):
        if value:
            try:
                datetime.strptime(value, "%Y-%m-%d")
            except ValueError as exc:
                raise ValueError("日期须为 YYYY-MM-DD") from exc
    if start_date and end_date and start_date > end_date:
        raise ValueError("起始日期不能晚于结束日期")
    if not 1 <= int(limit) <= 200:
        raise ValueError("单批上限应为 1 至 200")
    client, app, table = _target()
    wanted_decl, wanted_contract = _tokens(declarations), _tokens(contracts)
    conditions = [_condition("采购金额", "isEmpty", [])]
    if start_date:
        conditions.append(_date_filter("出口日期", start_date, True))
    if end_date:
        conditions.append(_date_filter("出口日期", end_date, False))
    if len(wanted_decl) == 1:
        conditions.append(_condition("报关单号", "is", [next(iter(wanted_decl))]))
    elif not wanted_decl and len(wanted_contract) == 1:
        conditions.append(_condition("合同号_1", "contains", [next(iter(wanted_contract))]))
    if len(wanted_decl) > 1:
        searches = ({"conjunction": "and", "conditions": conditions + [_condition("报关单号", "is", [value])]}
                    for value in sorted(wanted_decl))
    elif not wanted_decl and len(wanted_contract) > 1:
        searches = ({"conjunction": "and", "conditions": conditions + [_condition("合同号_1", "contains", [value])]}
                    for value in sorted(wanted_contract))
    else:
        searches = iter(({"conjunction": "and", "conditions": conditions},))
    candidates = itertools.chain.from_iterable(client.iter_search_records(app, table, query) for query in searches)
    selected, skipped = [], []
    scanned = 0
    seen_candidates = set()
    for record in candidates:
        if record.get("record_id") in seen_candidates:
            continue
        seen_candidates.add(record.get("record_id"))
        scanned += 1
        fields = record.get("fields") or {}
        decl, contract = _text(fields.get("报关单号")), _text(fields.get("合同号_1"))
        if wanted_decl and decl.upper() not in wanted_decl:
            continue
        if wanted_contract and not (_tokens(contract) & wanted_contract):
            continue
        if start_date or end_date:
            date = _date(fields.get("出口日期"))
            if not date or (start_date and date < start_date) or (end_date and date > end_date):
                continue
        if _text(fields.get(META_SOURCE)) or _text(fields.get("采购金额")):
            continue
        item = _record_input(record)
        reason = _input_reason(item)
        if reason:
            skipped.append({"record_id": item["record_id"], "报关单号": decl,
                            "合同号_1": contract, "tradeType": "内销" if not decl else "外销",
                            "reason": reason})
            continue
        selected.append(item)
        if len(selected) >= int(limit):
            break
    if not selected:
        return {"selected": [], "context": [], "skipped": skipped, "scanned": scanned,
                "selection": {"total": 0, "declarations": 0, "supplemental": 0,
                              "merged": 0, "skipped": len(skipped), "scanned": scanned}}

    # 同单同商品的历史拆分仅作兜底：合并计算，同时阻止自动写回。
    related = {item["record_id"]: item for item in _search_many(
        client, app, table, "报关单号", {item["报关单号"] for item in selected})}
    groups = {}
    for record in related.values():
        fields = record.get("fields") or {}
        if _text(fields.get(META_SOURCE)):
            continue
        item = _record_input(record)
        groups.setdefault(_item_key(item), []).append(item)
    collapsed = []
    consumed = set()
    for source in selected:
        if source["record_id"] in consumed:
            continue
        siblings = groups.get(_item_key(source), [])
        if len(siblings) > 1:
            source["sourceMembers"] = [{key: item.get(key, "") for key in
                                        ("record_id", "报关单号", "合同号_1", "报关品名",
                                         "供应商简称", "报关金额", "报关重量", "采购金额")}
                                       for item in siblings]
            source["mergedRecordIds"] = [item["record_id"] for item in siblings
                                         if item["record_id"] != source["record_id"]]
            consumed.update(source["mergedRecordIds"])
            if all(not _input_reason(item) for item in siblings):
                source["报关金额"] = f"{sum((_amount(item['报关金额']) for item in siblings), Decimal(0)):.2f}"
                source["报关重量"] = f"{sum((feishu_compare.number(item['报关重量']) for item in siblings), Decimal(0))}"
                if any(item["采购金额"] for item in siblings):
                    source["writebackBlocked"] = "同商品历史行已有采购金额，请复核后处理"
            else:
                source["writebackBlocked"] = "同商品历史行字段不完整，暂无法安全合并"
        collapsed.append(source)
    selected = collapsed

    context = []
    # 同合同补齐：只要本批选中的行涉及某张出运单，就把它在该出运单上的其它商品行一并拉进来。
    # 早先只在"填了日期范围"时才补，导致只按报关单号/合同号筛选时拆单池不完整、成本算不出来。
    if any(_tokens(item["合同号_1"]) for item in selected):
        matched = _search_many(client, app, table, "合同号_1",
                               {token for item in selected for token in _tokens(item["合同号_1"])},
                               operator="contains")
        selected_ids = {record_id for item in selected
                        for record_id in [item["record_id"], *item.get("mergedRecordIds", [])]}
        contracts_needed = {token for item in selected for token in _tokens(item["合同号_1"])}
        for record in matched:
            if record["record_id"] in selected_ids:
                continue
            fields = record.get("fields") or {}
            if not (_tokens(_text(fields.get("合同号_1"))) & contracts_needed):
                continue
            if _text(fields.get(META_SOURCE)):
                continue
            item = _record_input(record)
            reason = _input_reason(item)
            if not reason:
                item["contextOnly"] = True
                item["writebackBlocked"] = "同合同补齐记录仅用于计算，不属于所选未核算区间"
                context.append(item)
            else:
                skipped.append({"record_id": record["record_id"],
                                "报关单号": _text(fields.get("报关单号")),
                                "合同号_1": _text(fields.get("合同号_1")),
                                "tradeType": "内销" if not item["报关单号"] else "外销",
                                "reason": f"同合同记录无法参与补齐：{reason}"})
        context_groups = {}
        for item in context:
            context_groups.setdefault(_item_key(item), []).append(item)
        compact_context = []
        for group in context_groups.values():
            if len(group) == 1:
                compact_context.extend(group)
                continue
            if all(not _input_reason(item) for item in group):
                merged = group[0]
                merged["报关金额"] = f"{sum((_amount(item['报关金额']) for item in group), Decimal(0)):.2f}"
                merged["报关重量"] = f"{sum((feishu_compare.number(item['报关重量']) for item in group), Decimal(0))}"
                merged["mergedRecordIds"] = [item["record_id"] for item in group if item is not merged]
                compact_context.append(merged)
            else:
                compact_context.extend(group)
                for source in selected:
                    if _tokens(source["合同号_1"]) & _tokens(group[0]["合同号_1"]):
                        source["writebackBlocked"] = "同合同补齐记录存在无法安全合并的重复商品行，请人工核对"
        context = compact_context
    return {"selected": selected, "context": context, "skipped": skipped,
            "scanned": scanned + len(related) + len(context),
            "selection": {"total": len(selected), "declarations": len({r["报关单号"] for r in selected}),
                          "supplemental": len(context),
                          "merged": sum(bool(item.get("mergedRecordIds")) for item in selected),
                          "skipped": len(skipped), "scanned": scanned + len(related) + len(context)}}


def start_scan(**filters):
    chosen = select_records(**filters)
    if not chosen["selected"]:
        return {"id": None, "selection": chosen["selection"], "message": "所选范围内没有可处理的未核算商品行。"}
    sources = chosen["selected"] + chosen["context"]
    declarations = [{"来源文件": "飞书副本商品行", **source, "总价": source["报关金额"]}
                    for source in sources]
    job_id = workspace.start_job(kind="sync", declarations=declarations, source_records=sources)
    workspace._update(job_id, skipped=chosen["skipped"])
    return {"id": job_id, "selection": chosen["selection"], "skipped": chosen["skipped"]}


def _amount(value):
    raw = _text(value).replace(",", "")
    if not raw:
        return None
    try:
        return Decimal(raw).quantize(Decimal("0.01"))
    except InvalidOperation:
        return None


def _source_for(row, sources):
    candidates = [source for source in sources
                  if source["报关单号"] == row.get("报关单号")
                  and source.get("合同号_1", "") == row.get("合同号_1", "")]
    if len(candidates) == 1:
        return candidates[0], ""
    exact = [source for source in candidates if source["报关品名"] == row.get("报关品名")]
    supplier = _text(row.get("供应商简称"))
    if supplier:
        identified = [source for source in exact if supplier in
                      {item.get("供应商简称", "") for item in source.get("sourceMembers", [source])}]
        if len(identified) == 1:
            return identified[0], ""
    if len(exact) == 1:
        return exact[0], ""
    return None, "同单商品无法唯一对应，已暂停写回" if candidates else "飞书没有对应来源行"


def _writeback_key(row):
    return tuple(_text(row.get(name)) for name in
                 ("报关单号", "合同号_1", "供应商简称", "报关品名"))


def _row_key(source_id, row, ordinal):
    raw = "|".join((source_id, str(row.get("采购订单号") or ""),
                    str(row.get("供应商简称") or ""), str(row.get("产品类型") or ""), str(ordinal)))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def plan(job_id):
    rows = workspace.result_rows(job_id)
    source_path = workspace.JOBS_DIR / job_id / "source_records.json"
    if rows is None or not source_path.exists():
        return None
    sources = json.loads(source_path.read_text(encoding="utf-8"))
    groups = {source["record_id"]: {"source": source, "rows": [], "reasons": []} for source in sources}
    unmatched = []
    for index, row in enumerate(rows):
        source, reason = _source_for(row, sources)
        if source is None:
            unmatched.append({"index": index, "报关单号": row.get("报关单号"),
                              "合同号_1": row.get("合同号_1"), "reason": reason})
        else:
            groups[source["record_id"]]["rows"].append({"index": index, **row})
    output = []
    for source_id, group in groups.items():
        source, children, reasons = group["source"], group["rows"], group["reasons"]
        original = _amount(source.get("报关金额"))
        if source.get("writebackBlocked"):
            reasons.append(source["writebackBlocked"])
        if any(item["报关单号"] and item["报关单号"] == source["报关单号"]
               and item.get("合同号_1") == source.get("合同号_1") for item in unmatched):
            reasons.append("同单存在无法唯一对应的结果，已暂停写回")
        if not children:
            reasons.append("没有得到对应拆单结果")
        if original is None:
            reasons.append("来源行缺少拆分前报关金额")
        if any(child.get("异常类型") not in ("正常", "人工调整待写回") for child in children):
            reasons.append("组内存在待复核记录")
        amounts = [_amount(child.get("报关金额")) for child in children]
        if any(value is None for value in amounts):
            reasons.append("拆单结果存在空报关金额")
        total = sum((value or Decimal(0) for value in amounts), Decimal(0))
        if original is not None and total != original:
            reasons.append(f"报关金额不守恒：原行 {original:.2f}，拆分后 {total:.2f}")
        keys = [_writeback_key(child) for child in children]
        if any(not all(key) for key in keys):
            reasons.append("拆单结果缺少报关单号、合同号、供应商或品名，无法唯一写回")
        if len(keys) != len(set(keys)):
            reasons.append("报关单号、合同号、供应商和品名组合重复，无法唯一写回")
        members = source.get("sourceMembers") or []
        if len(members) > 1:
            member_keys = [_writeback_key(member) for member in members]
            if (any(not all(key) for key in member_keys) or len(set(member_keys)) != len(member_keys)
                    or not set(member_keys).issubset(set(keys))):
                reasons.append("历史拆分行无法全部按报关单号、合同号、供应商和品名对应，已暂停写回")
        numbered = [{**child, "splitKey": _row_key(source_id, child, ordinal)}
                    for ordinal, child in enumerate(children)]
        output.append({"sourceId": source_id, "declaration": source["报关单号"],
                       "contract": source.get("合同号_1", ""),
                       "sourceProduct": source.get("报关品名", ""),
                       "memberCount": len(source.get("sourceMembers", [source])),
                       "contextOnly": bool(source.get("contextOnly")),
                       "pdfTokens": source.get("pdfTokens", []),
                       "originalAmount": str(original) if original is not None else "",
                       "resultAmount": f"{total:.2f}", "status": "ready" if not reasons else "blocked",
                       "reasons": reasons, "children": numbered})
    return {"summary": {"ready": sum(g["status"] == "ready" for g in output),
                         "blocked": sum(g["status"] != "ready" for g in output),
                         "unmatched": len(unmatched)}, "groups": output, "unmatched": unmatched}


def set_mapping(job_id: str, row_index: int, source_id: str, operator: str) -> dict:
    raise ValueError("来源商品行已改为自动匹配；请重新打开写回预览")


def _fields(row, available, source_id, original):
    values = {"报关单号": row.get("报关单号"), "合同号_1": row.get("合同号_1"),
              "报关品名": row.get("报关品名"), "合同号（应收表格）": row.get("合同号（应收表格）"),
              "供应商简称": row.get("供应商简称"), "产品类型": row.get("产品类型"),
              "报关金额": _amount(row.get("报关金额")), "采购金额": _amount(row.get("采购金额")),
              "报关重量": _weight(row.get("报关重量")), META_SOURCE: source_id,
              META_KEY: row["splitKey"], META_ORIGINAL: _amount(original)}
    return {key: float(value) if isinstance(value, Decimal) else value for key, value in values.items()
            if key in available and value not in (None, "")}


def _weight(value):
    raw = _text(value).replace(",", "")
    try:
        return Decimal(raw) if raw else None
    except InvalidOperation:
        return None


def _write_targets(group, source, live):
    """在写入前用联合主键确定旧行；有歧义时整组停止。"""
    source_id = group["sourceId"]
    members = source.get("sourceMembers") or [{"record_id": source_id,
                                                "报关金额": group["originalAmount"]}]
    member_ids = {item["record_id"] for item in members}
    for member in members:
        record = live.get(member["record_id"])
        if not record:
            raise ValueError("飞书原商品行已被删除，请重新读取")
        actual = _amount((record.get("fields") or {}).get("报关金额"))
        if actual != _amount(member.get("报关金额")):
            raise ValueError("飞书原商品行金额已变化，请重新读取")
    eligible = member_ids | {record_id for record_id, record in live.items()
                             if _text((record.get("fields") or {}).get(META_SOURCE)) == source_id}
    targets, used = [], set()
    for index, child in enumerate(group["children"]):
        key = _writeback_key(child)
        if not all(key):
            raise ValueError("结果缺少报关单号、合同号、供应商或品名，无法唯一写回")
        outsiders = [record_id for record_id, record in live.items()
                    if record_id not in eligible and _writeback_key(record.get("fields") or {}) == key]
        if outsiders:
            raise ValueError("飞书出现新的同业务键商品行，请重新读取")
        matches = [record_id for record_id in eligible - used
                   if record_id in live and _writeback_key(live[record_id].get("fields") or {}) == key]
        if len(matches) > 1:
            raise ValueError("同业务键对应多条飞书原行，已暂停写回")
        target = matches[0] if matches else None
        if target is None and index == 0 and source_id not in used:
            target = source_id
        if target:
            used.add(target)
        targets.append(target)
    if member_ids - used:
        raise ValueError("仍有旧拆分行未被结果覆盖，已暂停写回")
    return targets


def push(job_id, operator=""):
    if not operator.strip():
        raise ValueError("写回前请填写操作员称呼")
    if not _PUSH_LOCK.acquire(blocking=False):
        raise ValueError("另一批写回正在进行，请稍后重试")
    try:
        proposal = plan(job_id)
        if proposal is None:
            raise ValueError("没有可写回的飞书任务")
        if not proposal["summary"]["ready"]:
            return {"results": [], "written": 0, "failed": 0, "message": "没有通过整组校验的记录"}
        client, app, table = _target()
        fields = {item["field_name"]: item for item in client.list_fields(app, table)}
        core_fields = {"报关单号", "合同号_1", "报关品名", "报关金额", "采购金额", "报关重量"}
        # 可直写字段类型：1=文本、2=数字、3=单选。
        # 「迈拓财务部门数据 副本」里 报关单号 / 报关品名 / 供应商简称 是**单选**字段
        # （旧 AI 副本是文本），单选值按选项名写入，飞书会自动补齐新选项。
        writable_types = (1, 2, 3)
        unavailable = {name for name in core_fields if fields.get(name, {}).get("type") not in writable_types}
        if unavailable:
            raise ValueError("目标表字段不可直接写入：" + "、".join(sorted(unavailable)))
        for name, kind in ((META_SOURCE, 1), (META_KEY, 1), (META_ORIGINAL, 2)):
            if name not in fields:
                client.create_field(app, table, name, kind)
        fields = {item["field_name"]: item for item in client.list_fields(app, table)}
        available = {name for name, item in fields.items() if item.get("type") in writable_types}
        required = {"报关单号", "合同号_1", "报关品名", "报关金额", "采购金额", "报关重量",
                    META_SOURCE, META_KEY, META_ORIGINAL}
        missing = required - available
        if missing:
            raise ValueError("目标表字段不可直接写入：" + "、".join(sorted(missing)))
        live = {record["record_id"]: record for record in client.iter_records(app, table)}
        source_path = workspace.JOBS_DIR / job_id / "source_records.json"
        saved_sources = json.loads(source_path.read_text(encoding="utf-8")) if source_path.exists() else []
        sources_by_id = {item["record_id"]: item for item in saved_sources}
        results = []
        for group in proposal["groups"]:
            source_id = group["sourceId"]
            if group["status"] != "ready":
                results.append({"sourceId": source_id, "status": "skipped", "reason": "；".join(group["reasons"])})
                continue
            if source_id not in live:
                results.append({"sourceId": source_id, "status": "failed", "reason": "来源记录已被删除"})
                continue
            try:
                children = group["children"]
                targets = _write_targets(group, sources_by_id.get(source_id, {}), live)
                created_count = 0
                for child, target_id in zip(children, targets):
                    if target_id:
                        continue
                    payload = _fields(child, available, source_id, group["originalAmount"])
                    created = client.create_record(app, table, payload)
                    target_id = created.get("record", created).get("record_id", "")
                    created_count += 1
                for child, target_id in zip(children, targets):
                    if not target_id:
                        continue
                    payload = _fields(child, available, "" if target_id == source_id else source_id,
                                      group["originalAmount"])
                    if target_id == source_id:
                        payload.pop(META_SOURCE, None)
                    client.update_record(app, table, target_id, payload)
                results.append({"sourceId": source_id, "status": "written", "created": created_count})
            except Exception as exc:  # noqa: BLE001
                results.append({"sourceId": source_id, "status": "failed", "reason": str(exc)[:200]})
        workspace.append_audit(job_id, "feishu_push", operator, {"results": results})
        (workspace.JOBS_DIR / job_id / "push_results.json").write_text(
            json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
        return {"results": results, "written": sum(r["status"] == "written" for r in results),
                "failed": sum(r["status"] == "failed" for r in results)}
    finally:
        _PUSH_LOCK.release()
