"""
比对引擎：将本次批量解析出的SKU记录与数据库现有记录比对，
得到 NEW / MATCH / CONFLICT 三种状态；同时检测同一批次内部
（同一SKU出现在多个文件里）品名/HS是否一致，不一致的单独列为
BATCH_CONFLICT，需人工选择采用哪个文件的值。

规则（已与用户确认）：
- 品名比对：忽略大小写后完全一致即视为一致（case-insensitive）
- HS比对：要求数值完全一致（字符串比较，已在解析阶段去除 .0 等后缀）
- 数据库中不存在该SKU -> NEW（人工确认汇总报告后才写入，不自动写入）
- 数据库中存在，品名和HS都完全一致 -> MATCH（自动忽略）
- 数据库中存在，但品名或HS有差异 -> CONFLICT（需人工集中确认，保留原记录或更新）
- 同一批次内部，同一SKU在不同文件里品名/HS不一致 -> BATCH_CONFLICT
  （需人工先选定本批次最终采用哪个文件的值，再继续跟数据库比对）
"""
import pandas as pd


def _variants_for_sku(group: pd.DataFrame) -> list[dict]:
    """同一SKU在本批次内出现的所有不同(description, hs_code)组合及对应来源文件
    品名比对忽略大小写；key 用 lower() 去重，展示值保留第一次出现的原始大小写。
    """
    seen = {}   # lower_key -> (original_desc, original_hs, files)
    for _, row in group.iterrows():
        desc = str(row["description"]).strip()
        hs   = str(row["hs_code"]).strip()
        key  = (desc.lower(), hs.lower())
        if key not in seen:
            seen[key] = (desc, hs, set())
        seen[key][2].add(row["source_file"])
    return [
        {"description": orig_desc, "hs_code": orig_hs, "files": sorted(files)}
        for (orig_desc, orig_hs, files) in seen.values()
    ]


def compare_batch(new_df: pd.DataFrame, db_df: pd.DataFrame) -> pd.DataFrame:
    """
    new_df: 本次批量解析结果，列 [sku, description, hs_code, po_number, source_file]
            可能包含同一SKU出现在多个文件里的情况（同一批次内部重复）
    db_df:  数据库现有记录，列至少包含 [sku, description, hs_code]

    返回 DataFrame，每个唯一SKU一行，新增列：
        status: NEW / MATCH / CONFLICT / BATCH_CONFLICT
        db_description, db_hs_code: 数据库里原有的值（NEW/BATCH_CONFLICT时可能为空）
        sources: 本批次中出现该SKU的所有来源文件名（逗号分隔）
        variants: 仅 BATCH_CONFLICT 时有效，本批次内出现的所有不同(品名,HS)组合
    """
    db_lookup = {}
    if db_df is not None and len(db_df) > 0:
        for _, row in db_df.iterrows():
            db_lookup[str(row["sku"]).strip()] = {
                "description": str(row.get("description", "")).strip(),
                "hs_code": str(row.get("hs_code", "")).strip(),
            }

    rows = []
    for sku, group in new_df.groupby("sku"):
        variants = _variants_for_sku(group)
        po_numbers = ", ".join(sorted(set(group["po_number"])))
        sources = ", ".join(sorted(set(group["source_file"])))

        if len(variants) > 1:
            # 本批次内部就不一致，先暴露出来让人工选择，暂不跟数据库比较
            rows.append({
                "sku": sku,
                "new_description": variants[0]["description"],
                "new_hs_code": variants[0]["hs_code"],
                "db_description": "",
                "db_hs_code": "",
                "status": "BATCH_CONFLICT",
                "po_numbers": po_numbers,
                "sources": sources,
                "variants": variants,
            })
            continue

        new_desc = variants[0]["description"]
        new_hs = variants[0]["hs_code"]
        db_rec = db_lookup.get(sku)

        if db_rec is None:
            status = "NEW"
            db_desc, db_hs = "", ""
        else:
            db_desc, db_hs = db_rec["description"], db_rec["hs_code"]
            if new_desc.lower() == db_desc.lower() and new_hs.lower() == db_hs.lower():
                status = "MATCH"
            else:
                status = "CONFLICT"

        rows.append({
            "sku": sku,
            "new_description": new_desc,
            "new_hs_code": new_hs,
            "db_description": db_desc,
            "db_hs_code": db_hs,
            "status": status,
            "po_numbers": po_numbers,
            "sources": sources,
            "variants": variants,
        })

    return pd.DataFrame(rows)


def resolve_batch_conflicts(result_df: pd.DataFrame, resolutions: dict, db_df: pd.DataFrame) -> pd.DataFrame:
    """
    resolutions: {sku: chosen_index}，chosen_index 对应 variants 列表里用户选定的那一项
    把 BATCH_CONFLICT 的行替换成用户选定的值，并重新跟数据库比对得到最终 NEW/MATCH/CONFLICT
    """
    db_lookup = {}
    if db_df is not None and len(db_df) > 0:
        for _, row in db_df.iterrows():
            db_lookup[str(row["sku"]).strip()] = {
                "description": str(row.get("description", "")).strip(),
                "hs_code": str(row.get("hs_code", "")).strip(),
            }

    result_df = result_df.copy()
    for idx, row in result_df.iterrows():
        if row["status"] != "BATCH_CONFLICT":
            continue
        sku = row["sku"]
        chosen_idx = resolutions.get(sku, 0)
        chosen = row["variants"][chosen_idx]
        new_desc, new_hs = chosen["description"], chosen["hs_code"]

        db_rec = db_lookup.get(sku)
        if db_rec is None:
            status = "NEW"
            db_desc, db_hs = "", ""
        else:
            db_desc, db_hs = db_rec["description"], db_rec["hs_code"]
            status = "MATCH" if (new_desc.lower() == db_desc.lower() and new_hs.lower() == db_hs.lower()) else "CONFLICT"

        result_df.at[idx, "new_description"] = new_desc
        result_df.at[idx, "new_hs_code"] = new_hs
        result_df.at[idx, "db_description"] = db_desc
        result_df.at[idx, "db_hs_code"] = db_hs
        result_df.at[idx, "status"] = status

    return result_df


def summarize(result_df: pd.DataFrame, file_count: int) -> dict:
    """生成当次录入报告所需的统计数字"""
    total_skus = len(result_df)
    new_count = (result_df["status"] == "NEW").sum()
    match_count = (result_df["status"] == "MATCH").sum()
    conflict_count = (result_df["status"] == "CONFLICT").sum()
    batch_conflict_count = (result_df["status"] == "BATCH_CONFLICT").sum()
    return {
        "file_count": file_count,
        "total_skus": int(total_skus),
        "new_count": int(new_count),
        "match_count": int(match_count),
        "conflict_count": int(conflict_count),
        "batch_conflict_count": int(batch_conflict_count),
    }
