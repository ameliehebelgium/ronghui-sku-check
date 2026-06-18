"""
比对引擎：将本次批量解析出的SKU记录与数据库现有记录比对，
得到 NEW / MATCH / CONFLICT 三种状态。

规则（已与用户确认）：
- 品名比对：要求文字完全一致才算一致，不做模糊匹配
- HS比对：要求数值完全一致（字符串比较，已在解析阶段去除 .0 等后缀）
- 数据库中不存在该SKU -> NEW（自动写入，不需要人工确认）
- 数据库中存在，品名和HS都完全一致 -> MATCH（自动忽略）
- 数据库中存在，但品名或HS有差异 -> CONFLICT（需人工集中确认，保留原记录或更新）
"""
import pandas as pd


def compare_batch(new_df: pd.DataFrame, db_df: pd.DataFrame) -> pd.DataFrame:
    """
    new_df: 本次批量解析结果，列 [sku, description, hs_code, po_number, source_file]
            可能包含同一SKU出现在多个文件里的情况（同一批次内部重复）
    db_df:  数据库现有记录，列至少包含 [sku, description, hs_code]

    返回 DataFrame，每个唯一SKU一行，新增列：
        status: NEW / MATCH / CONFLICT
        db_description, db_hs_code: 数据库里原有的值（NEW时为空）
        sources: 本批次中出现该SKU的所有来源文件名（逗号分隔）
    """
    db_lookup = {}
    if db_df is not None and len(db_df) > 0:
        for _, row in db_df.iterrows():
            db_lookup[str(row["sku"]).strip()] = {
                "description": str(row.get("description", "")).strip(),
                "hs_code": str(row.get("hs_code", "")).strip(),
            }

    # 同一批次内，同一SKU可能出现在多个文件里：按SKU分组，记录所有来源文件
    grouped = (
        new_df.groupby("sku")
        .agg({
            "description": "first",
            "hs_code": "first",
            "po_number": lambda x: ", ".join(sorted(set(x))),
            "source_file": lambda x: ", ".join(sorted(set(x))),
        })
        .reset_index()
        .rename(columns={"po_number": "po_numbers", "source_file": "sources"})
    )

    rows = []
    for _, r in grouped.iterrows():
        sku = r["sku"]
        new_desc = str(r["description"]).strip()
        new_hs = str(r["hs_code"]).strip()
        db_rec = db_lookup.get(sku)

        if db_rec is None:
            status = "NEW"
            db_desc, db_hs = "", ""
        else:
            db_desc, db_hs = db_rec["description"], db_rec["hs_code"]
            if new_desc == db_desc and new_hs == db_hs:
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
            "po_numbers": r["po_numbers"],
            "sources": r["sources"],
        })

    return pd.DataFrame(rows)


def summarize(result_df: pd.DataFrame, file_count: int) -> dict:
    """生成当次录入报告所需的统计数字"""
    total_skus = len(result_df)
    new_count = (result_df["status"] == "NEW").sum()
    match_count = (result_df["status"] == "MATCH").sum()
    conflict_count = (result_df["status"] == "CONFLICT").sum()
    return {
        "file_count": file_count,
        "total_skus": int(total_skus),
        "new_count": int(new_count),
        "match_count": int(match_count),
        "conflict_count": int(conflict_count),
    }
