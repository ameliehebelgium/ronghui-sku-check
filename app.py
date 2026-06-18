"""
RongHui_SKU_HS_Check
独立于 EU Risk App / UK SKU Tracker 的系统，专门处理 RongHui（HONGKONG SICHANG
INTERNATIONAL GROUP）提供的 packing list，做 SKU 品名/HS 的一致性校验与主数据库维护。
"""
import io
import zipfile
from datetime import datetime, date

import pandas as pd
import streamlit as st

from pl_parser import parse_packing_list
from compare_engine import compare_batch, summarize
import sheets_db

st.set_page_config(page_title="RongHui_SKU_HS_Check", layout="wide")

APP_USERNAME = "Admin"
APP_PASSWORD = "Admin123"


def login_gate():
    if st.session_state.get("authenticated"):
        return True
    st.title("RongHui_SKU_HS_Check")
    with st.form("login_form"):
        u = st.text_input("用户名")
        p = st.text_input("密码", type="password")
        submitted = st.form_submit_button("登录")
    if submitted:
        if u == APP_USERNAME and p == APP_PASSWORD:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("用户名或密码错误")
    return False


def extract_excels_from_zip(zip_file) -> list[tuple[str, io.BytesIO]]:
    out = []
    with zipfile.ZipFile(zip_file) as zf:
        for name in zf.namelist():
            if name.lower().endswith((".xlsx", ".xls")) and not name.startswith("__MACOSX"):
                data = zf.read(name)
                out.append((name.split("/")[-1], io.BytesIO(data)))
    return out


def page_upload_compare():
    st.header("上传与比对")
    upload_mode = st.radio("上传方式", ["ZIP批量上传", "单个/多个Excel文件"], horizontal=True)

    files_to_parse = []  # list of (filename, buffer)

    if upload_mode == "ZIP批量上传":
        zip_file = st.file_uploader("上传ZIP（包含多个packing list）", type=["zip"])
        if zip_file:
            files_to_parse = extract_excels_from_zip(zip_file)
            st.info(f"ZIP中识别到 {len(files_to_parse)} 个Excel文件")
    else:
        uploaded = st.file_uploader(
            "上传一个或多个packing list", type=["xlsx", "xls"], accept_multiple_files=True
        )
        if uploaded:
            files_to_parse = [(f.name, f) for f in uploaded]

    if not files_to_parse:
        return

    if st.button("开始解析与比对", type="primary"):
        all_records = []
        parse_errors = []
        progress = st.progress(0.0)
        for i, (fname, buf) in enumerate(files_to_parse):
            try:
                df = parse_packing_list(buf, fname)
                all_records.append(df)
            except Exception as e:
                parse_errors.append((fname, str(e)))
            progress.progress((i + 1) / len(files_to_parse))

        if parse_errors:
            with st.expander(f"⚠️ {len(parse_errors)} 个文件解析失败", expanded=True):
                for fname, err in parse_errors:
                    st.error(f"{fname}: {err}")

        if not all_records:
            st.warning("没有成功解析的文件")
            return

        new_df = pd.concat(all_records, ignore_index=True)
        db_df = sheets_db.load_master_db()
        result_df = compare_batch(new_df, db_df)

        st.session_state["last_result_df"] = result_df
        st.session_state["last_file_count"] = len(files_to_parse) - len(parse_errors)
        st.session_state["pending_conflicts"] = result_df[result_df["status"] == "CONFLICT"].copy()
        st.session_state["pending_new"] = result_df[result_df["status"] == "NEW"].copy()

    if "last_result_df" not in st.session_state:
        return

    result_df = st.session_state["last_result_df"]
    stats = summarize(result_df, st.session_state["last_file_count"])

    st.subheader("本次录入报告")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("录入文件数", stats["file_count"])
    c2.metric("合计SKU数", stats["total_skus"])
    c3.metric("已存在且一致", stats["match_count"])
    c4.metric("新SKU", stats["new_count"])
    c5.metric("需确认冲突", stats["conflict_count"])

    pending_new = st.session_state.get("pending_new", pd.DataFrame())
    if len(pending_new) > 0 and not st.session_state.get("new_skus_written", False):
        new_rows = [
            {
                "sku": r["sku"], "description": r["new_description"],
                "hs_code": r["new_hs_code"], "sources": r["sources"],
            }
            for _, r in pending_new.iterrows()
        ]
        sheets_db.append_new_skus(new_rows)
        st.session_state["new_skus_written"] = True
        st.session_state["today_new_skus"] = pending_new
        st.success(f"已自动写入 {len(new_rows)} 个新SKU到数据库")

    pending_conflicts = st.session_state.get("pending_conflicts", pd.DataFrame())
    if len(pending_conflicts) > 0:
        st.subheader(f"需人工确认的冲突（{len(pending_conflicts)} 个）")
        st.caption("勾选「更新为新记录」的SKU将用本次packing list的数据覆盖数据库；未勾选的默认保留数据库原记录")

        display_df = pending_conflicts.copy()
        display_df.insert(0, "更新为新记录", False)
        edited = st.data_editor(
            display_df[["更新为新记录", "sku", "db_description", "db_hs_code", "new_description", "new_hs_code", "sources"]],
            column_config={
                "更新为新记录": st.column_config.CheckboxColumn(),
                "sku": "SKU",
                "db_description": "数据库原品名",
                "db_hs_code": "数据库原HS",
                "new_description": "本次品名",
                "new_hs_code": "本次HS",
                "sources": "来源文件",
            },
            disabled=["sku", "db_description", "db_hs_code", "new_description", "new_hs_code", "sources"],
            hide_index=True,
            use_container_width=True,
            key="conflict_editor",
        )

        if st.button("提交冲突处理结果", type="primary"):
            decisions = []
            for _, row in edited.iterrows():
                orig = pending_conflicts[pending_conflicts["sku"] == row["sku"]].iloc[0]
                decisions.append({
                    "sku": row["sku"],
                    "action": "UPDATE" if row["更新为新记录"] else "KEEP",
                    "old_description": orig["db_description"],
                    "old_hs_code": orig["db_hs_code"],
                    "new_description": orig["new_description"],
                    "new_hs_code": orig["new_hs_code"],
                })
            sheets_db.apply_conflict_decisions(decisions)
            update_count = sum(1 for d in decisions if d["action"] == "UPDATE")
            keep_count = len(decisions) - update_count
            st.success(f"已处理 {len(decisions)} 个冲突：{update_count} 个更新，{keep_count} 个保留原记录")
            st.session_state.pop("pending_conflicts", None)
    elif "last_result_df" in st.session_state:
        st.info("本次没有需要人工确认的冲突")

    today_new = st.session_state.get("today_new_skus")
    if today_new is not None and len(today_new) > 0:
        st.subheader("当天新SKU总结表")
        export_df = today_new[["sku", "new_description", "new_hs_code", "sources"]].rename(
            columns={"new_description": "description", "new_hs_code": "hs_code"}
        )
        st.dataframe(export_df, use_container_width=True, hide_index=True)
        buf = io.BytesIO()
        export_df.to_excel(buf, index=False)
        st.download_button(
            "下载当天新SKU总结表 (Excel)",
            data=buf.getvalue(),
            file_name=f"RongHui_New_SKUs_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def page_database():
    st.header("主数据库")
    db_df = sheets_db.load_master_db()
    st.caption(f"共 {len(db_df)} 条SKU记录")

    search = st.text_input("搜索SKU或品名")
    display_df = db_df
    if search:
        mask = db_df["sku"].str.contains(search, case=False, na=False) | db_df["description"].str.contains(search, case=False, na=False)
        display_df = db_df[mask]

    st.dataframe(display_df, use_container_width=True, hide_index=True)

    buf = io.BytesIO()
    db_df.to_excel(buf, index=False)
    st.download_button(
        "一键下载主数据库 (Excel)",
        data=buf.getvalue(),
        file_name=f"RongHui_SKU_Master_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def page_change_log():
    st.header("变更日志")
    log_df = sheets_db.load_change_log()
    st.caption(f"共 {len(log_df)} 条记录")
    action_filter = st.multiselect("筛选操作类型", options=sorted(log_df["action"].unique()) if len(log_df) else [])
    display_df = log_df
    if action_filter:
        display_df = log_df[log_df["action"].isin(action_filter)]
    st.dataframe(display_df.sort_values("timestamp", ascending=False) if len(display_df) else display_df,
                 use_container_width=True, hide_index=True)


def main():
    if not login_gate():
        return

    st.sidebar.title("RongHui_SKU_HS_Check")
    page = st.sidebar.radio("导航", ["上传与比对", "主数据库", "变更日志"])
    if st.sidebar.button("退出登录"):
        st.session_state["authenticated"] = False
        st.rerun()

    if page == "上传与比对":
        page_upload_compare()
    elif page == "主数据库":
        page_database()
    elif page == "变更日志":
        page_change_log()


if __name__ == "__main__":
    main()
