"""
RongHui_SKU_HS_Check
独立于 EU Risk App / UK SKU Tracker 的系统，专门处理 RongHui（HONGKONG SICHANG
INTERNATIONAL GROUP）提供的 packing list，做 SKU 品名/HS 的一致性校验与主数据库维护。
"""
import io
import zipfile
from datetime import date

import pandas as pd
import streamlit as st

from pl_parser import parse_packing_list
from compare_engine import compare_batch, resolve_batch_conflicts, summarize
import sheets_db
from logo_const import VEVOR_LOGO_B64

st.set_page_config(page_title="RongHui_SKU_HS_Check", layout="wide")

APP_USERNAME = "Admin"
APP_PASSWORD = "Admin123"

CUSTOM_CSS = """
<style>
:root {
    --navy-deep: #0B1220;
    --navy-mid: #14233D;
    --brand-blue: #1E3A5F;
    --surface: #F7F8FA;
    --border-soft: #D8DEE4;
    --text-primary: #15202B;
    --text-muted: #5B6B7C;
}

/* 侧边栏：深墨蓝，企业系统感 */
section[data-testid="stSidebar"] {
    background-color: var(--navy-deep);
    border-right: 1px solid var(--navy-mid);
}
section[data-testid="stSidebar"] * {
    color: #E7ECF2 !important;
}
section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] .stRadio label {
    color: #C7D2DD !important;
}

/* 主内容区背景 */
.main .block-container {
    background-color: var(--surface);
    padding-top: 2rem;
}

/* 标题样式 */
h1, h2, h3 {
    color: var(--brand-blue);
    font-weight: 700;
    letter-spacing: -0.01em;
}

/* 指标卡片化 */
div[data-testid="stMetric"] {
    background-color: white;
    border: 1px solid var(--border-soft);
    border-radius: 6px;
    padding: 0.9rem 1rem;
}
div[data-testid="stMetric"] label {
    color: var(--text-muted) !important;
    font-weight: 600;
}

/* 信息/成功提示条，去掉默认圆润感，改为左侧色条风格 */
div[data-testid="stAlert"] {
    border-radius: 4px;
    border-left: 4px solid var(--brand-blue);
}

/* 表格容器边框 */
div[data-testid="stDataFrame"] {
    border: 1px solid var(--border-soft);
    border-radius: 6px;
}

/* 页脚签名 */
.app-footer {
    margin-top: 3rem;
    padding-top: 1rem;
    border-top: 1px solid var(--border-soft);
    text-align: center;
    color: var(--text-muted);
    font-size: 0.78rem;
    letter-spacing: 0.02em;
}
</style>
"""


def inject_global_style():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def render_sidebar_header():
    st.sidebar.markdown(
        f"""
        <div style="margin-bottom:1.2rem;">
            <img src="data:image/png;base64,{VEVOR_LOGO_B64}"
                 style="width:100%; border-radius:6px; display:block;" />
        </div>
        <hr style="border-color:#22324A; margin:0.8rem 0 1.2rem 0;">
        """,
        unsafe_allow_html=True,
    )


def render_sidebar_footer():
    st.sidebar.markdown(
        """
        <hr style="border-color:#22324A; margin:1.5rem 0 0.8rem 0;">
        <div style="color:#7C8CA0; font-size:0.78rem;">
            Designed and created by<br>
            <span style="color:#C7D2DD; font-weight:600;">Amélie — Vevor EU</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_footer():
    st.markdown(
        '<div class="app-footer">Designed and created by Amélie — Vevor EU</div>',
        unsafe_allow_html=True,
    )


def render_page_header():
    st.markdown(
        """
        <div style="color:#5B6B7C; font-size:0.85rem; margin-top:-0.6rem; margin-bottom:1.2rem;">
            RongHui_SKU_HS_Check &nbsp;·&nbsp; Vevor EU — Internal Use Only
        </div>
        """,
        unsafe_allow_html=True,
    )


def login_gate():
    inject_global_style()
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
    render_footer()
    return False


def extract_excels_from_zip(zip_file) -> list[tuple[str, io.BytesIO]]:
    out = []
    with zipfile.ZipFile(zip_file) as zf:
        for name in zf.namelist():
            if name.lower().endswith((".xlsx", ".xls")) and not name.startswith("__MACOSX"):
                data = zf.read(name)
                out.append((name.split("/")[-1], io.BytesIO(data)))
    return out


def _reset_batch_state():
    for key in [
        "last_result_df", "last_file_count", "pending_conflicts", "pending_new",
        "new_skus_written", "today_new_skus", "batch_conflicts_resolved",
    ]:
        st.session_state.pop(key, None)


def page_upload_compare():
    st.header("上传与比对")
    render_page_header()
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
        _reset_batch_state()
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
        st.session_state["last_db_df"] = db_df

    if "last_result_df" not in st.session_state:
        return

    result_df = st.session_state["last_result_df"]

    # ---------- 第一步：批次内部冲突必须先处理 ----------
    batch_conflicts = result_df[result_df["status"] == "BATCH_CONFLICT"]
    if len(batch_conflicts) > 0 and not st.session_state.get("batch_conflicts_resolved", False):
        st.subheader(f"⚠️ 批次内部冲突（{len(batch_conflicts)} 个SKU）")
        st.caption(
            "以下SKU在本次上传的不同文件里，品名或HS不一致。请为每个SKU选择本次最终采用哪个版本"
            "（默认已预选「出现文件数最多」的版本，可在下拉框里修改）。处理完才能继续后续比对。"
        )

        choice_rows = []
        for _, row in batch_conflicts.iterrows():
            variants = row["variants"]
            # 默认预选出现次数最多的变体
            default_idx = max(range(len(variants)), key=lambda i: len(variants[i]["files"]))
            options = [
                f"[{i+1}] {v['description']} | HS={v['hs_code']} | 来自{len(v['files'])}个文件: {', '.join(v['files'])}"
                for i, v in enumerate(variants)
            ]
            choice_rows.append({
                "SKU": row["sku"],
                "选择采用哪个版本": options[default_idx],
                "_options": options,
                "_sku": row["sku"],
            })

        for i, item in enumerate(choice_rows):
            choice_rows[i]["选择采用哪个版本"] = st.selectbox(
                f"SKU: {item['SKU']}",
                options=item["_options"],
                index=item["_options"].index(item["选择采用哪个版本"]),
                key=f"batch_conflict_{item['_sku']}",
            )

        if st.button("✅ 确认批次内部冲突的选择，继续比对", type="primary"):
            resolutions = {}
            for item in choice_rows:
                chosen_label = item["选择采用哪个版本"]
                chosen_idx = item["_options"].index(chosen_label)
                resolutions[item["_sku"]] = chosen_idx

            resolved_df = resolve_batch_conflicts(result_df, resolutions, st.session_state["last_db_df"])
            st.session_state["last_result_df"] = resolved_df
            st.session_state["batch_conflicts_resolved"] = True
            st.session_state["pending_conflicts"] = resolved_df[resolved_df["status"] == "CONFLICT"].copy()
            st.session_state["pending_new"] = resolved_df[resolved_df["status"] == "NEW"].copy()
            st.rerun()
        return  # 批次内部冲突没处理完，不展示后续报告

    if len(batch_conflicts) == 0:
        st.session_state.setdefault("pending_conflicts", result_df[result_df["status"] == "CONFLICT"].copy())
        st.session_state.setdefault("pending_new", result_df[result_df["status"] == "NEW"].copy())
        st.session_state.setdefault("batch_conflicts_resolved", True)

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
        st.info(f"本次发现 {len(pending_new)} 个新SKU，预览如下，确认无误后点击下方按钮写入数据库")
        preview_search = st.text_input("搜索新SKU预览（按SKU或品名）", key="preview_search")
        preview_df = pending_new[["sku", "new_description", "new_hs_code", "sources"]].rename(
            columns={"new_description": "品名", "new_hs_code": "HS编码", "sku": "SKU", "sources": "来源文件"}
        )
        if preview_search:
            mask = (
                preview_df["SKU"].str.contains(preview_search, case=False, na=False)
                | preview_df["品名"].str.contains(preview_search, case=False, na=False)
            )
            preview_df = preview_df[mask]
        st.dataframe(preview_df, use_container_width=True, hide_index=True, height=350)

        if st.button("✅ 确认写入新SKU到数据库", type="primary"):
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
            st.success(f"已写入 {len(new_rows)} 个新SKU到数据库")
            st.rerun()

    pending_conflicts = st.session_state.get("pending_conflicts", pd.DataFrame())
    if len(pending_conflicts) > 0:
        st.subheader(f"需人工确认的数据库冲突（{len(pending_conflicts)} 个）")
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

        if st.button("提交数据库冲突处理结果", type="primary"):
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
    elif st.session_state.get("batch_conflicts_resolved", False):
        st.info("本次没有需要人工确认的数据库冲突")

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
    render_page_header()
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
    render_page_header()
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

    render_sidebar_header()
    page = st.sidebar.radio("导航", ["上传与比对", "主数据库", "变更日志"])
    if st.sidebar.button("退出登录"):
        st.session_state["authenticated"] = False
        st.rerun()
    render_sidebar_footer()

    if page == "上传与比对":
        page_upload_compare()
    elif page == "主数据库":
        page_database()
    elif page == "变更日志":
        page_change_log()


if __name__ == "__main__":
    main()
