"""
app.py
------
Streamlit Dashboard – Test Report Comparison Tool

Run locally (LAN accessible):
    streamlit run app.py --server.address 0.0.0.0 --server.port 8501

Colleagues on the same network open:
    http://<your-PC-IP>:8501
"""
import io
import zipfile
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import streamlit as st

from file_matcher import match_files
from agent import compare, load_prompt, chat_reply
from report_generator import generate_word_report
import token_tracker

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Test Report Comparison Tool",
    page_icon="🔍",
    layout="wide",
)

# ── Sidebar: Token usage dashboard ──────────────────────────────────────────────
with st.sidebar:
    st.header("📊 Token 用量看板")
    usage_df = token_tracker.load_usage_log()
    if usage_df.empty:
        st.caption("暂无调用记录 / No LLM calls logged yet.")
    else:
        c1, c2 = st.columns(2)
        c1.metric("调用次数 / Calls", len(usage_df))
        c2.metric("总 Token", f"{int(usage_df['total_tokens'].sum()):,}")
        c3, c4 = st.columns(2)
        c3.metric("Input Tokens", f"{int(usage_df['input_tokens'].sum()):,}")
        c4.metric("Output Tokens", f"{int(usage_df['output_tokens'].sum()):,}")
        with st.expander("查看明细 / Details"):
            st.dataframe(
                usage_df.sort_values("timestamp", ascending=False),
                use_container_width=True,
                hide_index=True,
            )
        if st.button("🗑 清空记录 / Clear log", use_container_width=True):
            token_tracker.clear_usage_log()
            st.rerun()

# ── Load comparison prompt once ────────────────────────────────────────────────
_prompt_candidates = [
    Path(__file__).parent / "test_report_prompt_v15.docx",
    Path(__file__).parent / "test_report_prompt_v12.docx",
    Path(__file__).parent / "EXP3_BO_prompt.docx",
    Path(__file__).parent / "test_report_prompt_v4.docx",
]
PROMPT_PATH = str(next((p for p in _prompt_candidates if p.exists()), _prompt_candidates[-1]))

@st.cache_resource
def get_prompt() -> str:
    return load_prompt(PROMPT_PATH)

PROMPT_TEXT = get_prompt()

# ── Session-state defaults ─────────────────────────────────────────────────────
if "results" not in st.session_state:
    st.session_state.results = {}

if "chat_history" not in st.session_state:
    st.session_state.chat_history = [
        {
            "role": "assistant",
            "content": (
                "你好！我是你的测试分析助手。请先上传文件，然后告诉我本次分析的重点方向。\n\n"
                "• 经验丰富的工程师：直接给出清晰指令，我会快速确认并准备就绪\n"
                "• 经验较少的工程师：我们可以多轮对话，我会逐步帮你理清分析方向"
            ),
        }
    ]

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 – Header
# ══════════════════════════════════════════════════════════════════════════════
st.title("🔍 Test Report Comparison Tool")
st.caption(
    "Upload Test Cases and Test Reports · Files are matched by filename · "
    "Results generated as Word documents"
)
st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 – File Upload
# ══════════════════════════════════════════════════════════════════════════════
col_tc, col_tr = st.columns(2)

ALLOWED_TYPES = ["xlsx", "xls", "docx", "pdf", "txt", "csv"]

with col_tc:
    st.subheader("📋 Test Cases")
    tc_files = st.file_uploader(
        "Upload the single Test Case file (contains all requirements)",
        accept_multiple_files=False,
        key="tc_upload",
        type=ALLOWED_TYPES,
    )
    tc_files = [tc_files] if tc_files else []
    if tc_files:
        for f in tc_files:
            st.caption(f"• {f.name}")

with col_tr:
    st.subheader("📄 Test Reports")
    tr_files = st.file_uploader(
        "Upload one or more Test Report files (each will be compared against the TC)",
        accept_multiple_files=True,
        key="tr_upload",
        type=ALLOWED_TYPES,
    )
    if tr_files:
        for f in tr_files:
            st.caption(f"• {f.name}")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 – File Pairing Preview
# ══════════════════════════════════════════════════════════════════════════════
matched: list = []

if tc_files and tr_files:
    matched, unmatched_tc, unmatched_tr = match_files(tc_files, tr_files)

    st.divider()
    st.subheader("🔗 File Pairing Preview")

    pair_col, warn_col = st.columns(2)

    with pair_col:
        if matched:
            st.write("**Matched pairs**")
            for pair in matched:
                st.success(f"✅  {pair.test_case.name}  ↔  {pair.test_report.name}")
        else:
            st.info("No pairs found. Please upload a Test Case file and at least one Test Report.")

    with warn_col:
        if unmatched_tc or unmatched_tr:
            st.write("**Unmatched files**")
            for name in unmatched_tc:
                st.warning(f"⚠️  TC without report : {name}")
            for name in unmatched_tr:
                st.warning(f"⚠️  Report without TC : {name}")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 – Pre-analysis Chat
# ══════════════════════════════════════════════════════════════════════════════
st.divider()
st.subheader("💬 与分析助手对话")
st.caption(
    "与助手沟通本次分析的重点方向；确认就绪后点击 **▶ Run Comparison** 开始分析。"
)

for msg in st.session_state.chat_history:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

if instruction := st.chat_input("输入消息…"):
    st.session_state.chat_history.append({"role": "user", "content": instruction})
    with st.spinner("助手思考中…"):
        reply = chat_reply(st.session_state.chat_history)
    st.session_state.chat_history.append({"role": "assistant", "content": reply})
    st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 – Run & Download Controls
# ══════════════════════════════════════════════════════════════════════════════
st.divider()
run_col, dl_col = st.columns(2)

with run_col:
    run_clicked = st.button(
        "▶  Run Comparison",
        type="primary",
        disabled=len(matched) == 0,
        use_container_width=True,
    )

with dl_col:
    if st.session_state.results:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in st.session_state.results.items():
                zf.writestr(f"{name}.docx", data["word_bytes"])
        st.download_button(
            label="📥  Download All Results (.zip)",
            data=zip_buf.getvalue(),
            file_name="test_results.zip",
            mime="application/zip",
            use_container_width=True,
            key="dl_all_zip",
        )
    else:
        st.button(
            "📥  Download All Results (.zip)",
            disabled=True,
            use_container_width=True,
        )

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 – Run Logic  (executes after button press)
# ══════════════════════════════════════════════════════════════════════════════
if run_clicked and matched:
    st.session_state.results = {}
    # Snapshot chat history (plain list, safe to copy)
    chat_history_snapshot = list(st.session_state.chat_history)

    # Snapshot all file bytes before threading (UploadedFile is not thread-safe)
    pairs_snapshot = []
    for pair in matched:
        pair.test_case.seek(0)
        pair.test_report.seek(0)
        pairs_snapshot.append({
            "name":         pair.name,
            "tc_bytes":     pair.test_case.read(),
            "tc_name":      pair.test_case.name,
            "tr_bytes":     pair.test_report.read(),
            "tr_name":      pair.test_report.name,
            "chat_history": chat_history_snapshot,
        })

    progress_bar = st.progress(0, text="Running comparisons…")
    total = len(pairs_snapshot)

    def _process(item: dict):
        result = compare(
            item["tc_bytes"], item["tc_name"],
            item["tr_bytes"], item["tr_name"],
            PROMPT_TEXT, item["chat_history"],
        )
        word_bytes = generate_word_report(result, item["tc_name"], item["tr_name"])
        return item["name"], result, word_bytes, item["tr_name"]

    with ThreadPoolExecutor(max_workers=min(total, 4)) as executor:
        futures = {executor.submit(_process, p): p["name"] for p in pairs_snapshot}
        done = 0
        for future in as_completed(futures):
            name, result, word_bytes, tr_name = future.result()
            st.session_state.results[name] = {
                "result":     result,
                "word_bytes": word_bytes,
                "tr_name":    tr_name,
            }
            done += 1
            progress_bar.progress(done / total, text=f"Processed {done}/{total}: {name}")

    progress_bar.empty()
    st.success(f"✅  {total} comparison(s) completed.")
    st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7 – Results Summary Table + Individual Downloads
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.results:
    st.divider()
    st.subheader("📊 Results Summary")

    rows = []
    for name, data in st.session_state.results.items():
        r = data["result"]
        rows.append({
            "Filename / 文件名":      name,
            "Overall / 整体结论":       r.get("overall", "N/A"),
            "Summary / 总结":       r.get("summary", ""),
        })

    def _highlight_overall_row(row):
        overall = str(row.get("Overall / 整体结论", ""))
        palette = {
            "符合":            "background-color: #C6EFCE",
            "不符合":          "background-color: #FFC7CE; color: #9C0006",
            "需人工复核后确认": "background-color: #FFEB9C",
            "error":          "background-color: #FFC7CE; color: #9C0006",
        }
        style = palette.get(overall, palette.get(overall.lower(), ""))
        return [style for _ in row]

    st.dataframe(
        pd.DataFrame(rows).style.apply(_highlight_overall_row, axis=1),
        use_container_width=True,
        hide_index=True,
    )

    # ── Per-file detailed issue table ──────────────────────────────
    def _color_severity_col(col):
        palette = {
            "严重":    "background-color: #FFC7CE",
            "一般":    "background-color: #FCE4D6",
            "轻微":    "background-color: #FFEB9C",
            "需人工复核": "background-color: #DDEBF7",
        }
        return [palette.get(str(v).strip(), "") for v in col]

    st.subheader("📋 Detailed Results")
    for name, data in st.session_state.results.items():
        r = data["result"]
        issues          = r.get("issues", [])
        manual_images   = r.get("manual_review_images", [])
        recommendations = r.get("recommendations", "")
        overall  = r.get("overall", "N/A")
        tr_name  = data.get("tr_name", name)
        is_error = str(overall).lower() == "error"
        label = f"**{tr_name}** — ⚠️ {overall} (系统/接口错误，非测试结论)" if is_error else f"**{tr_name}** — {overall}"
        with st.expander(label, expanded=True):
            if is_error:
                st.error(r.get("summary", ""))
            if issues:
                st.caption("Issue list / 问题清单")
                df_issues = pd.DataFrame([
                    {
                        "#":                       it.get("no", idx),
                        "问题描述 / Description":     it.get("description", ""),
                        "测试要求的内容 / Requirement": it.get("requirement_text", ""),
                        "测试报告的内容 / Report":     it.get("report_text", ""),
                        "问题归类 / Category":        it.get("category", ""),
                        "严重程度 / Severity":        it.get("severity", ""),
                    }
                    for idx, it in enumerate(issues, 1)
                ])
                st.dataframe(
                    df_issues.style.apply(_color_severity_col, subset=["严重程度 / Severity"]),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info("No issues found — see the downloaded Word report for details. / 未发现问题")

            if manual_images:
                st.caption("Manual review image list / 待人工核查图片清单")
                df_images = pd.DataFrame([
                    {
                        "位置 / Location":         img.get("location", ""),
                        "对应要求 / Requirement":    img.get("requirement_text", ""),
                        "报告描述 / Report":         img.get("report_text", ""),
                        "无法识别原因 / Reason":      img.get("reason", ""),
                    }
                    for img in manual_images
                ])
                st.dataframe(df_images, use_container_width=True, hide_index=True)

            if recommendations:
                st.caption("Recommendations / 建议")
                st.markdown(recommendations)

    st.subheader("📄 Individual Downloads")
    n_cols = min(len(st.session_state.results), 4)
    dl_cols = st.columns(n_cols)
    for idx, (name, data) in enumerate(st.session_state.results.items()):
        with dl_cols[idx % n_cols]:
            st.download_button(
                label=f"⬇  {name}.docx",
                data=data["word_bytes"],
                file_name=f"{name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key=f"dl_{idx}_{Path(name).stem}",
            )

