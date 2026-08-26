"""
app.py — Streamlit UI for Excel Reader.

Upload any Excel/CSV → Ask anything → Get exact text-readable answers.
Uses Google Gemini (FREE) or Anthropic Claude (paid).
"""

import os
import tempfile
import streamlit as st
from pathlib import Path

from excel_parser import parse_excel, sheets_to_text, parsed_to_json
from llm_analyzer import (
    ask_question, load_dataframes, build_schema_prompt,
    summarize_excel, convert_json_to_text_llm, convert_text_to_json_llm,
)

# ── Page config ─────────────────────────────────────────────────────────
st.set_page_config(page_title="Excel Reader", page_icon="📊", layout="wide")

st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Barlow:wght@500;700;800&display=swap">
<style>
.eyebrow-tag{display:inline-block;background:#FFE600;color:#1A1A1A;font-family:'Barlow',sans-serif;
  font-weight:700;font-size:.72rem;letter-spacing:.08em;text-transform:uppercase;
  padding:4px 11px;border-radius:2px;margin-bottom:.7rem}
.main-hdr{font-family:'Barlow',sans-serif;font-size:2.3rem;font-weight:800;
  color:#1A1A1A;letter-spacing:-0.01em;margin-bottom:.25rem}
.sub-hdr{font-size:1rem;color:#4A4A4A;margin-bottom:1rem}
.hdr-rule{border:none;border-top:5px solid #FFE600;margin:0 0 1.7rem;width:64px}
.badge-code{background:#E3F1EC;color:#1F6F54;padding:2px 10px;border-radius:2px;font-size:.75rem;font-weight:700;letter-spacing:.02em;text-transform:uppercase}
.badge-llm{background:#FBF0DC;color:#8A5A00;padding:2px 10px;border-radius:2px;font-size:.75rem;font-weight:700;letter-spacing:.02em;text-transform:uppercase}
</style>
""", unsafe_allow_html=True)

# ── Session state ───────────────────────────────────────────────────────
for k, v in {"parsed":None,"text":None,"dfs":None,"schema":None,"history":[],"fname":None}.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ── Sidebar ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Settings")

    provider = st.radio("LLM Provider", ["Gemini (FREE)", "Anthropic (Paid)"], index=0)

    if provider == "Gemini (FREE)":
        api_key = st.text_input("Gemini API Key", type="password",
                                value=os.getenv("GEMINI_API_KEY", ""),
                                help="Get FREE at aistudio.google.com/apikey")
        if api_key:
            os.environ["GEMINI_API_KEY"] = api_key
            os.environ.pop("ANTHROPIC_API_KEY", None)
        model = st.selectbox("Model", [
            "gemini-3.6-flash",
            "gemini-3.6-flash",
        ])
        st.caption("✅ Gemini is completely FREE")
    else:
        api_key = st.text_input("Anthropic API Key", type="password",
                                value=os.getenv("ANTHROPIC_API_KEY", ""))
        if api_key:
            os.environ["ANTHROPIC_API_KEY"] = api_key
            os.environ.pop("GEMINI_API_KEY", None)
        model = st.selectbox("Model", [
            "claude-sonnet-4-20250514",
            "claude-haiku-4-5-20251001",
        ])

    show_code = st.toggle("Show pandas code", value=False)

    st.divider()
    if st.session_state.dfs:
        st.markdown(f"**📁 {st.session_state.fname}**")
        for name, df in st.session_state.dfs.items():
            st.markdown(f"📋 **{name}** — {df.shape[0]:,}R × {df.shape[1]}C")
        total = sum(d.shape[0] for d in st.session_state.dfs.values())
        st.caption(f"✅ All {total:,} rows loaded")

    st.divider()
    c1, c2 = st.columns(2)
    if c1.button("🗑️ Clear Chat"):
        st.session_state.history = []; st.rerun()
    if c2.button("📤 New File"):
        for k in ("parsed","text","dfs","schema","fname"): st.session_state[k] = None
        st.session_state.history = []; st.rerun()


def _has_key():
    return bool(os.getenv("GEMINI_API_KEY") or os.getenv("ANTHROPIC_API_KEY"))


# ── Header ──────────────────────────────────────────────────────────────
st.markdown(
    '<span class="eyebrow-tag">AI-Powered Analysis</span>'
    '<p class="main-hdr">Excel Reader</p>'
    '<p class="sub-hdr">Upload any Excel/CSV → ask anything → '
    'get <b>exact</b> answers (computed by Python, formatted by AI)</p>'
    '<hr class="hdr-rule">',
    unsafe_allow_html=True,
)


# ── File upload helper ──────────────────────────────────────────────────
def _load(uploaded):
    suf = Path(uploaded.name).suffix
    with tempfile.NamedTemporaryFile(delete=False, suffix=suf) as tmp:
        tmp.write(uploaded.getbuffer()); tmp_path = tmp.name
    parsed = parse_excel(tmp_path, uploaded.name)
    dfs = load_dataframes(tmp_path, uploaded.name)
    os.unlink(tmp_path)
    st.session_state.parsed = parsed
    st.session_state.text = sheets_to_text(parsed)
    st.session_state.dfs = dfs
    st.session_state.schema = build_schema_prompt(dfs)
    st.session_state.fname = uploaded.name
    st.session_state.history = []
    return len(dfs), sum(d.shape[0] for d in dfs.values())


# ── Upload zone ─────────────────────────────────────────────────────────
if not st.session_state.dfs:
    st.markdown("### 📂 Upload Your File")
    up = st.file_uploader("Drop Excel/CSV here", type=["xlsx","xls","xlsm","csv","tsv"])
    if up:
        with st.spinner("🔍 Parsing (merged cells, headers, all sheets)..."):
            try:
                ns, nr = _load(up)
                st.success(f"✅ **{up.name}** — {ns} sheet(s), {nr:,} rows loaded!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ {e}")
else:
    with st.expander("📂 Upload a different file"):
        up = st.file_uploader("New file", type=["xlsx","xls","xlsm","csv","tsv"], key="re")
        if up:
            with st.spinner("Parsing..."):
                try: _load(up); st.rerun()
                except Exception as e: st.error(f"❌ {e}")


# ── Tabs ────────────────────────────────────────────────────────────────
tab_qa, tab_conv, tab_raw = st.tabs(["💬 Q&A Chat", "🔄 JSON ↔ Text", "📋 Data Preview"])

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB 1 — Q&A
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab_qa:
    if not st.session_state.dfs:
        st.info("👆 Upload a file first.")
    else:
        dfs = st.session_state.dfs
        mc = st.columns(4)
        mc[0].metric("Sheets", len(dfs))
        mc[1].metric("Total Rows", f"{sum(d.shape[0] for d in dfs.values()):,}")
        mc[2].metric("Max Cols", max(d.shape[1] for d in dfs.values()))
        merged = sum(len(s.get("merged_regions",[])) for s in (st.session_state.parsed or {}).get("sheets",[]))
        mc[3].metric("Merged", merged)
        st.divider()

        # Quick actions
        q1, q2, q3 = st.columns(3)
        if q1.button("📝 Auto-Summarize", use_container_width=True):
            if not _has_key(): st.error("⚠️ Set API key in sidebar.")
            else:
                with st.spinner("Summarizing..."):
                    r = summarize_excel(dfs, st.session_state.schema, model)
                st.session_state.history += [
                    {"role":"user","content":"Summarize this data."},
                    {"role":"assistant",**r},
                ]; st.rerun()

        if q2.button("📊 Column Details", use_container_width=True):
            parts = []
            for n,d in dfs.items():
                parts.append(f"**{n}** ({d.shape[0]} rows)")
                for c in d.columns:
                    parts.append(f"  - `{c}` — {d[c].dtype} ({d[c].notna().sum()} vals)")
            st.session_state.history += [
                {"role":"user","content":"Show column details."},
                {"role":"assistant","answer":"\n".join(parts),"method":"direct"},
            ]; st.rerun()

        if q3.button("🔢 Quick Stats", use_container_width=True):
            if not _has_key(): st.error("⚠️ Set API key in sidebar.")
            else:
                with st.spinner("Computing..."):
                    r = ask_question(dfs, st.session_state.schema,
                        "For every numeric column show count, min, max, mean, sum as a table.",
                        st.session_state.text, model)
                st.session_state.history += [
                    {"role":"user","content":"Show statistics for all numeric columns."},
                    {"role":"assistant",**r},
                ]; st.rerun()

        st.divider()

        # Chat messages
        for msg in st.session_state.history:
            with st.chat_message(msg["role"]):
                st.markdown(msg.get("answer", msg.get("content","")))
                if msg["role"] == "assistant":
                    m = msg.get("method","")
                    if m == "code_execution":
                        st.markdown('<span class="badge-code">✅ Computed by Python (exact)</span>', unsafe_allow_html=True)
                    elif m == "direct_llm":
                        st.markdown('<span class="badge-llm">⚠️ LLM analysis (verify numbers)</span>', unsafe_allow_html=True)
                    if show_code and msg.get("code"):
                        with st.expander("🐍 Code"):
                            st.code(msg["code"], language="python")
                        if msg.get("raw_output"):
                            with st.expander("📋 Raw Output"):
                                st.text(msg["raw_output"])

        # Input
        if prompt := st.chat_input("Ask anything about your data..."):
            if not _has_key():
                st.error("⚠️ Enter API key in sidebar.")
            else:
                with st.chat_message("user"):
                    st.markdown(prompt)
                st.session_state.history.append({"role":"user","content":prompt})
                with st.chat_message("assistant"):
                    with st.spinner("🧠 Understanding → 🐍 Computing → ✍️ Formatting..."):
                        r = ask_question(dfs, st.session_state.schema, prompt,
                                         st.session_state.text, model)
                    st.markdown(r["answer"])
                    m = r.get("method","")
                    if m == "code_execution":
                        st.markdown('<span class="badge-code">✅ Computed by Python (exact)</span>', unsafe_allow_html=True)
                    elif m == "direct_llm":
                        st.markdown('<span class="badge-llm">⚠️ LLM analysis (verify numbers)</span>', unsafe_allow_html=True)
                    if show_code and r.get("code"):
                        with st.expander("🐍 Code"):
                            st.code(r["code"], language="python")
                st.session_state.history.append({"role":"assistant",**r})


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB 2 — JSON ↔ Text
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab_conv:
    st.markdown("### 🔄 JSON ↔ Text Converter")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### JSON → Text")
        ji = st.text_area("Paste JSON", height=200, key="ji",
                          placeholder='{"name":"Vedansh","role":"intern"}')
        if st.button("Convert →", key="j2t"):
            if ji.strip() and _has_key():
                with st.spinner("Converting..."):
                    st.text_area("Result", convert_json_to_text_llm(ji, model), height=200, key="jo")
    with c2:
        st.markdown("#### Text → JSON")
        ti = st.text_area("Paste text", height=200, key="ti",
                          placeholder="Name: Vedansh\nRole: Intern")
        if st.button("Convert →", key="t2j"):
            if ti.strip() and _has_key():
                with st.spinner("Converting..."):
                    st.code(convert_text_to_json_llm(ti, model), language="json")

    st.divider()
    if st.session_state.parsed:
        st.markdown("#### 📤 Export Parsed Excel")
        e1, e2 = st.columns(2)
        stem = Path(st.session_state.fname).stem
        e1.download_button("⬇️ JSON", parsed_to_json(st.session_state.parsed),
                           f"{stem}.json", "application/json")
        e2.download_button("⬇️ Text", st.session_state.text,
                           f"{stem}.txt", "text/plain")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB 3 — Data Preview
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab_raw:
    if not st.session_state.dfs:
        st.info("👆 Upload a file first.")
    else:
        dfs = st.session_state.dfs
        sel = st.selectbox("Sheet", list(dfs.keys()))
        df = dfs[sel]
        st.markdown(f"**{df.shape[0]:,} rows × {df.shape[1]} columns**")
        st.dataframe(df, use_container_width=True, height=500)
        with st.expander("📊 Statistics"):
            num = df.select_dtypes(include=["number"])
            if not num.empty: st.dataframe(num.describe(), use_container_width=True)
            else: st.info("No numeric columns.")
        if st.session_state.parsed:
            ps = next((s for s in st.session_state.parsed["sheets"] if s["name"]==sel), None)
            if ps and ps.get("merged_regions"):
                with st.expander(f"🔗 Merged Regions ({len(ps['merged_regions'])})"):
                    st.code("\n".join(ps["merged_regions"]))

st.divider()
st.caption("**How it works:** Question → LLM writes pandas code → Python computes exact answer → LLM formats readable text. Numbers are never guessed.")
