import streamlit as st
import anthropic
import json
from excel_builder import build_excel
import tempfile, os

st.set_page_config(page_title="Climate CBA Tool", page_icon="🌍", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500&display=swap');
* { font-family: 'IBM Plex Sans', sans-serif; }
.stApp { background: #f8f9fa; }
.block-container { max-width: 860px; padding: 2.5rem 2rem; }
h1 { font-size: 1.6rem; font-weight: 500; color: #1a1a2e; letter-spacing: -0.02em; }
.chat-msg-user {
    background: #1a1a2e; color: #fff;
    border-radius: 12px 12px 2px 12px;
    padding: 0.75rem 1rem; margin: 0.5rem 0; max-width: 80%; margin-left: auto;
    font-size: 0.92rem;
}
.chat-msg-ai {
    background: #fff; color: #1a1a2e; border: 1px solid #e2e8f0;
    border-radius: 12px 12px 12px 2px;
    padding: 0.75rem 1rem; margin: 0.5rem 0; max-width: 85%;
    font-size: 0.92rem; line-height: 1.6;
}
.stTextInput input {
    border: 1.5px solid #e2e8f0 !important; border-radius: 8px !important;
    font-size: 0.95rem !important; padding: 0.6rem 0.9rem !important;
}
.stTextInput input:focus { border-color: #3b82f6 !important; box-shadow: none !important; }
.stButton > button {
    background: #1a1a2e !important; color: #fff !important; border: none !important;
    border-radius: 8px !important; font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.8rem !important; padding: 0.55rem 1.5rem !important;
}
.stButton > button:hover { background: #16213e !important; }
.stDownloadButton > button {
    background: #059669 !important; color: #fff !important; border: none !important;
    border-radius: 8px !important; font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.85rem !important; padding: 0.6rem 1.5rem !important; width: 100%;
}
.status-badge {
    display: inline-block; background: #eff6ff; color: #3b82f6;
    font-family: 'IBM Plex Mono', monospace; font-size: 0.7rem;
    padding: 0.15rem 0.5rem; border-radius: 4px; margin-bottom: 1rem;
}
</style>
""", unsafe_allow_html=True)

# ── Session state init ──────────────────────────────────────────────────────────
for k, v in {
    "messages": [], "stage": "problem", "analysis_data": None,
    "api_key": "", "problem_text": ""
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("## 🌍 Climate Adaptation CBA Tool")

# API key
with st.expander("⚙️ API Key", expanded=not st.session_state.api_key):
    k = st.text_input("Anthropic API Key", type="password", value=st.session_state.api_key)
    if k: st.session_state.api_key = k

st.markdown("---")

# ── Stage badge ───────────────────────────────────────────────────────────────
stages = {"problem": "1/3 · Problem", "measures": "2/3 · Measures", "data": "3/3 · Analysis", "done": "✓ Complete"}
st.markdown(f'<div class="status-badge">{stages.get(st.session_state.stage, "")}</div>', unsafe_allow_html=True)

# ── Chat display ───────────────────────────────────────────────────────────────
chat_container = st.container()
with chat_container:
    if not st.session_state.messages:
        st.markdown("""
        <div class="chat-msg-ai">
        Hello! I'm here to help you build a cost-benefit analysis for climate adaptation measures.<br><br>
        <b>Describe the climate problem</b> you're working on — include location, population, climate hazard, and any economic context you have.
        </div>
        """, unsafe_allow_html=True)
    for msg in st.session_state.messages:
        css = "chat-msg-user" if msg["role"] == "user" else "chat-msg-ai"
        st.markdown(f'<div class="{css}">{msg["content"]}</div>', unsafe_allow_html=True)

# ── Input ──────────────────────────────────────────────────────────────────────
if st.session_state.stage != "done":
    col1, col2 = st.columns([5, 1])
    with col1:
        user_input = st.text_input("", placeholder="Type your message...", label_visibility="collapsed", key="user_input")
    with col2:
        send = st.button("Send →")

    if send and user_input.strip() and st.session_state.api_key:
        st.session_state.messages.append({"role": "user", "content": user_input})
        client = anthropic.Anthropic(api_key=st.session_state.api_key)

        # ── Stage: problem description ─────────────────────────────────────────
        if st.session_state.stage == "problem":
            st.session_state.problem_text = user_input
            with st.spinner("Analyzing problem..."):
                resp = client.messages.create(
                    model="claude-opus-4-6", max_tokens=1200,
                    messages=[{
                        "role": "user",
                        "content": f"""You are an expert in climate adaptation economics.

The user described this climate problem:
"{user_input}"

Do two things:
1. Briefly confirm your understanding of the problem (2-3 sentences)
2. Propose 3-4 concrete adaptation measures relevant to this problem, based on the academic literature. For each measure give:
   - Name
   - 1-sentence description
   - Typical cost range
   - Main benefit categories

Then ask the user: "Which of these measures would you like to analyze? You can select all or a subset."

Keep your response concise and practical."""
                    }]
                )
            reply = resp.content[0].text
            st.session_state.messages.append({"role": "assistant", "content": reply})
            st.session_state.stage = "measures"
            st.rerun()

        # ── Stage: user selects measures ───────────────────────────────────────
        elif st.session_state.stage == "measures":
            with st.spinner("Identifying data needs..."):
                history = [{"role": m["role"], "content": m["content"]} for m in st.session_state.messages]
                history.append({
                    "role": "user",
                    "content": f"""The user selected: "{user_input}"

Now ask for the specific numerical data you need to run the CBA. 
Ask for the MINIMUM required inputs — only what you truly need for NPV and BCR calculations.
Group questions logically (costs, benefits, parameters).
Be specific about units.
Ask all questions in one message."""
                })
                resp = client.messages.create(model="claude-opus-4-6", max_tokens=800, messages=history)
            reply = resp.content[0].text
            st.session_state.messages.append({"role": "assistant", "content": reply})
            st.session_state.stage = "data"
            st.rerun()

        # ── Stage: user provides data → build analysis ─────────────────────────
        elif st.session_state.stage == "data":
            with st.spinner("Building analysis..."):
                history = [{"role": m["role"], "content": m["content"]} for m in st.session_state.messages]
                history.append({
                    "role": "user",
                    "content": f"""User provided: "{user_input}"

Extract all data and produce a JSON object for the CBA model. Respond ONLY with valid JSON, no markdown fences:

{{
  "problem_title": "short title",
  "problem_summary": "2-sentence summary",
  "discount_rate": 0.035,
  "time_horizon": 30,
  "currency": "USD",
  "currency_unit": "millions",
  "measures": [
    {{
      "name": "Measure name",
      "description": "What it involves",
      "category": "Infrastructure|Nature-based|Policy|Technology",
      "capex": 10.0,
      "annual_opex": 0.5,
      "annual_benefit": 3.0,
      "benefit_types": ["flood protection", "tourism"],
      "co_benefits": "Non-monetary benefits",
      "lifetime_years": 30,
      "feasibility": "High|Medium|Low",
      "uncertainty": "Low|Medium|High",
      "data_source": "Literature estimate / User provided"
    }}
  ],
  "sensitivity_vars": [
    {{"name": "Discount Rate", "base": 0.035, "low": 0.02, "high": 0.07, "unit": "%"}},
    {{"name": "Annual Benefit", "base": 1.0, "low": 0.5, "high": 1.5, "unit": "multiplier"}}
  ],
  "key_assumptions": "Main assumptions and limitations",
  "data_gaps": "What additional data would improve this"
}}

If user data is missing for some fields, use reasonable literature-based estimates and note them in data_source."""
                })
                resp = client.messages.create(model="claude-opus-4-6", max_tokens=2000, messages=history)

            raw = resp.content[0].text.strip()
            if raw.startswith("```"):
                raw = raw.split("```")[1]
                if raw.startswith("json"): raw = raw[4:]
                raw = raw.rsplit("```", 1)[0]

            try:
                data = json.loads(raw)
                st.session_state.analysis_data = data
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": f"✅ Analysis complete for **{data['problem_title']}**. I've analyzed {len(data['measures'])} measure(s). Click **Download Excel** below to get your full CBA model with sensitivity analysis."
                })
                st.session_state.stage = "done"
            except Exception as e:
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": f"I had trouble parsing the analysis. Could you provide the data again more clearly? (Error: {e})"
                })
            st.rerun()

# ── Excel download ─────────────────────────────────────────────────────────────
if st.session_state.stage == "done" and st.session_state.analysis_data:
    st.markdown("---")
    col_l, col_r = st.columns([1, 2])
    with col_l:
        with st.spinner("Building Excel..."):
            tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
            build_excel(st.session_state.analysis_data, tmp.name)
            with open(tmp.name, "rb") as f:
                excel_bytes = f.read()
            os.unlink(tmp.name)

        title = st.session_state.analysis_data.get("problem_title", "CBA").replace(" ", "_")
        st.download_button(
            "⬇ Download Excel CBA Model",
            excel_bytes,
            file_name=f"CBA_{title}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col_r:
        st.markdown("**Your Excel file includes:**")
        st.markdown("- 📋 Inputs & Assumptions sheet (blue = editable)\n- 📊 CBA Results with NPV & BCR\n- 📈 Sensitivity Analysis (tornado chart data)\n- 📁 Summary sheet")

    if st.button("🔄 Start New Analysis"):
        for k in ["messages", "stage", "analysis_data", "problem_text"]:
            st.session_state[k] = [] if k == "messages" else ("problem" if k == "stage" else None if k == "analysis_data" else "")
        st.rerun()
