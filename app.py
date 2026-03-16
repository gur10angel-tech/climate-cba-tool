import streamlit as st
import anthropic
import boto3
import json
import re
import copy
from excel_builder import build_excel
import tempfile, os
import markdown as md_lib
import PyPDF2
import pandas as pd
import io
import plotly.graph_objects as go


def query_bedrock_kb(query: str) -> str:
    """Query the Bedrock Knowledge Base. Returns retrieved text or empty string on failure."""
    kb_id     = st.session_state.get("kb_id", "").strip()
    aws_key   = st.session_state.get("aws_key", "").strip()
    aws_secret= st.session_state.get("aws_secret", "").strip()
    region    = st.session_state.get("aws_region", "eu-north-1").strip()

    if not (kb_id and aws_key and aws_secret):
        return ""

    try:
        client = boto3.client(
            "bedrock-agent-runtime",
            region_name=region,
            aws_access_key_id=aws_key,
            aws_secret_access_key=aws_secret,
        )
        resp = client.retrieve(
            knowledgeBaseId=kb_id,
            retrievalQuery={"text": query},
            retrievalConfiguration={"vectorSearchConfiguration": {"numberOfResults": 6}},
        )
        chunks = []
        for r in resp.get("retrievalResults", []):
            text = r.get("content", {}).get("text", "")
            src  = r.get("location", {}).get("s3Location", {}).get("uri", "")
            src_name = src.split("/")[-1] if src else "unknown"
            if text:
                chunks.append(f"[{src_name}]\n{text}")
        return "\n\n---\n\n".join(chunks)
    except Exception as e:
        return f"[KB unavailable: {e}]"

st.set_page_config(page_title="Climate CBA Tool", page_icon="🌍", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
* { font-family: 'Inter', 'Segoe UI', sans-serif; }

.stApp {
    background: #f8fafc;
}

.block-container {
    max-width: 860px;
    padding: 2.5rem 2rem;
    background: #ffffff;
    border-radius: 16px;
    box-shadow: 0 18px 45px rgba(15, 23, 42, 0.06);
}

h1 {
    font-size: 1.6rem;
    font-weight: 500;
    color: #052e16;
    letter-spacing: -0.02em;
}

.chat-msg-user {
    background: #15803d;
    color: #f9fafb;
    border-radius: 12px 12px 4px 12px;
    padding: 0.75rem 1rem;
    margin: 0.5rem 0;
    max-width: 80%;
    margin-left: auto;
    font-size: 0.92rem;
}

.chat-msg-ai {
    background: #f9fafb;
    color: #022c22;
    border: 1px solid #e2e8f0;
    border-radius: 12px 12px 12px 4px;
    padding: 0.75rem 1rem;
    margin: 0.5rem 0;
    max-width: 85%;
    font-size: 0.92rem;
    line-height: 1.6;
}

.stTextInput input {
    border: 1.5px solid #d4d4d4 !important;
    border-radius: 999px !important;
    font-size: 0.95rem !important;
    padding: 0.6rem 0.9rem !important;
    background: #f9fafb !important;
}

.stTextInput input:focus {
    border-color: #16a34a !important;
    box-shadow: 0 0 0 1px rgba(22, 163, 74, 0.35) !important;
    background: #ffffff !important;
}

.stTextInput input:focus-visible {
    outline: none !important;
}

[data-baseweb="input"]:focus-within {
    outline: none !important;
    box-shadow: none !important;
}

.stTextInput > div:focus-within {
    outline: none !important;
    border: none !important;
}


[data-baseweb="input"] > div,
[data-baseweb="input"] > div:focus,
[data-baseweb="input"] > div:focus-within {
    outline: none !important;
    box-shadow: none !important;
    border-color: #d4d4d4 !important;
}
[data-baseweb="input"]:focus-within > div {
    border-color: #16a34a !important;
    box-shadow: none !important;
    outline: none !important;
}

.stTextArea textarea {
    border: 1.5px solid #d4d4d4 !important;
    border-radius: 12px !important;
    font-size: 0.95rem !important;
    padding: 0.6rem 0.9rem !important;
    background: #f9fafb !important;
    resize: vertical;
}

.stTextArea textarea:focus {
    border-color: #16a34a !important;
    box-shadow: 0 0 0 1px rgba(22, 163, 74, 0.35) !important;
    background: #ffffff !important;
    outline: none !important;
}

.chat-msg-ai table {
    width: 100%;
    border-collapse: collapse;
    margin: 0.5rem 0;
}

.chat-msg-ai th, .chat-msg-ai td {
    padding: 6px 10px;
    border: 1px solid #d1fae5;
    font-size: 0.88rem;
}

.chat-msg-ai th {
    background: #ecfdf5;
    font-weight: 600;
    font-size: 0.9rem;
    color: #022c22;
}

.chat-msg-ai td:first-child {
    font-size: 0.97rem;
    font-weight: 600;
}

.stButton > button {
    background: #15803d !important;
    color: #f9fafb !important;
    border: none !important;
    border-radius: 999px !important;
    font-family: 'Inter', 'Segoe UI', sans-serif !important;
    font-size: 0.8rem !important;
    padding: 0.55rem 1.5rem !important;
}

.stButton > button:hover {
    background: #166534 !important;
}


[data-testid="stFormSubmitButton"] > button {
    background: #15803d !important;
    color: #f9fafb !important;
    border: none !important;
    border-radius: 999px !important;
    font-family: 'Inter', 'Segoe UI', sans-serif !important;
    font-size: 0.8rem !important;
    padding: 0.55rem 1.5rem !important;
}
[data-testid="stFormSubmitButton"] > button:hover {
    background: #166534 !important;
}

.stDownloadButton > button {
    background: #15803d !important;
    color: #f9fafb !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'Inter', 'Segoe UI', sans-serif !important;
    font-size: 0.85rem !important;
    padding: 0.6rem 1.5rem !important;
    width: 100%;
}

.status-badge {
    display: inline-block;
    background: #ecfdf3;
    color: #15803d;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    padding: 0.15rem 0.6rem;
    border-radius: 999px;
    margin-bottom: 0.4rem;
}

.specialist-badge {
    display: inline-block;
    background: #bbf7d0;
    color: #14532d;
    font-family: 'Inter', 'Segoe UI', sans-serif;
    font-size: 0.7rem;
    padding: 0.15rem 0.6rem;
    border-radius: 999px;
    margin-bottom: 1rem;
    margin-left: 0.5rem;
}

.status-badge {
    font-family: 'Inter', 'Segoe UI', sans-serif;
}

/* st.metric styling */
[data-testid="stMetric"] label { color: #475569 !important; font-size: 0.82rem !important; }
[data-testid="stMetricValue"] { color: #14532d !important; font-size: 1.6rem !important; font-weight: 600 !important; }
[data-testid="stMetricDelta"] { font-size: 0.78rem !important; }

/* Tab accent */
.stTabs [data-baseweb="tab-list"] { border-bottom: 2px solid #e2e8f0; }
.stTabs [aria-selected="true"] { border-bottom: 3px solid #15803d !important; color: #14532d !important; font-weight: 600; }

/* ── Feature 4a: Formula rendering — no monospace code blocks in chat ── */
.chat-msg-ai pre, .chat-msg-ai code {
    font-family: 'Inter', 'Segoe UI', sans-serif !important;
    background: transparent !important;
    border: none !important;
    padding: 0 !important;
    font-size: 0.92rem !important;
    color: inherit !important;
    white-space: pre-wrap !important;
}
.chat-msg-ai pre {
    border-left: 3px solid #bbf7d0 !important;
    padding-left: 0.75rem !important;
    margin: 0.5rem 0 !important;
    background: #f0fdf4 !important;
}

/* ── Feature 1: Landing banner ── */
.landing-hero {
    text-align: center;
    padding: 1.8rem 1rem 1.2rem;
    border-bottom: 1px solid #e2e8f0;
    margin-bottom: 1.2rem;
}
.landing-hero h2 {
    font-size: 1.7rem;
    font-weight: 600;
    color: #052e16;
    margin: 0 0 0.4rem;
    letter-spacing: -0.02em;
}
.landing-hero p {
    color: #64748b;
    font-size: 0.93rem;
    max-width: 520px;
    margin: 0 auto;
    line-height: 1.6;
}
.cap-grid {
    display: flex;
    gap: 0.85rem;
    margin: 1.2rem 0 0.8rem;
    flex-wrap: wrap;
}
.cap-card {
    flex: 1;
    min-width: 155px;
    background: #f0fdf4;
    border: 1px solid #bbf7d0;
    border-radius: 12px;
    padding: 0.9rem 1rem;
    font-size: 0.86rem;
    color: #14532d;
    line-height: 1.5;
}
.cap-card strong {
    display: block;
    font-size: 0.92rem;
    margin-bottom: 0.25rem;
    color: #052e16;
}
.usecase-row {
    display: flex;
    gap: 0.6rem;
    margin: 0.6rem 0 1.2rem;
    flex-wrap: wrap;
}
.usecase-tag {
    background: #fefce8;
    border: 1px solid #fde68a;
    color: #713f12;
    border-radius: 999px;
    padding: 0.22rem 0.8rem;
    font-size: 0.8rem;
}
.cta-hint {
    font-size: 0.86rem;
    color: #64748b;
    margin-top: 1rem;
    font-style: italic;
}
</style>
""", unsafe_allow_html=True)

# ── Climate Scenario Engine constants (must be before sidebar) ───────────────────
CLIMATE_SCENARIOS = {
    "Stable (Baseline)": {
        "heat_days_growth_rate": 0.0,
        "cdd_increment": 0.0,
        "flood_frequency_multiplier": 0.0,
        "sea_level_rise_cm_per_decade": 0.0,
        "drought_intensity_multiplier": 0.0,
        "source": "No climate change assumed — static benefits",
    },
    "RCP 4.5 (Moderate)": {
        "heat_days_growth_rate": 0.02,
        "cdd_increment": 4.5,
        "flood_frequency_multiplier": 0.15,
        "sea_level_rise_cm_per_decade": 4.0,
        "drought_intensity_multiplier": 0.08,
        "source": "IPCC AR6 WGI (2021); IMS Coastal Plain 2050 projection",
    },
    "RCP 8.5 (Accelerated)": {
        "heat_days_growth_rate": 0.04,
        "cdd_increment": 9.0,
        "flood_frequency_multiplier": 0.30,
        "sea_level_rise_cm_per_decade": 8.5,
        "drought_intensity_multiplier": 0.18,
        "source": "IPCC AR6 WGI (2021); IMS high-emission trajectory",
    },
    "Custom": {
        "heat_days_growth_rate": 0.02,
        "cdd_increment": 5.0,
        "flood_frequency_multiplier": 0.15,
        "sea_level_rise_cm_per_decade": 4.0,
        "drought_intensity_multiplier": 0.08,
        "source": "User-defined",
    },
}

BENEFIT_DRIVER_MAP = {
    "avoided_mortality":          "exponential",
    "morbidity_savings":          "exponential",
    "energy_savings":             "linear",
    "property_value_uplift":      "static",
    "generic_annual":             "static",
    "avoided_mortality_npv":      "exponential",
    "morbidity_savings_npv":      "exponential",
    "skin_cancer_prevention_npv": "exponential",
    "energy_savings_npv":         "linear",
    "carbon_sequestration_npv":   "static",
    "runoff_reduction_npv":       "static",
    "air_quality_npv":            "static",
    "habitat_creation_npv":       "static",
    "property_value_uplift_npv":  "static",
    "roof_longevity_npv":         "static",
}

# ── Sidebar ─────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## Settings")

    with st.expander("Climate Scenario", expanded=False):
        selected = st.selectbox(
            "Select Climate Scenario",
            list(CLIMATE_SCENARIOS.keys()),
            key="scenario",
        )
        rates = CLIMATE_SCENARIOS[selected].copy()
        challenge = st.session_state.get("challenge_type", "general")

        if selected == "Custom":
            rates["heat_days_growth_rate"] = st.number_input(
                "Heat Days Growth Rate (r/yr)", 0.0, 0.10,
                value=float(rates["heat_days_growth_rate"]), step=0.005, format="%.3f",
                help="Exponential growth factor for heatwave days. Affects Mortality & Morbidity.",
            )
            rates["cdd_increment"] = st.number_input(
                "CDD Annual Increment", 0.0, 20.0,
                value=float(rates["cdd_increment"]), step=0.5,
                help="Additional Cooling Degree Days per year (linear). Affects Energy Savings.",
            )
            rates["flood_frequency_multiplier"] = st.number_input(
                "Flood Freq. Multiplier (/decade)", 0.0, 1.0,
                value=float(rates["flood_frequency_multiplier"]), step=0.05, format="%.2f",
                help="Fractional increase in flood event frequency per decade (e.g. 0.15 = +15%/decade).",
            )
            rates["sea_level_rise_cm_per_decade"] = st.number_input(
                "Sea Level Rise (cm/decade)", 0.0, 30.0,
                value=float(rates["sea_level_rise_cm_per_decade"]), step=0.5,
                help="Mean sea level rise in cm per decade.",
            )
            rates["drought_intensity_multiplier"] = st.number_input(
                "Drought Intensity Mult. (/decade)", 0.0, 0.5,
                value=float(rates["drought_intensity_multiplier"]), step=0.02, format="%.2f",
                help="Fractional increase in drought intensity per decade.",
            )
        else:
            if challenge in ("heat", "general"):
                st.caption(f"🌡️ Heat Days growth: **{rates['heat_days_growth_rate']*100:.1f}%/yr** (exponential)")
                st.caption(f"🌡️ CDD increment: **{rates['cdd_increment']:.1f} CDDs/yr** (linear)")
            if challenge in ("flood", "general"):
                st.caption(f"🌊 Flood freq. increase: **{rates['flood_frequency_multiplier']*100:.0f}%/decade**")
                st.caption(f"🌊 Sea level rise: **{rates['sea_level_rise_cm_per_decade']:.1f} cm/decade**")
            if challenge in ("drought", "general"):
                st.caption(f"💧 Drought intensity: **+{rates['drought_intensity_multiplier']*100:.0f}%/decade**")
            if challenge in ("heat", "general"):
                st.caption("Property / Carbon: **static** (asset-value driven)")
        st.caption(f"*Source: {rates['source']}*")
        st.session_state.escalation_rates = rates

    with st.expander("⚗️ What-If Analysis", expanded=False):
        st.caption("Move sliders to update all portfolio charts instantly. These values are also applied to the downloaded Excel.")
        st.slider("Discount Rate (%)", min_value=0.5, max_value=10.0, value=3.5, step=0.1,
                  key="sidebar_dr",
                  help="Sets the discount rate for portfolio charts AND the downloaded Excel model.")
        st.caption(
            f"Active: **DR = {st.session_state.get('sidebar_dr', 3.5):.1f}%** — applied to Excel on download"
        )

    with st.expander("Financial Parameters", expanded=False):
        st.number_input("Time Horizon (years)", min_value=5, max_value=100, value=50, step=5,
                        key="sidebar_horizon", help="Sets the time horizon for the downloaded Excel (overrides Claude's JSON default).")
        st.selectbox("Currency", ["NIS", "EUR", "USD"], key="sidebar_currency",
                     help="Reference value — actual currency is set by Claude from your data.")

    with st.expander("Methodology Reference", expanded=False):
        st.markdown("""
**VSL Derivation Chain**

| Step | Parameter | Value |
|------|-----------|-------|
| 1 | OECD Base VSL (2005 USD) | $3.0M |
| 2 | × CPI Multiplier (2005→2023) | 1.68 |
| 3 | × PPP Ratio (Israel/OECD) | 0.89 |
| 4 | × Income Elasticity | 1.00 |
| 5 | × FX Rate (NIS/USD) | 3.70 |
| 6 | = VSL (NIS) | ~16.6M |
| 7 | ÷ Life Expectancy (years) | 35 |
| 8 | = VSLY (NIS/yr) | ~474K |

**CDD / Heat-Mortality**
- CDD baseline (Tel Aviv, 21°C): 735
- Heat-mortality factor: 0.00083 (Gasparrini 2017)
- Morbidity multiplier: 10×

**NPV Formula**
`NPV = PV(Benefits) − PV(Costs)`
`BCR = PV(Benefits) / PV(Costs)`
""")

    with st.expander("About", expanded=False):
        st.markdown("""
**Climate Adaptation CBA Tool**

Produces a fully auditable Excel cost-benefit model for urban climate adaptation measures.

Every calculation is a live Excel formula traceable to peer-reviewed literature.

**Key sources:** Viscusi & Masterman (2017), Gasparrini et al. (2017) Lancet, WHO Heat Health Action Plan (2008), OECD ENV/WKP(2012)3.
""")

# ── Specialist keyword detection ────────────────────────────────────────────────
SHADE_KEYWORDS = [
    "shade", "shading", "tree", "trees", "boulevard", "canopy",
    "urban forest", "street tree", "street trees", "avenue", "allee",
    "urban canopy", "urban shade", "natural shade",
    "צל", "עצים", "עצי רחוב", "הצללה", "סככה",
]
ROOF_KEYWORDS = [
    "green roof", "green roofs", "rooftop vegetation", "rooftop garden",
    "living roof", "vegetated roof", "extensive roof", "intensive roof",
    "גג ירוק", "גגות ירוקים", "גינת גג",
]


def detect_specialist_type(text: str):
    t = text.lower()
    # Check green roof first — more specific, avoids "green" matching shading
    for kw in ROOF_KEYWORDS:
        if kw in t:
            return "green_roof"
    for kw in SHADE_KEYWORDS:
        if kw in t:
            return "natural_shading"
    return None


# ── Challenge type detection ─────────────────────────────────────────────────
_CHALLENGE_KEYWORDS = {
    "heat":    ["heat", "temperature", "cooling", "heatwave", "shade", "urban heat",
                "thermal comfort", "hot day", "cool down"],
    "flood":   ["flood", "sea level", "storm surge", "coastal", "drainage", "inundation",
                "precipitation", "runoff", "stormwater", "wetland"],
    "drought": ["drought", "water scarcity", "rainfall deficit", "irrigation",
                "dry spell", "water stress", "aquifer"],
}

def detect_challenge_type(text: str) -> str:
    """Returns 'heat' | 'flood' | 'drought' | 'general' based on problem description keywords."""
    tl = text.lower()
    scores = {ct: sum(1 for kw in kws if kw in tl) for ct, kws in _CHALLENGE_KEYWORDS.items()}
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "general"


# ── File parsing helpers ─────────────────────────────────────────────────────────
def _parse_uploaded_file(f) -> str:
    """Extract text from an uploaded PDF, Excel, or CSV file. Returns up to 8000 chars."""
    name = f.name.lower()
    try:
        raw = f.read()
        if name.endswith(".pdf"):
            reader = PyPDF2.PdfReader(io.BytesIO(raw))
            pages = [p.extract_text() or "" for p in reader.pages]
            return "\n\n".join(pages)[:8000]
        elif name.endswith((".xlsx", ".xls")):
            df_dict = pd.read_excel(io.BytesIO(raw), sheet_name=None)
            parts = [f"[Sheet: {sn}]\n{df.to_string(index=False, max_rows=50)}"
                     for sn, df in df_dict.items()]
            return "\n\n".join(parts)[:8000]
        elif name.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(raw))
            return df.to_string(index=False, max_rows=80)[:8000]
    except Exception as e:
        return f"[File parse error: {e}]"
    return ""


def _extract_structured_data(text: str) -> str:
    """Regex pre-parser: extract common CBA parameters from free text."""
    patterns = {
        "CAPEX":         r"(?:capex|capital cost|investment)[^\d]*([\d,\.]+)",
        "OPEX":          r"(?:opex|annual cost|maintenance)[^\d]*([\d,\.]+)",
        "Area (m²)":     r"([\d,\.]+)\s*(?:sq\.?\s*m|m²|sqm|square meter)",
        "Population":    r"population\s+(?:of\s+)?([\d,\.]+)",
        "VSL":           r"vsl[^\d]*([\d,\.]+)",
        "Discount rate": r"discount\s+rate[^\d]*([\d\.]+)\s*%?",
    }
    found = {}
    for label, pat in patterns.items():
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            found[label] = m.group(1).replace(",", "")
    if not found:
        return ""
    lines = ["STRUCTURED DATA EXTRACTED:"] + [f"  {k}: {v}" for k, v in found.items()]
    return "\n".join(lines)


# ── Benefit composition pie chart ────────────────────────────────────────────────
_GREEN_PALETTE = ["#15803d", "#16a34a", "#22c55e", "#4ade80", "#86efac", "#bbf7d0", "#dcfce7"]


def _build_benefit_pie(measure: dict, financials_entry: dict = None):
    """Build a donut pie chart of benefit composition for one measure. Returns go.Figure or None.
    When financials_entry is provided, uses PV-weighted health/energy/other breakdown."""
    labels, values = [], []

    # Prefer PV-weighted breakdown from financials engine
    if financials_entry:
        ph = financials_entry.get("pv_health", 0)
        pe = financials_entry.get("pv_energy", 0)
        po = financials_entry.get("pv_other",  0)
        if ph + pe + po > 0:
            if ph > 0: labels.append("Health Benefits"); values.append(ph)
            if pe > 0: labels.append("Energy Benefits"); values.append(pe)
            if po > 0: labels.append("Other / Ecosystem"); values.append(po)

    # Fallback: original component-level breakdown
    if not labels:
        adv = measure.get("advanced_benefits")
        if adv and isinstance(adv, dict):
            for k, v in adv.items():
                if isinstance(v, (int, float)) and v > 0:
                    labels.append(k.replace("_npv", "").replace("_", " ").title())
                    values.append(v)
        else:
            for comp in (measure.get("benefit_components") or []):
                val = comp.get("value", comp.get("annual_value", 0)) or 0
                if isinstance(val, (int, float)) and val > 0:
                    labels.append(comp.get("name", "Benefit"))
                    values.append(val)
    if not labels or sum(values) == 0:
        return None
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=0.35,
        marker=dict(colors=_GREEN_PALETTE[:len(labels)]),
        textinfo="percent+label",
        textposition="inside",
        insidetextorientation="radial",
        hovertemplate="%{label}<br>%{value:.3f} M<br>%{percent}<extra></extra>",
    )])
    fig.update_layout(
        title=dict(
            text=f"Benefit Composition — {measure.get('name', 'Measure')}",
            font=dict(family="Inter, Segoe UI, sans-serif", size=14, color="#052e16"),
            x=0.5,
        ),
        legend=dict(orientation="v", x=1.02, y=0.5),
        margin=dict(t=50, b=20, l=20, r=20),
        paper_bgcolor="white",
        plot_bgcolor="white",
        height=360,
        showlegend=True,
    )
    return fig


# ── 50-year benefit projection chart ─────────────────────────────────────────
def _build_projection_chart(analysis_data: dict, escalation_rates: dict, vsl_mult: float = 1.0):
    """Build Plotly line chart of 50-yr undiscounted annual benefits per measure."""
    measures = analysis_data.get("measures", [])
    if not measures:
        return None

    horizon = int(analysis_data.get("time_horizon", 50))
    r_exp   = escalation_rates.get("heat_days_growth_rate", 0.0)
    cdd_inc = escalation_rates.get("cdd_increment", 0.0)
    baseline_cdd = (analysis_data.get("cdd_params") or {}).get("annual_cdd", 735) or 735

    sp      = analysis_data.get("specialist_params") or {}
    mat_years = int(sp.get("maturity_years", 8) or 8)

    def _annuity_to_annual(npv_val, dr, yrs):
        """Reverse-engineer approximate annual base from a total NPV."""
        if dr > 0 and yrs > 0:
            af = (1 - (1 + dr) ** (-yrs)) / dr
            return npv_val / af if af > 0 else 0
        return npv_val / yrs if yrs > 0 else 0

    dr = analysis_data.get("discount_rate", 0.035) or 0.035

    years = list(range(1, horizon + 1))
    traces = []
    portfolio_total = [0.0] * horizon

    for mi, m in enumerate(measures):
        health_base = energy_base = other_base = 0.0
        adv = m.get("advanced_benefits")
        is_specialist = bool(adv and isinstance(adv, dict))

        if is_specialist:
            life = m.get("lifetime_years", horizon) or horizon
            for key, val in adv.items():
                if not isinstance(val, (int, float)) or val <= 0:
                    continue
                annual = _annuity_to_annual(val, dr, life)
                driver = BENEFIT_DRIVER_MAP.get(key, "static")
                if driver == "exponential":
                    health_base += annual
                elif driver == "linear":
                    energy_base += annual
                else:
                    other_base += annual
        else:
            for comp in (m.get("benefit_components") or []):
                val = comp.get("value", comp.get("annual_value", 0)) or 0
                if not isinstance(val, (int, float)) or val <= 0:
                    continue
                driver = BENEFIT_DRIVER_MAP.get(comp.get("type", "generic_annual"), "static")
                if driver == "exponential":
                    health_base += val
                elif driver == "linear":
                    energy_base += val
                else:
                    other_base += val

        health_base *= vsl_mult  # Apply What-If VSL multiplier

        stype = analysis_data.get("specialist_type")
        benefit_series = []
        for t in years:
            mat = min(t / mat_years, 1.0) if stype == "natural_shading" else 1.0
            exp_factor = (1 + r_exp) ** t
            cdd_factor = (baseline_cdd + cdd_inc * t) / baseline_cdd if baseline_cdd > 0 else 1.0
            h_t = health_base * mat * exp_factor
            e_t = energy_base * cdd_factor
            o_t = other_base
            benefit_series.append(h_t + e_t + o_t)

        for i, v in enumerate(benefit_series):
            portfolio_total[i] += v

        traces.append(go.Scatter(
            x=years, y=benefit_series,
            mode="lines", name=m.get("name", f"Measure {mi+1}"),
            line=dict(width=2),
        ))

    if not any(v > 0 for v in portfolio_total):
        return None

    traces.append(go.Scatter(
        x=years, y=portfolio_total,
        mode="lines", name="Total Portfolio",
        line=dict(color="#052e16", dash="dot", width=3),
    ))

    cur = analysis_data.get("currency", "")
    cur_unit = analysis_data.get("currency_unit", "millions")
    fig = go.Figure(data=traces)
    fig.update_layout(
        title=dict(
            text="50-Year Annual Benefit Projection (Undiscounted)",
            font=dict(family="Inter, Segoe UI, sans-serif", size=14, color="#052e16"),
            x=0.5,
        ),
        xaxis=dict(title="Year", gridcolor="#e2e8f0"),
        yaxis=dict(title=f"Annual Benefit ({cur_unit} {cur})", gridcolor="#e2e8f0"),
        paper_bgcolor="white",
        plot_bgcolor="white",
        height=420,
        legend=dict(orientation="v", x=1.02, y=0.5),
        margin=dict(t=60, b=40, l=60, r=20),
    )
    return fig


# ── Portfolio calculation engine (What-If aware) ─────────────────────────────────
def _compute_measure_financials(data: dict, dr=None, vsl_mult: float = 1.0) -> list:
    """
    Compute PV-based financials for each measure, respecting What-If DR and VSL overrides.
    Returns list[dict] with: name, capex, opex, life, pv_ben, pv_cost, npv, bcr,
    pv_health, pv_energy, pv_other.
    """
    orig_dr = data.get("discount_rate", 0.035) or 0.035
    eff_dr  = dr if (dr is not None and dr > 0) else orig_dr

    results = []
    for m in data.get("measures", []):
        life    = m.get("lifetime_years", 30) or 30
        capex   = m.get("capex", 0) or 0
        opex    = m.get("annual_opex", 0) or 0
        af      = (1 - (1 + eff_dr) ** (-life)) / eff_dr if eff_dr > 0 else life
        orig_af = (1 - (1 + orig_dr) ** (-life)) / orig_dr if orig_dr > 0 else life

        pv_cost   = capex + opex * af
        pv_health = pv_energy = pv_other = 0.0

        adv = m.get("advanced_benefits")
        if adv and isinstance(adv, dict):
            # Specialist: Claude's NPVs are at orig_dr → convert to annual → re-discount at eff_dr
            for key, val in adv.items():
                if not isinstance(val, (int, float)) or val <= 0:
                    continue
                annual = val / orig_af if orig_af > 0 else 0
                pv     = annual * af
                driver = BENEFIT_DRIVER_MAP.get(key, "static")
                if driver == "exponential":
                    pv_health += pv * vsl_mult
                elif driver == "linear":
                    pv_energy += pv
                else:
                    pv_other += pv
        else:
            for comp in (m.get("benefit_components") or []):
                val = comp.get("value", comp.get("annual_value", 0)) or 0
                if not isinstance(val, (int, float)) or val <= 0:
                    continue
                pv     = val * af
                driver = BENEFIT_DRIVER_MAP.get(comp.get("type", "generic_annual"), "static")
                if driver == "exponential":
                    pv_health += pv * vsl_mult
                elif driver == "linear":
                    pv_energy += pv
                else:
                    pv_other += pv

        pv_ben = pv_health + pv_energy + pv_other
        npv    = pv_ben - pv_cost
        bcr    = pv_ben / pv_cost if pv_cost > 0 else 0.0
        results.append({
            "name":      m.get("name", "Measure"),
            "capex":     capex,
            "opex":      opex,
            "life":      life,
            "pv_ben":    pv_ben,
            "pv_cost":   pv_cost,
            "npv":       npv,
            "bcr":       bcr,
            "pv_health": pv_health,
            "pv_energy": pv_energy,
            "pv_other":  pv_other,
        })
    return results


def _build_bcr_bar_chart(financials: list, currency: str = "") -> "go.Figure | None":
    """Horizontal bar chart comparing BCR across all measures with threshold lines."""
    if not financials:
        return None
    names  = [f["name"] for f in financials]
    bcrs   = [f["bcr"]  for f in financials]
    colors = ["#15803d" if b >= 1.5 else "#d97706" if b >= 1.0 else "#dc2626" for b in bcrs]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=names, x=bcrs,
        orientation="h",
        marker_color=colors,
        text=[f"{b:.2f}" for b in bcrs],
        textposition="outside",
        hovertemplate="%{y}<br>BCR: %{x:.2f}<extra></extra>",
    ))
    fig.add_vline(x=1.0, line_dash="dash", line_color="#64748b",
                  annotation_text="BCR=1.0 (break-even)",
                  annotation_position="top right")
    fig.add_vline(x=1.5, line_dash="dot", line_color="#15803d",
                  annotation_text="BCR=1.5 (recommended)",
                  annotation_position="top right")
    fig.update_layout(
        title=dict(text="Benefit-Cost Ratio by Measure",
                   font=dict(family="Inter, Segoe UI, sans-serif", color="#052e16", size=14), x=0.5),
        xaxis=dict(title="BCR", gridcolor="#e2e8f0"),
        yaxis=dict(title=""),
        paper_bgcolor="white", plot_bgcolor="white",
        height=max(280, 60 * len(financials) + 100),
        margin=dict(t=60, b=40, l=20, r=90),
        showlegend=False,
    )
    return fig


def _build_investment_map(financials: list, currency: str = "") -> "go.Figure | None":
    """Scatter plot: X=CAPEX, Y=NPV — identify low-hanging fruit (high NPV, low cost)."""
    if not financials:
        return None
    names  = [f["name"]  for f in financials]
    capexs = [f["capex"] for f in financials]
    npvs   = [f["npv"]   for f in financials]
    bcrs   = [f["bcr"]   for f in financials]
    sizes  = [max(12, min(f["pv_ben"] * 5, 60)) for f in financials]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=capexs, y=npvs,
        mode="markers+text",
        text=names,
        textposition="top center",
        marker=dict(
            size=sizes,
            color=bcrs,
            colorscale=[[0, "#dc2626"], [0.4, "#d97706"], [1, "#15803d"]],
            cmin=0, cmax=3,
            colorbar=dict(title="BCR", thickness=12),
            showscale=True,
            line=dict(width=1, color="white"),
        ),
        hovertemplate=(
            "<b>%{text}</b><br>"
            f"CAPEX: %{{x:.1f}} {currency}M<br>"
            f"NPV: %{{y:.1f}} {currency}M<br>"
            "BCR: %{marker.color:.2f}<extra></extra>"
        ),
    ))
    fig.add_hline(y=0, line_dash="dash", line_color="#64748b",
                  annotation_text="NPV=0", annotation_position="right")
    if len(financials) >= 2:
        fig.add_annotation(
            x=min(capexs) if capexs else 0,
            y=max(npvs) * 0.92 if any(v > 0 for v in npvs) else 1,
            text="🌟 Low-Hanging Fruit", showarrow=False,
            font=dict(color="#15803d", size=10),
        )
    fig.update_layout(
        title=dict(text=f"Investment Map — CAPEX vs NPV ({currency}M)",
                   font=dict(family="Inter, Segoe UI, sans-serif", color="#052e16", size=14), x=0.5),
        xaxis=dict(title=f"CAPEX ({currency}M)", gridcolor="#e2e8f0"),
        yaxis=dict(title=f"NPV ({currency}M)", gridcolor="#e2e8f0"),
        paper_bgcolor="white", plot_bgcolor="white",
        height=420,
        margin=dict(t=60, b=50, l=60, r=100),
    )
    return fig


def _build_waterfall_chart(financials: list, data: dict, currency: str = "") -> "go.Figure | None":
    """Waterfall chart: Initial Investment → O&M → Benefit categories → Net NPV."""
    if not financials:
        return None
    total_capex      = sum(f["capex"]             for f in financials)
    total_pv_opex    = sum(max(f["pv_cost"] - f["capex"], 0) for f in financials)
    total_pv_health  = sum(f["pv_health"]          for f in financials)
    total_pv_energy  = sum(f["pv_energy"]          for f in financials)
    total_pv_other   = sum(f["pv_other"]           for f in financials)
    total_npv        = sum(f["npv"]                for f in financials)

    x_labels = ["Initial\nInvestment", "Discounted\nO&M"]
    y_vals   = [-total_capex,           -total_pv_opex]
    m_types  = ["relative",             "relative"]

    if total_pv_health > 0:
        x_labels.append("Health\nBenefits (PV)")
        y_vals.append(total_pv_health)
        m_types.append("relative")
    if total_pv_energy > 0:
        x_labels.append("Energy\nBenefits (PV)")
        y_vals.append(total_pv_energy)
        m_types.append("relative")
    if total_pv_other > 0:
        x_labels.append("Other\nBenefits (PV)")
        y_vals.append(total_pv_other)
        m_types.append("relative")

    x_labels.append("Net NPV")
    y_vals.append(total_npv)
    m_types.append("total")

    fig = go.Figure(go.Waterfall(
        name="Portfolio",
        orientation="v",
        measure=m_types,
        x=x_labels,
        y=y_vals,
        connector=dict(line=dict(color="#94a3b8", width=1, dash="dot")),
        decreasing=dict(marker=dict(color="#dc2626")),
        increasing=dict(marker=dict(color="#15803d")),
        totals=dict(marker=dict(color="#052e16" if total_npv >= 0 else "#dc2626")),
        text=[f"{v:+,.1f}" for v in y_vals],
        textposition="outside",
        hovertemplate=f"%{{x}}<br>%{{y:+,.2f}} {currency}M<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text=f"Portfolio Cash-Flow Waterfall ({currency}M)",
                   font=dict(family="Inter, Segoe UI, sans-serif", color="#052e16", size=14), x=0.5),
        yaxis=dict(title=f"{currency}M", gridcolor="#e2e8f0"),
        paper_bgcolor="white", plot_bgcolor="white",
        height=420,
        margin=dict(t=60, b=60, l=60, r=40),
        showlegend=False,
    )
    return fig


# ── Methodology auditor ──────────────────────────────────────────────────────────
_SOURCE_GRADE_A = [
    "oecd", "world bank", "ims", "israeli", "ministry of finance",
    "cbs", "bank of israel", "ipcc", "who ", "unep", "bls",
    "ministry of health", "government", "official",
]
_SOURCE_GRADE_B = [
    "lancet", "nature ", "journal", "jama", "bmj", "environ",
    "science ", "et al", "pubmed", "peer-review", "wiley", "springer",
    "elsevier", "plos", "epidemiology", "health economics",
]
_MISSING_SOURCE_VALUES = {"", "unknown", "default", "user provided", "literature estimate", "n/a"}

def _grade_source(source_text: str) -> tuple:
    """Grade a source citation. Returns (grade, label): A=Official, B=Peer-Reviewed, C=Grey, ?=Missing."""
    if not source_text or source_text.strip().lower() in _MISSING_SOURCE_VALUES:
        return ("?", "Missing")
    sl = source_text.lower()
    if any(k in sl for k in _SOURCE_GRADE_A):
        return ("A", "Official / Institutional")
    if any(k in sl for k in _SOURCE_GRADE_B):
        return ("B", "Peer-Reviewed")
    return ("C", "Grey Literature")


def _apply_audit_corrections(data: dict) -> dict:
    """Return a deep-corrected copy of analysis_data with ERROR-level params fixed."""
    d = copy.deepcopy(data)
    for m in d.get("measures", []):
        for comp in (m.get("benefit_components") or []):
            # VSL: if looks like millions (e.g. 3.8) convert to full units
            vsl = comp.get("vsl", 0) or 0
            if isinstance(vsl, (int, float)) and 0 < vsl < 1000:
                comp["vsl"] = int(vsl * 1_000_000)
            # HMF: cap at 0.10 → recommended 0.035 (Mediterranean baseline)
            hmf = comp.get("heat_mortality_factor", 0)
            if isinstance(hmf, (int, float)) and hmf > 0.10:
                comp["heat_mortality_factor"] = 0.035
                comp["heat_mortality_factor_source"] = (
                    (comp.get("heat_mortality_factor_source") or "") +
                    f" [CORRECTED from {hmf:.3f} to 0.035 — Gasparrini 2017 Mediterranean median]"
                )
            # Efficiency: cap at 0.85 → 0.77 (Bouchama 2007)
            eff = comp.get("heat_reduction_efficiency", 0)
            if isinstance(eff, (int, float)) and eff > 0.85:
                comp["heat_reduction_efficiency"] = 0.77
            # Uplift fraction: if given as percentage not decimal
            uplift = comp.get("uplift_fraction", 0)
            if isinstance(uplift, (int, float)) and uplift > 1.0:
                comp["uplift_fraction"] = round(uplift / 100, 4)
    return d


def _run_methodology_audit(data: dict, challenge_type: str = "general") -> list:
    """
    UI-facing parameter sanity checks. Runs before Excel build.
    Returns list of {"level":"ERROR"|"WARNING","measure":str,"component":str,"issue":str}.
    Complements _sanitise_mortality_params() in excel_builder which auto-corrects silently.
    """
    findings = []
    for m in data.get("measures", []):
        mname = m.get("name", "Unknown Measure")
        for comp in (m.get("benefit_components") or []):
            cname = comp.get("name", "component")
            ctype = comp.get("type", "")

            if ctype == "avoided_mortality":
                vsl = comp.get("vsl", 0)
                if isinstance(vsl, (int, float)) and 0 < vsl < 100_000:
                    findings.append({"level": "ERROR", "measure": mname, "component": cname,
                        "issue": f"VSL = {vsl} — appears to be in millions, not full units. "
                                 "Expected e.g. 3,800,000. Excel will auto-correct."})

                hmf = comp.get("heat_mortality_factor", 0)
                if isinstance(hmf, (int, float)) and challenge_type != "flood":
                    if hmf > 0.5:
                        findings.append({"level": "ERROR", "measure": mname, "component": cname,
                            "issue": f"heat_mortality_factor = {hmf:.3f} — this means {hmf*100:.0f}% of all deaths are heat-related "
                                     f"(max is ~10%). Likely a prompt artefact. Click 'Apply Corrections' to set → 0.035 "
                                     "(Gasparrini 2017 Mediterranean median)."})
                    elif hmf > 0.128:
                        findings.append({"level": "WARNING", "measure": mname, "component": cname,
                            "issue": f"heat_mortality_factor = {hmf:.3f} ({hmf*100:.1f}%) — exceeds 12.8% "
                                     "upper extreme. Gasparrini 2017 Mediterranean median ≈ 3.5%. "
                                     "Excel will auto-correct to 0.067."})

                eff = comp.get("heat_reduction_efficiency", 0)
                if isinstance(eff, (int, float)) and eff > 1.0:
                    findings.append({"level": "ERROR", "measure": mname, "component": cname,
                        "issue": f"heat_reduction_efficiency = {eff} — must be a fraction 0.0–1.0, not a percentage."})

                mr = comp.get("mortality_rate", 0)
                if isinstance(mr, (int, float)) and mr > 0.1:
                    findings.append({"level": "WARNING", "measure": mname, "component": cname,
                        "issue": f"mortality_rate = {mr:.3f} ({mr*100:.1f}%/yr) — exceeds 10%. "
                                 "Typical all-cause rate age 65-74 ≈ 1.2% (0.012)."})

                pop = comp.get("population_at_risk", 0) or 0
                if challenge_type != "flood":
                    hmf2 = min(hmf if isinstance(hmf, (int, float)) else 0, 0.10)
                    eff2 = min(eff if isinstance(eff, (int, float)) else 0, 1.0)
                    deaths = pop * (mr if isinstance(mr, (int, float)) else 0) * hmf2 * eff2
                    if deaths > 500:
                        findings.append({"level": "WARNING", "measure": mname, "component": cname,
                            "issue": f"Implied deaths_avoided ≈ {deaths:.0f}/yr — exceeds plausible range "
                                     "(typically 1–200/yr per measure). "
                                     "Verify population_at_risk is only the sub-group served by this measure."})

            if ctype == "property_value_uplift":
                uplift = comp.get("uplift_fraction", 0)
                if isinstance(uplift, (int, float)):
                    if uplift > 1.0:
                        findings.append({"level": "ERROR", "measure": mname, "component": cname,
                            "issue": f"uplift_fraction = {uplift} — must be a decimal (0.03 = 3%), "
                                     "not a percentage. Excel will auto-correct by dividing by 100."})
                    elif uplift > 0.5:
                        findings.append({"level": "WARNING", "measure": mname, "component": cname,
                            "issue": f"uplift_fraction = {uplift:.2f} — implies {uplift*100:.0f}% property value uplift. "
                                     "Literature range 2–8% (Fuerst & McAllister 2011)."})

        # ── Check Group A: Source completeness ──────────────────────────────
        _AM_SOURCES_HEAT  = ["vsl_source", "mortality_rate_source", "heat_mortality_factor_source"]
        _AM_SOURCES_FLOOD = ["vsl_source", "mortality_rate_source"]
        _KEY_SOURCE_FIELDS = {
            "avoided_mortality":    _AM_SOURCES_FLOOD if challenge_type == "flood" else _AM_SOURCES_HEAT,
            "energy_savings":       ["electricity_tariff_source"],
            "morbidity_savings":    ["hospitalization_cost_source"],
            "property_value_uplift":["property_value_per_m2_source", "uplift_fraction_source"],
        }
        for comp in (m.get("benefit_components") or []):
            ctype_a = comp.get("type", "generic_annual")
            cname_a = comp.get("name", "")
            for src_field in _KEY_SOURCE_FIELDS.get(ctype_a, []):
                grade, label = _grade_source(comp.get(src_field, ""))
                if grade == "?":
                    findings.append({"level": "WARNING", "measure": mname, "component": cname_a,
                        "issue": f"Missing citation for `{src_field}` — mark 'Low Confidence'"})
                elif grade == "C":
                    findings.append({"level": "INFO", "measure": mname, "component": cname_a,
                        "issue": f"`{src_field}` uses grey literature ({label}) — prefer Grade A/B source"})

        # ── Check Group B: VSL outlier detection (upper bounds) ─────────────
        for comp in (m.get("benefit_components") or []):
            if comp.get("type") == "avoided_mortality":
                vsl = comp.get("vsl", 0) or 0
                cname_b = comp.get("name", "")
                if vsl > 100_000_000:
                    findings.append({"level": "ERROR", "measure": mname, "component": cname_b,
                        "issue": f"VSL {vsl:,.0f} exceeds 100M — likely unit error (should be ~3M–20M in full currency units)"})
                elif vsl > 30_000_000:
                    findings.append({"level": "WARNING", "measure": mname, "component": cname_b,
                        "issue": f"VSL {vsl:,.0f} is >{vsl/16_600_000:.1f}× the NIS benchmark (~16.6M) — verify source"})

        # ── Check Group C: Benefit-transfer audit ───────────────────────────
        _FOREIGN_KW  = ["oecd", "us epa", "european", "eu ", "uk ", "lancet", "who"]
        _LOCAL_ADJ   = ["ppp", "adjusted", "elasticity", "transfer", "nis", "israel", "shekel"]
        for comp in (m.get("benefit_components") or []):
            cname_c = comp.get("name", "")
            all_sources = " ".join(
                str(v) for k, v in comp.items() if k.endswith("_source") and v
            ).lower()
            has_foreign = any(kw in all_sources for kw in _FOREIGN_KW)
            has_adj     = any(kw in all_sources for kw in _LOCAL_ADJ)
            if has_foreign and not has_adj:
                findings.append({"level": "WARNING", "measure": mname, "component": cname_c,
                    "issue": "Foreign source detected — confirm PPP / income-elasticity adjustment applied for Israeli context"})

        # ── Check Group D: Population ethics & reasonableness ───────────────
        for comp in (m.get("benefit_components") or []):
            if comp.get("type") == "avoided_mortality":
                cname_d = comp.get("name", "")
                pop = comp.get("population_at_risk", 0) or 0
                if pop > 5_000_000:
                    findings.append({"level": "WARNING", "measure": mname, "component": cname_d,
                        "issue": f"population_at_risk = {pop:,} exceeds 5M — confirm this is a vulnerable sub-group, not the total city population"})
                mr_d = comp.get("mortality_rate", 0) or 0
                if isinstance(mr_d, (int, float)) and 0 < mr_d < 0.001:
                    if challenge_type == "flood":
                        findings.append({"level": "WARNING", "measure": mname, "component": cname_d,
                            "issue": f"mortality_rate = {mr_d} is very low for flood mortality — "
                                     "use general population rate ~0.008 (CBS 2022 Life Tables, age-standardised)"})
                    else:
                        findings.append({"level": "WARNING", "measure": mname, "component": cname_d,
                            "issue": f"mortality_rate = {mr_d} is very low (<0.1%) — verify demographic group (elderly: 1–4.5%, general: 0.8%)"})

        # ── Check Group E: Structural integrity (orphan types) ──────────────
        _HANDLED_TYPES = {"avoided_mortality", "energy_savings", "morbidity_savings",
                          "property_value_uplift", "generic_annual"}
        for comp in (m.get("benefit_components") or []):
            ctype_e = comp.get("type", "generic_annual")
            if ctype_e not in _HANDLED_TYPES:
                findings.append({"level": "WARNING", "measure": mname, "component": comp.get("name", ""),
                    "issue": f"Unrecognized benefit type `{ctype_e}` — this component will NOT appear in NPV (golden thread broken)"})

    return findings


def _render_validation_report(findings: list, data: dict = None) -> bool:
    """Render methodology audit findings in the UI. Returns True if any ERROR-level finding (blocks download)."""
    has_errors   = any(f["level"] == "ERROR"   for f in findings)
    has_warnings = any(f["level"] == "WARNING" for f in findings)
    has_info     = any(f["level"] == "INFO"    for f in findings)

    # ── Corrections-applied banner ───────────────────────────────────────────
    if st.session_state.get("audit_corrections_applied"):
        st.success("✅ Recommended corrections were applied — parameters updated to literature benchmarks.")

    # ── Source confidence tally ──────────────────────────────────────────────
    grade_counts: dict = {"A": 0, "B": 0, "C": 0, "?": 0}
    for m in (data.get("measures", []) if data else []):
        for comp in (m.get("benefit_components") or []):
            for k, v in comp.items():
                if k.endswith("_source"):
                    g, _ = _grade_source(str(v) if v else "")
                    grade_counts[g] = grade_counts.get(g, 0) + 1

    # ── Header banner ────────────────────────────────────────────────────────
    if not findings:
        st.success("✅ All checks passed — parameters within expected ranges.")
        with st.expander("📋 Methodology Audit Report", expanded=False):
            st.markdown(
                f"📚 **Source Confidence:** "
                f"**{grade_counts['A']}×** Grade A (Official)  ·  "
                f"**{grade_counts['B']}×** Grade B (Peer-Reviewed)  ·  "
                f"**{grade_counts['C']}×** Grade C (Grey)  ·  "
                f"**{grade_counts['?']}×** Missing"
            )
            st.markdown("*No issues found.*")
        return False

    n_err  = sum(1 for f in findings if f["level"] == "ERROR")
    n_warn = sum(1 for f in findings if f["level"] == "WARNING")

    if has_errors:
        st.error(f"❌ Methodology check: **{n_err} error(s)**, {n_warn} warning(s). "
                 "Review below. Excel will attempt auto-correction but download is blocked until confirmed.")
    elif has_warnings:
        st.warning(f"⚠️ Methodology check: **{n_warn} warning(s)** — review recommended. You may still download.")
    elif has_info:
        st.info("ℹ️ Methodology check: minor notes — source quality advisory only. Download allowed.")

    with st.expander("📋 Methodology Audit Report", expanded=has_errors):
        # Source confidence summary
        st.markdown(
            f"📚 **Source Confidence:** "
            f"**{grade_counts['A']}×** Grade A (Official)  ·  "
            f"**{grade_counts['B']}×** Grade B (Peer-Reviewed)  ·  "
            f"**{grade_counts['C']}×** Grade C (Grey)  ·  "
            f"**{grade_counts['?']}×** Missing"
        )
        st.markdown("---")
        lines = []
        for f in findings:
            if f["level"] == "ERROR":
                lines.append(f"**:red[❌ ERROR]** &nbsp; `{f['measure']}` / `{f['component']}` → {f['issue']}")
            elif f["level"] == "WARNING":
                lines.append(f"**:orange[⚠️ WARNING]** &nbsp; `{f['measure']}` / `{f['component']}` → {f['issue']}")
            elif f["level"] == "INFO":
                lines.append(f"**:blue[ℹ️ INFO]** &nbsp; `{f['measure']}` / `{f['component']}` → {f['issue']}")
        st.markdown("\n\n".join(lines))

    if has_errors:
        st.markdown("---")
        _ac1, _ac2 = st.columns(2)
        with _ac1:
            if st.button("🔧 Apply Recommended Corrections & Regenerate",
                         key="btn_apply_corrections", type="primary",
                         help="Auto-corrects: VSL units, heat_mortality_factor >10%, uplift fractions"):
                corrected = _apply_audit_corrections(data)
                st.session_state.analysis_data = corrected
                st.session_state.audit_corrections_applied = True
                st.session_state.audit_acknowledged = False
                st.rerun()
        with _ac2:
            if st.button("⚠️ Acknowledge & Download Anyway",
                         key="btn_acknowledge_audit",
                         help="Download the Excel model with a warning — unverified parameters flagged"):
                st.session_state.audit_acknowledged = True
                st.rerun()

    return has_errors


# ── Specialist prompt constants ─────────────────────────────────────────────────
# Use USER_INPUT_PLACEHOLDER as a safe substitution token (avoids f-string brace conflicts)

GENERIC_DATA_PROMPT = """User provided: "USER_INPUT_PLACEHOLDER"

OUTPUT ONLY valid JSON. No text before or after. Start with { and end with }.

ARCHITECTURE RULE: Every benefit component MUST use a typed formula — not a single pre-computed number.
Use type "avoided_mortality", "energy_savings", "morbidity_savings", or "property_value_uplift".
Use "generic_annual" ONLY as a last resort when none of the above apply.

═══════════════════════════════════════════════════════════════
AVOIDED MORTALITY — PARAMETER GUIDE (read carefully)
═══════════════════════════════════════════════════════════════
Formula: Annual Benefit (EUR M) = population_at_risk × mortality_rate × heat_mortality_factor × heat_reduction_efficiency × vsl ÷ 1,000,000

PARAMETER DEFINITIONS AND CORRECT RANGES:
• population_at_risk      — number of PEOPLE in the vulnerable group (e.g. 30000 elderly residents)
• mortality_rate          — annual ALL-CAUSE death rate per person, e.g. 0.012 for age 65-74 (1.2% die per year)
• heat_mortality_factor   — fraction of ALL annual deaths that are attributable to heat (NOT a per-person rate)
                            Literature range: 0.02–0.10. Typical: 0.05–0.07 for temperate/Mediterranean cities.
                            Source: Gasparrini 2017 Lancet gives ~2–5% for Mediterranean; Hyyrynen 2025 uses 6.7%.
                            DO NOT use 0.128 (12.8%) — this is the upper bound for extreme heat climates only.
• heat_reduction_efficiency — fraction of heat deaths prevented by this intervention (0.0–1.0)
                            Typical: 0.28–0.85 depending on measure. AC in homes: ~0.77. DC network: ~0.50.
• vsl                     — Value of Statistical Life in FULL currency units (not millions!)
                            e.g. 3800000 for EUR 3.8M. NEVER write 3.8.

SANITY CHECK — COMPUTE THIS BEFORE SUBMITTING:
  deaths_avoided = population_at_risk × mortality_rate × heat_mortality_factor × efficiency
  This number should be REALISTIC: typically 1–200 deaths/year per measure for a city.
  If deaths_avoided > 500/year for ONE measure → your parameters are too high, reduce them.

WORKED EXAMPLE (realistic for Tel Aviv, 30,000 elderly 75+):
  deaths_avoided = 30000 × 0.045 × 0.035 × 0.50 = 23.6 deaths/yr  ✓ plausible
  annual_benefit = 23.6 × 3,800,000 / 1,000,000 = EUR 89.7M/yr
  (Note: heat_mortality_factor = 0.035 from Gasparrini 2017 Mediterranean median, NOT 0.128)

CRITICAL DEFAULT GUARD — ISRAEL BASELINES:
- If user says "use defaults" for Israel/Tel Aviv:
  mortality_rate = 0.008 (CBS 2022 general population) or 0.012 (age 65-74)
  heat_mortality_factor = 0.035 (Gasparrini 2017 Mediterranean cluster median)
- NEVER output heat_mortality_factor = 1.0 or any value above 0.10.
  A value of 1.0 would mean 100% of all annual deaths are from heat — physically impossible.
  If you are uncertain, default to 0.035.
═══════════════════════════════════════════════════════════════

For "energy_savings" provide: area_m2, energy_reduction_kwh_m2, electricity_tariff
For "morbidity_savings" provide: cases_avoided_per_year, hospitalization_cost, avg_length_of_stay_days
For "property_value_uplift" provide: affected_area_m2, property_value_per_m2, uplift_fraction

Every numeric parameter needs a "_source" field: e.g. "vsl_source": "OECD 2012 ENV/WKP median"
Use literature defaults for any parameter the user did not provide.

{
  "problem_title": "...",
  "problem_summary": "2-sentence summary",
  "discount_rate": 0.035,
  "time_horizon": 30,
  "currency": "EUR",
  "currency_unit": "millions",
  "specialist_type": null,
  "measures": [
    {
      "name": "Measure name",
      "description": "What it does",
      "category": "Infrastructure|Nature-based|Policy|Technology",
      "capex": 0.0,
      "capex_source": "User provided / Author Year",
      "annual_opex": 0.0,
      "opex_source": "User provided / Author Year",
      "cost_breakdown": [
        {"type": "capex", "item": "Site preparation", "unit_cost": 5000, "qty": 1, "unit": "lump sum", "note": "optional — include if known"},
        {"type": "opex",  "item": "Annual maintenance", "unit_cost": 10000, "qty": 1, "unit": "yr", "note": "optional — include if known"}
      ],
      "co_benefits": "Non-monetised benefits",
      "lifetime_years": 30,
      "feasibility": "High|Medium|Low",
      "uncertainty": "Low|Medium|High",
      "benefit_components": [
        {
          "name": "Avoided Heat Mortality (Age 65-74)",
          "type": "avoided_mortality",
          "population_at_risk": 30000,
          "population_at_risk_source": "User provided / CBS estimate",
          "mortality_rate": 0.012,
          "mortality_rate_source": "Israel CBS Life Tables 2022, age 65-74",
          "heat_mortality_factor": 0.035,
          "heat_mortality_factor_source": "Gasparrini et al. 2017 Lancet, Table 3 Mediterranean median",
          "heat_reduction_efficiency": 0.50,
          "heat_reduction_efficiency_source": "WHO 2013 / literature estimate for this measure type",
          "vsl": 3800000,
          "vsl_source": "OECD 2012 ENV/WKP meta-analysis EUR median"
        }
      ]
    }
  ],
  "sensitivity_vars": [
    {"name": "Discount Rate",              "base": 0.035, "low": 0.015, "high": 0.07,  "unit": "%"},
    {"name": "VSL",                        "base": 1.0,   "low": 0.5,   "high": 1.5,   "unit": "multiplier"},
    {"name": "Heat Mortality Factor",      "base": 1.0,   "low": 0.5,   "high": 1.5,   "unit": "multiplier"},
    {"name": "CAPEX Variation",            "base": 1.0,   "low": 0.8,   "high": 1.3,   "unit": "multiplier"},
    {"name": "Vulnerable Population",      "base": 1.0,   "low": 0.7,   "high": 1.3,   "unit": "multiplier"},
    {"name": "Electricity Tariff",         "base": 1.0,   "low": 0.7,   "high": 1.5,   "unit": "multiplier"}
  ],
  "key_assumptions": [
    {"text": "assumption text", "source": "Author Year, Table/Page"}
  ],
  "data_gaps": "What local data would most improve accuracy"
}

CRITICAL UNITS:
- "vsl" MUST be in full currency units: write 3800000 (not 3.8 or 3,800,000)
- "capex" and "annual_opex" are in currency millions (e.g. 10.5 = EUR 10.5M)
- "population_at_risk", "area_m2", "hospitalization_cost" are in natural units (full numbers)
- "mortality_rate", "heat_mortality_factor", "heat_reduction_efficiency", "uplift_fraction" are fractions (0.0 to 1.0)"""

NATURAL_SHADING_DATA_PROMPT = """User provided: "USER_INPUT_PLACEHOLDER"

You are a specialist in climate adaptation economics with expertise in urban heat island mitigation and nature-based solutions. The user is analyzing a natural shading / boulevard tree / urban canopy intervention. Apply the following peer-reviewed methodology precisely.

METHODOLOGY PARAMETERS (use these exact values unless user provides better local data):
- Functional unit: 1 linear meter of shaded boulevard
- Time horizon: 50 years (system boundary for all costs and benefits)
- Discount rate: 3.5% (unless user specifies otherwise)
- Currency: NIS (New Israeli Shekel) — use 3.7 NIS/USD exchange rate

INTERVENTION COST & BENEFIT STRUCTURE:
- Costs: initial CAPEX (planting, pit preparation, irrigation systems) + annual OPEX (maintenance, pruning, water) + any replacement costs within the 50-year horizon.
- Direct benefits: thermal comfort and reduced heat stress for pedestrians, UV radiation shading and skin cancer risk reduction.
- Indirect ecosystem service benefits: carbon sequestration, urban runoff reduction (avoided water treatment and drainage costs), improved air quality, and habitat creation — all monetised over the 50-year horizon.

VSL ENGINE (Value of Statistical Life):
1. Base VSL: $3,000,000 (2005 USD, OECD baseline)
2. CPI adjustment 2005→2024: multiply by 1.68 (US CPI ratio)
3. GDP PPP adjustment: multiply by ratio of Israel GDP per capita PPP to OECD average. Use ~0.89 if no user data.
4. Income elasticity: 1.0 (standard for developed economies)
5. Convert to NIS: multiply by 3.7
6. Computed VSL ≈ 12,800,000 NIS
7. VSLY = VSL / remaining life expectancy of affected demographic (default: 35 years). VSLY ≈ 365,714 NIS/life-year.

CDD PARAMETERS (Cooling Degree Days, Tel Aviv baseline):
- Annual CDD: 735 (21°C base temperature)
- Heat-mortality factor: 0.00083 (deaths per person per CDD above threshold)
- As climate worsens and CDDs increase, the avoided-damage value of the intervention rises, increasing the BCR.

BENEFIT CALCULATION RULES:
1. Avoided Mortality: vulnerable_population × base_mortality_rate × heat_mortality_factor × heat_reduction_efficiency(0.50) × maturity_factor(year) × VSL — sum over 50 years discounted.
2. Morbidity Savings: hospitalization_cost(3,928 NIS/day) × avg_length_stay(5.2 days) × heat_attributable_cases_avoided × efficiency × maturity_factor.
3. Skin Cancer Prevention: pedestrians_per_hour × operating_hours(8) × UV_reduction(0.75) × skin_cancer_incidence_rate × (avg_treatment_cost + VSLY_loss) × maturity_factor.
4. Ecosystem Services (use literature estimates if no user data):
   - Carbon sequestration: 450 NIS/tree/year × tree_density
   - Runoff reduction: infrastructure_cost_avoided × runoff_coefficient_improvement
   - Air quality: PM2.5 reduction × health_cost_per_unit
   - Habitat creation: 200–500 NIS/m²/year biodiversity value

MATURITY CURVE — CRITICAL:
Benefits ramp up linearly over 8 years (biological maturity of trees):
- Years 1–8: maturity_factor = year_number / 8 (e.g. year 1 = 12.5%, year 4 = 50%, year 8 = 100%)
- Years 9–50: maturity_factor = 1.0
Apply this factor to ALL benefit streams every year.

CAPEX AND OPEX GUIDANCE:
- Literature range: 800–2,500 NIS/linear meter for planting
- Annual maintenance: 150–400 NIS/linear meter/year

Compute each advanced_benefits NPV over 50 years at the stated discount rate, applying the maturity ramp. Express all monetary values in NIS millions.

Respond ONLY with valid JSON, no markdown fences:

{
  "problem_title": "short descriptive title",
  "problem_summary": "2-sentence summary of the shading intervention and context",
  "discount_rate": 0.035,
  "time_horizon": 50,
  "currency": "NIS",
  "currency_unit": "millions",
  "specialist_type": "natural_shading",
  "vsl_params": {
    "base_vsl_usd_2005": 3000000,
    "cpi_multiplier": 1.68,
    "gdp_ppp_ratio": 0.89,
    "income_elasticity": 1.0,
    "usd_to_local_currency": 3.7,
    "life_expectancy_remaining": 35,
    "computed_vsl_local": 12800000,
    "computed_vsly_local": 365714
  },
  "cdd_params": {
    "annual_cdd": 735,
    "base_temp_celsius": 21,
    "heat_mortality_factor": 0.00083,
    "population_density_or_pedestrians": 1200
  },
  "specialist_params": {
    "maturity_years": 8,
    "heat_reduction_efficiency": 0.50,
    "uv_reduction_factor": 0.75,
    "pedestrians_per_hour": 1200,
    "functional_unit": "linear meter of shaded boulevard"
  },
  "measures": [
    {
      "name": "Natural Shading Boulevard",
      "description": "What the intervention involves",
      "category": "Nature-based",
      "capex": 1.5,
      "annual_opex": 0.15,
      "benefit_types": ["avoided mortality", "morbidity savings", "skin cancer prevention", "carbon sequestration", "runoff reduction", "air quality", "habitat creation"],
      "co_benefits": "Urban heat island reduction, aesthetic value, pedestrian comfort",
      "lifetime_years": 50,
      "feasibility": "High",
      "uncertainty": "Medium",
      "data_source": "Specialist methodology — VSL/CDD model",
      "advanced_benefits": {
        "avoided_mortality_npv": 0.0,
        "morbidity_savings_npv": 0.0,
        "skin_cancer_prevention_npv": 0.0,
        "carbon_sequestration_npv": 0.0,
        "runoff_reduction_npv": 0.0,
        "air_quality_npv": 0.0,
        "habitat_creation_npv": 0.0
      }
    }
  ],
  "sensitivity_vars": [],
  "key_assumptions": [
    {"text": "50-year time horizon with 1 linear meter of shaded boulevard as functional unit", "source": "Methodology default — standard urban CBA horizon"},
    {"text": "VSL from OECD 2005 baseline adjusted to 2024 Israel via CPI (×1.68) and PPP (×0.89), converted at 3.7 NIS/USD", "source": "OECD (2005); US BLS CPI; World Bank PPP data"},
    {"text": "CDD baseline 735 (Tel Aviv, 21°C base), heat-mortality factor 0.00083", "source": "Israeli Meteorological Service; Gasparrini et al. (2017)"},
    {"text": "8-year linear biological maturity ramp applied to all benefit streams (year n ÷ 8)", "source": "Nowak et al. (2002), urban tree growth literature"},
    {"text": "Heat reduction efficiency fixed at 50%; UV reduction factor 0.75", "source": "Shashua-Bar & Hoffman (2000); WHO UV Index guidelines"},
    {"text": "Ecosystem service values (carbon 450 NIS/tree/yr, habitat 200-500 NIS/m²/yr) from literature", "source": "TEEB (2010); Israeli carbon market reference prices"}
  ],
  "data_gaps": "Localized mortality and morbidity rates, precise pedestrians-per-hour counts, city-specific CAPEX/OPEX cost data, and locally derived ecosystem service valuation coefficients."
}

Fill in all numeric fields based on user data and the methodology above. If user data is unavailable, use methodology defaults and note in data_source."""

GREEN_ROOF_DATA_PROMPT = """User provided: "USER_INPUT_PLACEHOLDER"

You are a specialist in climate adaptation economics with expertise in urban green infrastructure and rooftop ecology. The user is analyzing a green roof / rooftop vegetation intervention. Apply the following peer-reviewed methodology precisely.

METHODOLOGY PARAMETERS (use these exact values unless user provides better local data):
- Functional unit: 1 sq meter of green roof
- Time horizon: 50 years
- Discount rate: 3.5% (unless user specifies otherwise)
- Currency: NIS (New Israeli Shekel) — use 3.7 NIS/USD exchange rate

VSL ENGINE (Value of Statistical Life):
1. Base VSL: $3,000,000 (2005 USD, OECD baseline)
2. CPI adjustment 2005→2024: multiply by 1.68
3. GDP PPP adjustment: multiply by ~0.89 (Israel vs OECD average)
4. Income elasticity: 1.0
5. Convert to NIS: multiply by 3.7
6. Computed VSL ≈ 12,800,000 NIS
7. VSLY = VSL / remaining_life_expectancy (default: 35 years) ≈ 365,714 NIS/life-year

CDD PARAMETERS:
- Annual CDD: 735 (21°C base, Tel Aviv)
- Heat-mortality factor: 0.00083
- Green roof reduces ambient temperature in building and immediate surroundings

POPULATION DRIVER:
- Residential density: 19,000 people/km² (typical Israeli urban residential)
- Scale to building catchment area for mortality and morbidity calculations

BENEFIT CALCULATION RULES:
1. Avoided Mortality: catchment_population × base_mortality_rate × heat_mortality_factor × heat_reduction_efficiency(0.28) × VSL — sum over 50 years discounted.
2. Morbidity Savings: hospitalization_cost(3,928 NIS/day) × avg_length_stay(5.2 days) × heat_attributable_cases_avoided × heat_reduction_efficiency(0.28).
3. Property Value Uplift (capital benefit, Year 1 only): roof_area_m2 × property_value_per_m2 × property_value_uplift_pct(0.03). One-time benefit, discounted to Year 1; represents a direct increase in real estate value.
4. Roof Longevity Extension: (roof_replacement_cost / conventional_roof_lifetime) × longevity_extension_years(15). Lump-sum avoided replacement cost at the conventional roof end-of-life.
5. Ecosystem Services (NIS/year):
   - Carbon sequestration: 350 NIS/m²/year × green_roof_area
   - Runoff reduction: stormwater_infrastructure_cost_avoided × runoff_reduction_coefficient(0.65)
   - Air quality: PM2.5 reduction × health_cost_per_unit
   - Habitat creation: 300 NIS/m²/year × green_roof_area

NO MATURITY RAMP for green roofs. Benefits are immediate at full capacity from Year 1.

CAPEX AND OPEX GUIDANCE:
- Extensive green roof: 300–600 NIS/m²
- Intensive green roof: 800–2,000 NIS/m²
- Annual maintenance: 50–120 NIS/m²/year

Compute each advanced_benefits NPV over 50 years at the stated discount rate. Express all monetary values in NIS millions.

Respond ONLY with valid JSON, no markdown fences:

{
  "problem_title": "short descriptive title",
  "problem_summary": "2-sentence summary of the green roof intervention and context",
  "discount_rate": 0.035,
  "time_horizon": 50,
  "currency": "NIS",
  "currency_unit": "millions",
  "specialist_type": "green_roof",
  "vsl_params": {
    "base_vsl_usd_2005": 3000000,
    "cpi_multiplier": 1.68,
    "gdp_ppp_ratio": 0.89,
    "income_elasticity": 1.0,
    "usd_to_local_currency": 3.7,
    "life_expectancy_remaining": 35,
    "computed_vsl_local": 12800000,
    "computed_vsly_local": 365714
  },
  "cdd_params": {
    "annual_cdd": 735,
    "base_temp_celsius": 21,
    "heat_mortality_factor": 0.00083,
    "population_density_or_pedestrians": 19000
  },
  "specialist_params": {
    "heat_reduction_efficiency": 0.28,
    "property_value_uplift_pct": 0.03,
    "roof_longevity_extension_years": 15,
    "roof_area_m2": 1000,
    "functional_unit": "sq meter of green roof"
  },
  "measures": [
    {
      "name": "Green Roof Installation",
      "description": "What the intervention involves",
      "category": "Nature-based",
      "capex": 0.5,
      "annual_opex": 0.06,
      "benefit_types": ["avoided mortality", "morbidity savings", "property value uplift", "roof longevity", "carbon sequestration", "runoff reduction", "air quality", "habitat creation"],
      "co_benefits": "Urban heat island reduction, stormwater management, building insulation",
      "lifetime_years": 50,
      "feasibility": "Medium",
      "uncertainty": "Medium",
      "data_source": "Specialist methodology — VSL/CDD model",
      "advanced_benefits": {
        "avoided_mortality_npv": 0.0,
        "morbidity_savings_npv": 0.0,
        "property_value_uplift_npv": 0.0,
        "roof_longevity_npv": 0.0,
        "carbon_sequestration_npv": 0.0,
        "runoff_reduction_npv": 0.0,
        "air_quality_npv": 0.0,
        "habitat_creation_npv": 0.0
      }
    }
  ],
  "sensitivity_vars": [],
  "key_assumptions": [
    {"text": "50-year time horizon with 1 m² of green roof as functional unit", "source": "Methodology default — standard green infrastructure CBA horizon"},
    {"text": "VSL from OECD 2005 baseline adjusted to 2024 Israel via CPI (×1.68) and PPP (×0.89), converted at 3.7 NIS/USD", "source": "OECD (2005); US BLS CPI; World Bank PPP data"},
    {"text": "CDD baseline 735 (Tel Aviv, 21°C base), heat-mortality factor 0.00083", "source": "Israeli Meteorological Service; Gasparrini et al. (2017)"},
    {"text": "No maturity ramp — green roof benefits operate at full capacity from Year 1", "source": "Berghage et al. (2009); green roof engineering literature"},
    {"text": "Heat reduction efficiency 28%; property value uplift 3% of roof area value", "source": "Oberndorfer et al. (2007); Fuerst & McAllister (2011)"},
    {"text": "Roof longevity extension 15 years over conventional roof; ecosystem services (carbon 350 NIS/m²/yr, habitat 300 NIS/m²/yr)", "source": "Berghage et al. (2009); TEEB (2010)"}
  ],
  "data_gaps": "Local building-level property values and turnover rates, detailed residential density and exposure, city-specific CAPEX/OPEX for green roofs, and refined local ecosystem service valuation coefficients."
}

Fill in all numeric fields based on user data and the methodology above. If user data is unavailable, use methodology defaults and note in data_source."""


# ── Session state init ──────────────────────────────────────────────────────────
# AWS/Bedrock credentials are loaded from Streamlit secrets (server-side, free for all users).
# Anthropic API key is always entered by the user — each person uses their own account.
def _secret(key, default=""):
    try:
        return st.secrets[key]
    except Exception:
        return default

for k, v in {
    "messages": [], "stage": "problem", "analysis_data": None,
    "api_key":    "",                          # always user-provided
    "problem_text": "", "specialist_type": None,
    "challenge_type": "general",
    "custom_measure_context": None,
    "uploaded_file_text": "",
    "kb_id":      _secret("BEDROCK_KB_ID"),    # server secret
    "aws_key":    _secret("AWS_ACCESS_KEY_ID"),
    "aws_secret": _secret("AWS_SECRET_ACCESS_KEY"),
    "aws_region": _secret("AWS_REGION", "eu-north-1"),
    "sidebar_dr": 3.5, "sidebar_horizon": 50, "sidebar_currency": "NIS",
    "scenario": "Stable (Baseline)",
    "escalation_rates": CLIMATE_SCENARIOS["Stable (Baseline)"],
    "audit_acknowledged": False,
    "audit_corrections_applied": False,
    "use_defaults_flag": False,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("## 🌍 Climate Adaptation CBA Tool")

# Anthropic key — always shown, always user-provided
_kb_from_server = bool(_secret("BEDROCK_KB_ID"))
with st.expander("⚙️ Enter your Anthropic API Key", expanded=not st.session_state.api_key):
    _k = st.text_input("Anthropic API Key", type="password",
                        value=st.session_state.api_key,
                        help="Get your key at console.anthropic.com — each user needs their own.")
    if _k:
        sanitized = _k.strip().encode("ascii", errors="ignore").decode("ascii")
        if sanitized != _k.strip():
            st.warning("⚠️ Non-ASCII characters were removed from your API key.")
        st.session_state.api_key = sanitized
    if _kb_from_server:
        st.caption("✅ Knowledge base (UHI papers) is pre-configured — no AWS setup needed.")

# AWS/Bedrock — only show manual fields if NOT loaded from server secrets
if not _kb_from_server:
    with st.expander("📚 Bedrock Knowledge Base (optional)", expanded=False):
        st.caption("Connect your UHI academic papers knowledge base for richer analysis")
        c1, c2 = st.columns(2)
        with c1:
            v = st.text_input("AWS Access Key ID", type="password", value=st.session_state.aws_key)
            if v: st.session_state.aws_key = v.strip()
            v = st.text_input("Knowledge Base ID", value=st.session_state.kb_id, placeholder="e.g. EWI8BYELS2")
            if v: st.session_state.kb_id = v.strip()
        with c2:
            v = st.text_input("AWS Secret Access Key", type="password", value=st.session_state.aws_secret)
            if v: st.session_state.aws_secret = v.strip()
            v = st.text_input("AWS Region", value=st.session_state.aws_region)
            if v: st.session_state.aws_region = v.strip()

st.markdown("---")

# ── Stage badge ───────────────────────────────────────────────────────────────
stages = {"problem": "1/3 · Problem", "measures": "2/3 · Measures", "data": "3/3 · Analysis", "done": "✓ Complete"}
badge_html = f'<div class="status-badge">{stages.get(st.session_state.stage, "")}</div>'
if st.session_state.specialist_type == "natural_shading":
    badge_html += '<div class="specialist-badge">🌿 Natural Shading Mode</div>'
elif st.session_state.specialist_type == "green_roof":
    badge_html += '<div class="specialist-badge">🏠 Green Roof Mode</div>'
st.markdown(badge_html, unsafe_allow_html=True)

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab_analysis, tab_portfolio, tab_tables, tab_method = st.tabs([
    "Analysis", "📊 Portfolio", "Detailed Tables", "Methodology"
])

with tab_analysis:
    # ── Chat display ──────────────────────────────────────────────────────────
    chat_container = st.container()
    with chat_container:
        if not st.session_state.messages:
            st.markdown("""
<div class="landing-hero">
  <h2>BcAi &mdash; Climate Benefit-Cost Intelligence</h2>
  <p>Automated, audit-ready cost-benefit analysis for climate adaptation decisions.
     Powered by peer-reviewed literature and live Excel formula outputs.</p>
</div>
<div class="cap-grid">
  <div class="cap-card"><strong>📊 CBA Automation</strong>
    Claude generates a full NPV/BCR analysis from your problem description in minutes.</div>
  <div class="cap-card"><strong>📚 Literature RAG</strong>
    Benefit parameters sourced automatically from peer-reviewed meta-analyses and OECD guidelines.</div>
  <div class="cap-card"><strong>🔒 Audit-Ready Excel</strong>
    Every figure is a live formula — change one input and all results cascade instantly.</div>
</div>
<div style="font-size:0.84rem;color:#374151;margin-bottom:0.35rem;font-weight:500;">Common use cases</div>
<div class="usecase-row">
  <span class="usecase-tag">🌡️ Urban Heat Island</span>
  <span class="usecase-tag">🌊 Coastal Flooding</span>
  <span class="usecase-tag">🏙️ Green Infrastructure</span>
  <span class="usecase-tag">💧 Drought Resilience</span>
  <span class="usecase-tag">🌿 Nature-Based Solutions</span>
</div>
<div class="cta-hint">↓ Describe your climate challenge below to begin — include location, affected population, and climate hazard.</div>
<hr style="border:none;border-top:1px solid #e2e8f0;margin:0.8rem 0 0.3rem;">
<div style="font-size:0.75rem;color:#9ca3af;text-align:center;">
  By Dan Brodsky, Gur Angel &amp; Nir Becker &nbsp;·&nbsp; Climate Economics Lab
</div>
""", unsafe_allow_html=True)
        for msg in st.session_state.messages:
            if msg["role"] == "user":
                st.markdown(f'<div class="chat-msg-user">{msg["content"]}</div>', unsafe_allow_html=True)
            else:
                html_content = md_lib.markdown(msg["content"], extensions=["tables", "nl2br"])
                st.markdown(f'<div class="chat-msg-ai">{html_content}</div>', unsafe_allow_html=True)

    # ── File upload (MUST be outside st.form — Streamlit constraint) ─────────
    if st.session_state.stage in ("problem", "measures", "data"):
        uploaded_file = st.file_uploader(
            "📎 Attach a document (PDF, Excel, CSV — optional)",
            type=["pdf", "xlsx", "xls", "csv"],
            key="file_uploader",
        )
        if uploaded_file is not None:
            extracted = _parse_uploaded_file(uploaded_file)
            if not extracted.startswith("[File parse error"):
                st.session_state.uploaded_file_text = extracted
                structured = _extract_structured_data(extracted)
                with st.expander(
                    f"✅ Parsed: {uploaded_file.name} ({len(extracted):,} chars)", expanded=False
                ):
                    if structured:
                        st.code(structured, language=None)
                    st.text_area("Raw extract preview", extracted[:500], height=80, disabled=True)
            else:
                st.warning(extracted)

    # ── Input form ────────────────────────────────────────────────────────────
    if st.session_state.stage != "done":
        with st.form("chat_form", clear_on_submit=True):
            col1, col2 = st.columns([6, 1])
            with col1:
                user_input = st.text_input(
                    "",
                    placeholder="Type your message…  (Enter to send)",
                    label_visibility="collapsed",
                    key="user_input",
                )
            with col2:
                send = st.form_submit_button("Send →", use_container_width=True)
    else:
        send = False
        user_input = ""

    # ── Excel download (shown inside Analysis tab when done) ──────────────────
    if st.session_state.stage == "done" and st.session_state.analysis_data:
        st.markdown("---")
        # ── Methodology Audit (runs before Excel build) ────────────────────────
        _challenge = st.session_state.get("challenge_type", "general")
        _audit_findings = _run_methodology_audit(st.session_state.analysis_data, challenge_type=_challenge)
        _download_blocked = _render_validation_report(_audit_findings, data=st.session_state.analysis_data)

        col_l, col_r = st.columns([1, 2])
        with col_l:
            with st.spinner("Building Excel..."):
                tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                tmp.close()  # must close on Windows before another process writes to the path
                _sp = st.session_state.get("escalation_rates", CLIMATE_SCENARIOS["Stable (Baseline)"]).copy()
                _sp["name"] = st.session_state.get("scenario", "Stable (Baseline)")
                # Apply sidebar DR and Horizon overrides so Excel matches the UI settings
                _effective_data = dict(st.session_state.analysis_data)
                _effective_data["discount_rate"] = st.session_state.get("sidebar_dr", 3.5) / 100
                _effective_data["time_horizon"]  = int(st.session_state.get("sidebar_horizon", 50))
                build_excel(_effective_data, tmp.name, scenario_params=_sp,
                            audit_acknowledged=st.session_state.get("audit_acknowledged", False),
                            challenge_type=st.session_state.get("challenge_type", "general"))
                with open(tmp.name, "rb") as f:
                    excel_bytes = f.read()
                os.unlink(tmp.name)

            title = st.session_state.analysis_data.get("problem_title", "CBA").replace(" ", "_")
            _acknowledged = st.session_state.get("audit_acknowledged", False)
            _can_download = not _download_blocked or _acknowledged
            if _can_download:
                st.download_button(
                    "⬇ Download Excel CBA Model",
                    excel_bytes,
                    file_name=f"CBA_{title}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                if _acknowledged and _download_blocked:
                    st.caption("⚠️ Downloaded with unverified parameters — review highlighted cells in Excel.")
            else:
                st.button(
                    "⬇ Download Excel CBA Model",
                    disabled=True,
                    help="Fix errors above or click 'Acknowledge & Download Anyway' to unlock.",
                )
        with col_r:
            st.markdown("**Your Excel file includes:**")
            bullets = (
                "- 📋 Executive Summary (first tab, NPV & BCR dashboard)\n"
                "- 📋 Inputs & Assumptions sheet (blue = editable)\n"
                "- 📊 CBA Results with NPV & BCR\n"
                "- 📈 Sensitivity Analysis with conditional formatting\n"
                "- 📁 Summary sheet\n"
                "- 📚 Parameter Registry (all 55+ parameters with citations)"
            )
            if st.session_state.analysis_data.get("specialist_type"):
                bullets += "\n- 🔬 Specialist Detail (VSL derivation, year-by-year benefit table)"
            bullets += "\n- 📈 Yearly_Projection_Model (50-yr dynamic benefit table)"
            st.markdown(bullets)

        if st.button("🔄 Start New Analysis"):
            for k in ["messages", "stage", "analysis_data", "problem_text", "specialist_type",
                      "challenge_type", "custom_measure_context",
                      "uploaded_file_text", "escalation_rates", "scenario",
                      "sidebar_dr", "audit_acknowledged", "audit_corrections_applied"]:
                if k == "messages":
                    st.session_state[k] = []
                elif k == "stage":
                    st.session_state[k] = "problem"
                elif k in ("specialist_type", "analysis_data", "custom_measure_context"):
                    st.session_state[k] = None
                elif k == "challenge_type":
                    st.session_state[k] = "general"
                elif k == "escalation_rates":
                    st.session_state[k] = CLIMATE_SCENARIOS["Stable (Baseline)"]
                elif k == "scenario":
                    st.session_state[k] = "Stable (Baseline)"
                elif k == "sidebar_dr":
                    st.session_state[k] = 3.5
                elif k in ("audit_acknowledged", "audit_corrections_applied"):
                    st.session_state[k] = False
                else:
                    st.session_state[k] = ""
            st.rerun()

# ── Helpers for safe API calls ─────────────────────────────────────────────────
def _trunc(text: str, max_chars: int = 6000) -> str:
    """Truncate KB text to prevent context overflow."""
    return text[:max_chars] + "\n…[truncated for length]" if len(text) > max_chars else text

def _trim_history(messages, max_msgs: int = 6):
    """Keep only the last max_msgs messages to prevent context window overflow."""
    recent = messages[-max_msgs:]
    return [{"role": m["role"], "content": m["content"]} for m in recent]

def _api_error_msg(e) -> str:
    status = getattr(e, "status_code", "?")
    return (
        f"⚠️ The AI provider returned an error (HTTP {status}). "
        "This usually means the prompt was too long or the API key is invalid. "
        "Try describing your problem more briefly, or type **'use defaults'** to skip data questions."
    )

# ── Form processing (module-level so st.rerun() is safe) ───────────────────────
if send and user_input.strip() and st.session_state.api_key:
        # Client-side "use defaults" detection — persists across reruns
        if "use default" in user_input.lower():
            st.session_state["use_defaults_flag"] = True
        st.session_state.messages.append({"role": "user", "content": user_input})
        _key = st.session_state.api_key.strip().encode("ascii", errors="ignore").decode("ascii")
        client = anthropic.Anthropic(api_key=_key)

        # ── Stage: problem description ─────────────────────────────────────────
        if st.session_state.stage == "problem":
            st.session_state.problem_text = user_input
            with st.spinner("Searching academic literature for relevant adaptation measures..."):

                # Run 3 targeted KB queries to find measures from the literature
                kb_measures   = query_bedrock_kb(
                    f"adaptation measures interventions urban heat island heatwave {user_input}"
                )
                kb_costs      = query_bedrock_kb(
                    f"cost effectiveness implementation cost per unit urban cooling {user_input}"
                )
                kb_mortality  = query_bedrock_kb(
                    f"heat mortality benefit VSL avoided deaths cooling intervention"
                )

                kb_section = ""
                for label, kb_text in [
                    ("ADAPTATION MEASURES FROM LITERATURE", kb_measures),
                    ("COST RANGES FROM LITERATURE",          kb_costs),
                    ("HEALTH BENEFIT DATA FROM LITERATURE",  kb_mortality),
                ]:
                    if kb_text and not kb_text.startswith("[KB unavailable"):
                        kb_section += f"\n\n--- {label} ---\n{_trunc(kb_text)}"
                # Cap total KB section to prevent context overflow
                kb_section = _trunc(kb_section, 16000)

                if kb_section:
                    kb_instruction = f"""
The following passages come DIRECTLY from peer-reviewed academic papers in the knowledge base.
Your proposed measures MUST be drawn from these papers — do not invent measures that are not mentioned.
Cite the paper name when proposing each measure.

{kb_section}

---"""
                else:
                    kb_instruction = ""

                try:
                    resp = client.messages.create(
                        model="claude-opus-4-6", max_tokens=1500,
                        messages=[{
                            "role": "user",
                            "content": f"""You are an expert in climate adaptation economics.
{kb_instruction}
The user described this climate problem:
"{user_input}"

Do two things:
1. Briefly confirm your understanding of the problem (2-3 sentences).

2. Propose 3-5 concrete adaptation measures for this problem.
   {"Each measure MUST come from the academic papers cited above — reference the paper." if kb_section else "Use best-practice literature."}
   For each measure provide:
   - Name
   - 1-sentence description
   - Source paper (if available)
   - Typical cost range from the literature
   - Main benefit categories (e.g. avoided mortality, energy savings, morbidity reduction)
   - The KEY formula used to calculate the main benefit (e.g. "Avoided mortality = VSL × population × mortality_rate × heat_reduction_efficiency")

Then ask: "Which of these measures would you like to include in the analysis?"

FORMAT RULES — follow exactly:
- For benefit formulas use: **Benefit Formula:** _Annual Benefit = A × B × C_
- List formula parameters in a compact table: | Parameter | Default | Source |
- Use × for multiplication, ÷ for division. Never use * or / in displayed formulas.
- Do NOT use triple-backtick code blocks for any formula or equation.

Be concise and cite sources."""
                        }]
                    )
                except (anthropic.BadRequestError, anthropic.APIStatusError) as e:
                    st.session_state.messages.append({"role": "assistant", "content": _api_error_msg(e)})
                    st.rerun()
            reply = resp.content[0].text
            st.session_state.messages.append({"role": "assistant", "content": reply})
            st.session_state.specialist_type  = detect_specialist_type(st.session_state.problem_text)
            st.session_state.challenge_type   = detect_challenge_type(st.session_state.problem_text)
            # Purge UV/skin-cancer specialist logic when challenge is non-heat
            if st.session_state.specialist_type == "natural_shading" and st.session_state.challenge_type not in ("heat", "general"):
                st.session_state.specialist_type = None
            st.session_state.stage = "measures"
            st.rerun()

        # ── Stage: user selects measures ───────────────────────────────────────
        elif st.session_state.stage == "measures":
            # Detect custom measure request
            _CUSTOM_TRIGGERS = [
                "custom measure", "my own measure", "propose a measure",
                "not in the list", "add a measure", "different measure",
                "own intervention", "new measure", "define a measure",
            ]
            _is_custom = any(kw in user_input.lower() for kw in _CUSTOM_TRIGGERS)

            with st.spinner("Reading literature for selected measures..."):

                # Targeted KB queries for the selected measures
                kb_formulas = query_bedrock_kb(
                    f"benefit formula VSL mortality rate heat reduction efficiency {user_input}"
                )
                kb_params   = query_bedrock_kb(
                    f"parameter value cost unit energy savings electricity {user_input} {st.session_state.problem_text}"
                )

                kb_section = ""
                for label, kb_text in [
                    ("BENEFIT FORMULAS & PARAMETERS", kb_formulas),
                    ("COST & UNIT VALUES",             kb_params),
                ]:
                    if kb_text and not kb_text.startswith("[KB unavailable"):
                        kb_section += f"\n\n--- {label} ---\n{_trunc(kb_text)}"
                kb_section = _trunc(kb_section, 12000)

                if kb_section:
                    kb_block = f"""
The following passages were retrieved from peer-reviewed papers.
Use them to identify:
  (a) The formula for calculating each benefit (e.g. mortality = VSL × pop × rate × efficiency)
  (b) Which parameters must come from the user (population, area, local costs)
  (c) Which parameter values are already in the literature (VSL, mortality factors, efficiency rates)

{kb_section}

---"""
                else:
                    kb_block = ""

                _custom_preamble = ""
                if _is_custom:
                    _custom_preamble = """The user wants to define a CUSTOM MEASURE not found in the standard literature list.

Your job in this turn:
1. Ask for: (a) measure name, (b) brief description, (c) estimated CAPEX, (d) annual OPEX, (e) expected lifetime (years), (f) what hazard it addresses.
2. Search the knowledge base passages above for ANY analogous measures with similar benefit pathways (e.g. if they say "smart pavement", look for cool pavement / albedo interventions).
3. Propose a benefit formula using the closest analogous literature, labelling it as "benefit transfer from [source]".
4. Ask the user: "Does this formula and these parameters look right? Type 'confirm' to proceed."

"""

                history = _trim_history(st.session_state.messages)
                history.append({
                    "role": "user",
                    "content": f"""{_custom_preamble}The user selected these measures: "{user_input}"
{kb_block}
For each selected measure:

1. State the benefit formula clearly.
   **Benefit Formula:** _Annual Benefit = A × B × C_ (use this exact Markdown style — no code blocks)
   Cite the source paper for each parameter value you will use as a default.

2. Split data needs into two explicit groups:
   **I need from you (local data):** e.g. population size, project area, local electricity price
   **I will use from literature (with source):** e.g. VSL = €3.8M (OECD 2012), heat_mortality_factor = 0.00083 (Gasparrini 2017)

3. Ask the user ONLY for the local data. Keep questions short and grouped by measure.

FORMAT RULES — follow exactly:
- Use × for multiplication, ÷ for division. Never use * or / in displayed formulas.
- Do NOT use triple-backtick code blocks for any formula or equation.
- List parameters in a compact Markdown table: | Parameter | Default | Source |

If the user has already said "use defaults", skip the questions and confirm you will use literature defaults for everything."""
                })
                try:
                    resp = client.messages.create(model="claude-opus-4-6", max_tokens=1500, messages=history)
                except (anthropic.BadRequestError, anthropic.APIStatusError) as e:
                    st.session_state.messages.append({"role": "assistant", "content": _api_error_msg(e)})
                    st.rerun()

            reply = resp.content[0].text
            st.session_state.messages.append({"role": "assistant", "content": reply})
            if _is_custom:
                st.session_state.custom_measure_context = _trunc(reply, 1500)
            st.session_state.stage = "data"
            st.rerun()

        # ── Stage: user provides data → build analysis ─────────────────────────
        elif st.session_state.stage == "data":
            # Reset audit flags for this new generation run
            st.session_state.audit_acknowledged = False
            st.session_state.audit_corrections_applied = False

            stype = st.session_state.specialist_type
            if stype == "natural_shading":
                prompt_text = NATURAL_SHADING_DATA_PROMPT.replace("USER_INPUT_PLACEHOLDER", user_input)
            elif stype == "green_roof":
                prompt_text = GREEN_ROOF_DATA_PROMPT.replace("USER_INPUT_PLACEHOLDER", user_input)
            else:
                prompt_text = GENERIC_DATA_PROMPT.replace("USER_INPUT_PLACEHOLDER", user_input)
            # Scale max_tokens with measure count; floor 12000, cap 16000
            _prev_msg = next((m["content"] for m in reversed(st.session_state.messages)
                              if m["role"] == "assistant"), "")
            _n_measures_hint = max(1, _prev_msg.count("Measure") + _prev_msg.count("measure"))
            max_tok = min(16000, max(12000, _n_measures_hint * 2000))

            # Inject flood-specific guidance when challenge type is flood
            _challenge = st.session_state.get("challenge_type", "general")
            if _challenge == "flood":
                flood_injection = """CHALLENGE TYPE: COASTAL FLOODING — read before generating JSON.

For flood mortality DO NOT use heat_mortality_factor or heat_reduction_efficiency.
Instead use one of:
  a) "generic_annual" type with value = annual_lives_saved × VSL / 1,000,000
  b) "avoided_mortality" with heat_mortality_factor = 0.0 and a note in the source field
     that this is flood mortality, not heat mortality.

Israel flood mortality defaults:
- mortality_rate (general population): 0.008 (CBS 2022 Life Tables, age-standardised)
- flood_mortality_reduction_efficiency: 0.40–0.70 (early warning systems)
- Do NOT copy heat_mortality_factor from heatwave literature.

"""
                prompt_text = flood_injection + prompt_text

            # Inject custom measure context if user defined a custom measure in Stage 2
            _cmc = st.session_state.get("custom_measure_context")
            if _cmc:
                prompt_text += f"""

CUSTOM MEASURE CONTEXT — the user defined a custom measure in the previous step.
Ensure it is included as a measure in the JSON output with the benefit formula proposed below.
Prior discussion: {_cmc}

---"""

            # Query KB for formula parameters from literature
            kb_cba = query_bedrock_kb(
                f"VSL mortality rate heat reduction efficiency cost per unit {st.session_state.problem_text}"
            )
            if kb_cba and not kb_cba.startswith("[KB unavailable"):
                prompt_text += f"""

LITERATURE PARAMETER VALUES — use these to populate formula fields and their _source fields:

{_trunc(kb_cba)}

---"""

            # Inject uploaded document content as local data source
            _file_text = st.session_state.get("uploaded_file_text", "").strip()
            if _file_text:
                _structured_hint = _extract_structured_data(_file_text)
                _file_section = (
                    f"\n\nUPLOADED DOCUMENT (treat as local data source — override defaults):\n"
                    f"{_trunc(_file_text, 4000)}"
                )
                if _structured_hint:
                    _file_section += f"\n\nPRE-PARSED PARAMETERS:\n{_structured_hint}"
                prompt_text += _file_section

            def _extract_json(text: str):
                text = text.strip()
                try:
                    return json.loads(text)
                except Exception:
                    pass
                fence = re.search(r"```(?:json)?\s*([\s\S]*?)```", text)
                if fence:
                    try:
                        return json.loads(fence.group(1).strip())
                    except Exception:
                        pass
                start = text.find("{")
                if start != -1:
                    depth = 0
                    for i, ch in enumerate(text[start:], start):
                        if ch == "{": depth += 1
                        elif ch == "}":
                            depth -= 1
                            if depth == 0:
                                try:
                                    return json.loads(text[start:i+1])
                                except Exception:
                                    break
                return None

            with st.spinner("Running Economic Simulations..."):
                history = _trim_history(st.session_state.messages)
                history.append({"role": "user", "content": prompt_text})
                try:
                    resp = client.messages.create(
                        model="claude-opus-4-6", max_tokens=max_tok, messages=history
                    )
                except (anthropic.BadRequestError, anthropic.APIStatusError) as e:
                    st.session_state.messages.append({"role": "assistant", "content": _api_error_msg(e)})
                    st.rerun()
                except Exception as e:
                    st.session_state.messages.append({"role": "assistant", "content": f"⚠️ Unexpected API error: {e}"})
                    st.rerun()

            raw = resp.content[0].text
            data = _extract_json(raw)

            if data is None:
                with st.spinner("Refining output format..."):
                    retry_messages = history + [
                        {"role": "assistant", "content": raw},
                        {"role": "user", "content":
                            "Your response was not valid JSON. "
                            "Output ONLY the JSON object. "
                            "No explanations, no markdown fences. "
                            "Start with { and end with }."}
                    ]
                    try:
                        resp2 = client.messages.create(
                            model="claude-opus-4-6", max_tokens=max_tok, messages=retry_messages
                        )
                        data = _extract_json(resp2.content[0].text)
                    except (anthropic.BadRequestError, anthropic.APIStatusError) as e:
                        st.session_state.messages.append({"role": "assistant", "content": _api_error_msg(e)})
                        st.rerun()
                    except Exception as e:
                        st.session_state.messages.append({"role": "assistant", "content": f"⚠️ Unexpected API error: {e}"})
                        st.rerun()

            # ── Israel hard-coded fallback (bypasses RAG drift) ────────────────
            _ISRAEL_DEFAULTS = {
                "vsl":                  11_500_000,   # NIS — OECD 2005 CPI/PPP/FX adjusted
                "heat_mortality_factor": 0.035,       # Gasparrini 2017, Mediterranean median
                "mortality_rate":        0.008,       # CBS 2022 general population
                "heatwave_days":         35,          # Israel Meteorological Service baseline
            }
            def _apply_israel_defaults(d: dict) -> dict:
                """If use_defaults_flag is set and currency is NIS, clamp benefit_component params."""
                if not st.session_state.get("use_defaults_flag"):
                    return d
                if (d.get("currency") or "").upper() != "NIS":
                    return d
                import copy; d = copy.deepcopy(d)
                for m in d.get("measures", []):
                    for comp in (m.get("benefit_components") or []):
                        if comp.get("type") == "avoided_mortality":
                            for param, val in _ISRAEL_DEFAULTS.items():
                                if param != "heatwave_days":  # not a direct component field
                                    comp[param] = val
                                    comp[f"{param}_source"] = (
                                        comp.get(f"{param}_source", "") +
                                        " [Israel hard-coded baseline]"
                                    ).strip()
                return d

            if data:
                data = _apply_israel_defaults(data)
                st.session_state.analysis_data = data
                sources_used = set()
                for m in data.get("measures", []):
                    for comp in (m.get("benefit_components") or []):
                        for k, v in comp.items():
                            if k.endswith("_source") and isinstance(v, str) and len(v) > 5:
                                sources_used.add(v[:60])
                specialist_note = ""
                if data.get("specialist_type") == "natural_shading":
                    specialist_note = " using specialist VSL/CDD Natural Shading methodology"
                elif data.get("specialist_type") == "green_roof":
                    specialist_note = " using specialist VSL/CDD Green Roof methodology"
                sources_note = ""
                if sources_used:
                    sources_note = "\n\n📚 **Sources used:** " + " | ".join(list(sources_used)[:4])
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": (
                        f"✅ Analysis complete for **{data['problem_title']}**. "
                        f"{len(data['measures'])} measure(s) analyzed{specialist_note}. "
                        f"Each benefit is computed via a live Excel formula with cited sources.{sources_note}\n\n"
                        f"Click **Download Excel** below."
                    )
                })
                st.session_state.stage = "done"
            else:
                # Check whether raw response looks truncated (no closing brace)
                _raw_text = raw if "raw" in dir() else ""
                _truncated = _raw_text and not _raw_text.rstrip().endswith("}")
                _msg = (
                    "⚠️ Analysis truncated mid-output — the JSON was incomplete. "
                    "Try reducing the number of measures (aim for 3–4) or type **'use defaults'** "
                    "so descriptions stay concise."
                    if _truncated else
                    "I wasn't able to generate a valid analysis. "
                    "Please try again or type **'use defaults'** to let me fill everything from literature."
                )
                st.session_state.messages.append({"role": "assistant", "content": _msg})
            st.rerun()

# ── Portfolio Summary tab ───────────────────────────────────────────────────────
with tab_portfolio:
    _port_data = st.session_state.analysis_data
    if _port_data and st.session_state.stage == "done":
        _cur  = _port_data.get("currency", "")
        _dr   = st.session_state.get("sidebar_dr", 3.5) / 100
        _vmlt = 1.0
        _fins = _compute_measure_financials(_port_data, dr=_dr, vsl_mult=_vmlt)

        # What-If active banner
        _orig_dr = _port_data.get("discount_rate", 0.035) or 0.035
        if abs(_dr - _orig_dr) > 0.001:
            st.info(
                f"⚗️ **What-If mode active** — DR={_dr*100:.1f}% applied to both charts and Excel download."
            )

        # ── KPI metrics ──────────────────────────────────────────────────────
        st.subheader("Portfolio Executive Summary")
        _best = max(_fins, key=lambda f: f["bcr"], default=None)
        _total_npv = sum(f["npv"] for f in _fins)
        _total_pvb = sum(f["pv_ben"] for f in _fins)
        _total_pvc = sum(f["pv_cost"] for f in _fins)
        _port_bcr  = _total_pvb / _total_pvc if _total_pvc > 0 else 0.0
        _viable_count = sum(1 for f in _fins if f["bcr"] >= 1.0)

        # Climate scenario label
        _scenario_label = _port_data.get("scenario_name", st.session_state.get("scenario", "Stable (Baseline)"))
        st.caption(f"🌡️ Climate scenario: **{_scenario_label}** · DR: **{_dr*100:.1f}%**")

        _k1, _k2, _k3, _k4, _k5 = st.columns(5)
        _k1.metric(
            "Most Viable Measure",
            _best["name"] if _best else "—",
            delta=f"BCR {_best['bcr']:.2f}" if _best else None,
        )
        _k2.metric(
            "Total Portfolio NPV",
            f"{_total_npv:,.1f} {_cur}M",
            delta="Positive ✓" if _total_npv > 0 else "Negative ✗",
        )
        _k3.metric(
            "Portfolio BCR",
            f"{_port_bcr:.2f}",
            delta="≥1.0 viable" if _port_bcr >= 1.0 else "< 1.0 review",
        )
        _k4.metric("Measures Analyzed", str(len(_fins)))
        _k5.metric("Viable (BCR ≥ 1.0)", f"{_viable_count} / {len(_fins)}")

        # ── BCR bar chart ─────────────────────────────────────────────────────
        st.divider()
        st.subheader("Benefit-Cost Ratio by Measure")
        st.caption("Green = BCR ≥ 1.5 (recommended) · Amber = BCR 1.0–1.5 · Red = BCR < 1.0")
        _bcr_fig = _build_bcr_bar_chart(_fins, _cur)
        if _bcr_fig:
            st.plotly_chart(_bcr_fig, use_container_width=True)

        # ── Investment Map ────────────────────────────────────────────────────
        st.divider()
        st.subheader("Investment Map — CAPEX vs NPV")
        st.caption("Bubble size = PV of Benefits. Colour = BCR. Top-left quadrant = Low-Hanging Fruit.")
        _inv_fig = _build_investment_map(_fins, _cur)
        if _inv_fig:
            st.plotly_chart(_inv_fig, use_container_width=True)

        # ── Waterfall ─────────────────────────────────────────────────────────
        st.divider()
        with st.expander("📊 Portfolio Cash-Flow Breakdown (Waterfall)", expanded=True):
            st.caption(
                "Aggregate portfolio view: upward bars = discounted benefits by category; "
                "downward bars = costs."
            )
            _wf_fig = _build_waterfall_chart(_fins, _port_data, _cur)
            if _wf_fig:
                st.plotly_chart(_wf_fig, use_container_width=True)

        # ── Payback period table ───────────────────────────────────────────────
        st.divider()
        st.subheader("Payback & Comparison Table")
        import pandas as _pd
        _rows = []
        for _f in _fins:
            _ann_ben = _f["pv_ben"] / max(_f["life"], 1)  # rough annualised benefit
            _ann_net = _ann_ben - _f.get("opex", 0)
            if _ann_net > 0 and _f["capex"] > 0:
                _pb = f"{_f['capex'] / _ann_net:.1f} yr"
            else:
                _pb = "Never"
            _rows.append({
                "Measure": _f["name"],
                f"NPV ({_cur}M)": round(_f["npv"], 2),
                "BCR": round(_f["bcr"], 2),
                f"CAPEX ({_cur}M)": round(_f["capex"], 2),
                "Payback": _pb,
                "Viable": "✅" if _f["bcr"] >= 1.0 else "❌",
            })
        if _rows:
            _df_port = _pd.DataFrame(_rows)
            st.dataframe(_df_port, use_container_width=True, hide_index=True)

        # ── Data Gaps & Next Steps ─────────────────────────────────────────────
        _gaps = _port_data.get("data_gaps", "")
        if _gaps:
            st.divider()
            with st.expander("📋 Data Gaps & Next Steps", expanded=False):
                st.info(_gaps)
    else:
        st.info("Complete the analysis in the **Analysis** tab first to see the portfolio dashboard.")


# ── Detailed Tables tab ────────────────────────────────────────────────────────
with tab_tables:
    data_snapshot = st.session_state.analysis_data
    if data_snapshot and st.session_state.stage == "done":
        measures = data_snapshot.get("measures", [])
        cur = data_snapshot.get("currency_unit", "M")
        dr = data_snapshot.get("discount_rate", 0.035)
        horizon = data_snapshot.get("time_horizon", 50)

        # Metric cards — use _compute_measure_financials with What-If sliders
        _tdr   = st.session_state.get("sidebar_dr", 3.5) / 100
        _tvmlt = 1.0
        fins_t = _compute_measure_financials(data_snapshot, dr=_tdr, vsl_mult=_tvmlt)
        best_bcr   = max((f["bcr"]  for f in fins_t), default=0.0)
        best_npv   = max((f["npv"]  for f in fins_t), default=0.0)
        total_capex = sum(f["capex"] for f in fins_t)

        st.subheader("Key Metrics")
        st.caption(f"What-If: DR={_tdr*100:.1f}% | VSL×{_tvmlt:.1f} — see 📊 Portfolio tab for full dashboard")
        mc1, mc2, mc3 = st.columns(3)
        mc1.metric("Best BCR", f"{best_bcr:.2f}", delta="≥1.0 = viable" if best_bcr >= 1.0 else "< 1.0 not viable")
        mc2.metric("Best NPV", f"{best_npv:,.1f} {cur}M")
        mc3.metric("Total CAPEX", f"{total_capex:,.1f} {cur}M")

        st.divider()

        # Measures dataframe
        st.subheader("Measures Summary")
        rows = []
        for m in measures:
            rows.append({
                "Measure": m.get("name", ""),
                "CAPEX": m.get("capex", ""),
                "Annual OPEX": m.get("annual_opex", ""),
                "Lifetime (yr)": m.get("lifetime_years", ""),
                "Feasibility": m.get("feasibility", ""),
                "Uncertainty": m.get("uncertainty", ""),
            })
        st.dataframe(rows, use_container_width=True)

        # Sensitivity vars
        svars = data_snapshot.get("sensitivity_vars", [])
        if svars:
            st.subheader("Sensitivity Variables")
            st.dataframe(svars, use_container_width=True)

        # Key assumptions
        kassumptions = data_snapshot.get("key_assumptions", [])
        if kassumptions:
            st.subheader("Key Assumptions")
            for ka in kassumptions:
                if isinstance(ka, dict):
                    st.markdown(f"- **{ka.get('text','')}** — *{ka.get('source','')}*")
                else:
                    st.markdown(f"- {ka}")

        # ── Benefit Composition Charts ─────────────────────────────────────────
        st.divider()
        st.subheader("Benefit Composition by Measure")
        st.caption(
            "PV-weighted share of Total Benefits (Health / Energy / Other). "
            "Responds to What-If sliders. See Excel for full formula audit trail."
        )
        for i, m in enumerate(measures):
            fig = _build_benefit_pie(m, financials_entry=fins_t[i] if i < len(fins_t) else None)
            if fig:
                st.plotly_chart(fig, use_container_width=True)

        # ── 50-Year Projection Chart ───────────────────────────────────────
        st.divider()
        st.subheader("50-Year Annual Benefit Projection")
        st.caption(
            f"Scenario: **{st.session_state.get('scenario', 'Stable (Baseline)')}**  "
            "— Undiscounted annual benefits. Rising curves = rising cost of inaction."
        )
        proj_fig = _build_projection_chart(
            data_snapshot,
            st.session_state.get("escalation_rates", CLIMATE_SCENARIOS["Stable (Baseline)"]),
        )
        if proj_fig:
            st.plotly_chart(proj_fig, use_container_width=True)
        else:
            st.info("No benefit data available for projection chart.")
    else:
        st.info("Complete the analysis in the **Analysis** tab first to see results here.")

# ── Methodology tab ─────────────────────────────────────────────────────────────
with tab_method:
    _ch = st.session_state.get("challenge_type", "general")

    # ── Block A: Challenge-specific parameter table ──────────────────────────
    if _ch == "flood":
        st.subheader("Flood Risk Assessment — Key Parameters")
        st.markdown("""
| Parameter | Default Value | Source |
|-----------|--------------|--------|
| Storm surge return period | 100 years (1% AEP) | IPCC AR6 WGI Ch.9 |
| Sea level rise (Mediterranean) | 3–5 cm/decade | IPCC AR6 WGI Table 9.9 |
| Flood mortality rate (general pop.) | 0.008 | CBS Israel Life Tables 2022 |
| Flood depth-damage coefficient | 0.40 | Huizinga et al. (2017) JRC EUR 28552 EN |
| Early warning system effectiveness | 40–70% | UNDRR Global Assessment Report 2022 |

**Annual Avoided Deaths (Flood)** = Population × Mortality Rate × Flood Reduction Efficiency

**Annual Flood Damage Avoided (NIS M)** = Exposed Assets × Depth-Damage Coefficient × Flood Frequency × (1 − Mitigation Efficiency)
""")
    elif _ch == "drought":
        st.subheader("Drought Risk — Key Parameters")
        st.markdown("""
| Parameter | Default Value | Source |
|-----------|--------------|--------|
| Drought intensity multiplier | scenario-dependent | IPCC AR6 WGI Ch.11 |
| Water stress index (Mediterranean) | 0.4–0.7 | FAO AQUASTAT |
| Agricultural yield loss per σ drought | 8–15% | Lesk et al. (2016) *Nature* |
| Groundwater recharge reduction | 15–30%/°C warming | IPCC SRCCL (2019) |

**Annual Economic Loss (Drought)** = Agricultural Output × Yield Loss Fraction × Drought Intensity × (1 − Adaptation Efficiency)
""")
    else:
        st.subheader("VSL Derivation Chain")
        st.markdown("""
| Step | Parameter | Default Value | Source |
|------|-----------|---------------|--------|
| 1 | OECD Base VSL (2005 USD) | $3,000,000 | Viscusi & Masterman (2017); OECD ENV/WKP(2012)3 |
| 2 | × CPI Multiplier (2005→2023) | 1.68 | BLS CPI-U Series CUUR0000SA0 |
| 3 | = CPI-Adjusted VSL (2023 USD) | $5,040,000 | Computed |
| 4 | × PPP Ratio (Israel/OECD) | 0.89 | World Bank WDI NY.GDP.PCAP.PP.CD |
| 5 | × Income Elasticity | 1.00 | Standard for developed economies |
| 6 | = PPP-Adjusted VSL (USD) | $4,485,600 | Computed |
| 7 | × FX Rate (NIS/USD) | 3.70 | Bank of Israel |
| 8 | = VSL in NIS | ~16,597,000 | Computed |
| 9 | ÷ Life Expectancy (yr) | 35 | UN WPP |
| 10 | = VSLY (NIS/yr) | ~474,200 | Computed |

*All steps are live Excel formulas — change any input cell and the entire chain updates.*
""")
        st.divider()
        st.subheader("CDD / Heat-Mortality Formula")
        st.markdown("""
**Annual Avoided Deaths** = Population × Mortality Rate × Heat-Mortality Factor × Heat Reduction Efficiency

**Annual Avoided Mortality Benefit (NIS M)** = Avoided Deaths × VSL / 1,000,000

| Parameter | Value | Source |
|-----------|-------|--------|
| Heat-Mortality Factor | 0.00083 deaths/°C/person | Gasparrini et al. (2017) Lancet, Mediterranean cluster |
| Morbidity Multiplier | 10× | WHO Europe Heat Health Action Plan (2008) |
| CDD Baseline (Tel Aviv) | 735 CDD/yr (21°C base) | Israel Meteorological Service, 1990–2020 |
| Skin Cancer Incidence | 0.000161/person/yr | Israeli Cancer Registry; WHO IARC Monograph 100D |
""")

    # ── Block B: NPV / BCR Formulas — shared ────────────────────────────────
    st.divider()
    st.subheader("NPV / BCR Formulas")
    st.markdown(r"""
$$\text{NPV} = \sum_{t=1}^{T} \frac{B_t - C_t}{(1+r)^t}$$

$$\text{BCR} = \frac{PV(\text{Benefits})}{PV(\text{Costs})} = \frac{B/r \cdot (1-(1+r)^{-T})}{CAPEX + OPEX/r \cdot (1-(1+r)^{-T})}$$

Where: *B* = Annual benefit, *r* = Discount rate, *T* = Time horizon, *CAPEX* = Capital cost, *OPEX* = Annual operating cost.

**BCR > 1.0** = project generates more benefit than cost → viable
**BCR > 1.5** = recommended threshold for public infrastructure
""")

    # ── Block C: Benefit Transfer & 8-Step VSL — shared ─────────────────────
    st.divider()
    st.subheader("Benefit Transfer & General Methodology")
    st.markdown("""
This tool derives local monetary values via the **Benefit Transfer Protocol** (8 steps):

1. **Source study value** — e.g. OECD 2012 VSL baseline ($3M, 2005 USD)
2. **CPI inflation** — adjust to current year using BLS CPI-U index
3. **PPP income adjustment** — scale by Israel/OECD GDP-per-capita ratio
4. **Income elasticity** — apply elasticity = 1.0 (standard for developed economies)
5. **FX conversion** — multiply by NIS/USD exchange rate
6. **VSLY derivation** — VSL ÷ remaining life expectancy of affected demographic
7. **Domain parameter** — multiply by hazard-specific factor (heat mortality fraction, flood depth-damage coefficient, etc.)
8. **Discounting** — convert annual benefit stream to NPV using project discount rate

All monetary values in the Excel model are live formulas — edit any blue cell to update the entire analysis.
""")

    # ── Block D: Key Citations — challenge-specific ──────────────────────────
    st.divider()
    st.subheader("Key Citations")
    if _ch == "flood":
        st.markdown("""
- **Huizinga, J. et al. (2017).** Global flood depth-damage functions. *JRC Technical Report EUR 28552 EN*. European Commission.
- **IPCC (2021).** AR6 WGI Chapter 9: Ocean, Cryosphere and Sea Level Change. *Sixth Assessment Report*.
- **UNDRR (2022).** Global Assessment Report on Disaster Risk Reduction 2022. United Nations.
- **CBS Israel (2022).** Life Tables 2020–2022. Central Bureau of Statistics, Israel.
- **Viscusi, W.K. & Masterman, C.W. (2017).** Income Elasticities and Global Values of a Statistical Life. *Journal of Benefit-Cost Analysis*, 8(2), 226–250.
- **OECD (2012).** Mortality Risk Valuation in Environment, Health and Transport Policies. ENV/WKP(2012)3.
""")
    elif _ch == "drought":
        st.markdown("""
- **Lesk, C. et al. (2016).** Influence of extreme weather disasters on global crop production. *Nature*, 529, 84–87.
- **IPCC (2019).** Special Report on Climate Change and Land (SRCCL). Chapter 5: Food Security.
- **FAO AQUASTAT.** Global Information System on Water and Agriculture. Food and Agriculture Organization.
- **IPCC (2021).** AR6 WGI Chapter 11: Weather and Climate Extreme Events in a Changing Climate.
- **Viscusi, W.K. & Masterman, C.W. (2017).** Income Elasticities and Global Values of a Statistical Life. *Journal of Benefit-Cost Analysis*, 8(2), 226–250.
- **OECD (2012).** Mortality Risk Valuation in Environment, Health and Transport Policies. ENV/WKP(2012)3.
""")
    else:
        st.markdown("""
- **Viscusi, W.K. & Masterman, C.W. (2017).** Income Elasticities and Global Values of a Statistical Life. *Journal of Benefit-Cost Analysis*, 8(2), 226–250.
- **Gasparrini, A. et al. (2017).** Projections of temperature-related excess mortality under climate change scenarios. *The Lancet Planetary Health*, 1(9), e360–e367.
- **OECD (2012).** Mortality Risk Valuation in Environment, Health and Transport Policies. OECD Publishing. ENV/WKP(2012)3.
- **WHO Europe (2008).** Heat–health action plans. WHO Regional Office for Europe, Copenhagen.
- **Shashua-Bar, L. & Hoffman, M.E. (2000).** Vegetation as a climatic component in the design of an urban street. *Energy and Buildings*, 31(3), 221–235.
- **Nowak, D.J. et al. (2002).** Brooklyn's Urban Forest. General Technical Report NE-290. USDA Forest Service.
- **Berghage, R. et al. (2009).** Green Roofs for Stormwater Runoff Control. EPA/600/R-09/026.
- **BLS CPI-U Series CUUR0000SA0.** U.S. Bureau of Labor Statistics. https://www.bls.gov/cpi/
- **World Bank WDI NY.GDP.PCAP.PP.CD.** World Development Indicators. https://data.worldbank.org
- **Israeli Cancer Registry.** Ministry of Health, State of Israel. https://www.health.gov.il
""")
