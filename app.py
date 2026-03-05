import streamlit as st
import anthropic
import json
import re
from excel_builder import build_excel
import tempfile, os
import markdown as md_lib

st.set_page_config(page_title="Climate CBA Tool", page_icon="🌍", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500&display=swap');
* { font-family: 'IBM Plex Sans', sans-serif; }

.stApp {
    background: #f4f7f5;
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
    background: #065f46;
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
    border: 1px solid #d1fae5;
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
    font-family: 'IBM Plex Mono', monospace !important;
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
    font-family: 'IBM Plex Mono', monospace !important;
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
    font-family: 'IBM Plex Mono', monospace !important;
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
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    padding: 0.15rem 0.6rem;
    border-radius: 999px;
    margin-bottom: 1rem;
    margin-left: 0.5rem;
}
</style>
""", unsafe_allow_html=True)

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


# ── Specialist prompt constants ─────────────────────────────────────────────────
# Use USER_INPUT_PLACEHOLDER as a safe substitution token (avoids f-string brace conflicts)

GENERIC_DATA_PROMPT = """User provided: "USER_INPUT_PLACEHOLDER"

Extract all data and produce a JSON object for the CBA model. Respond ONLY with valid JSON, no markdown fences:

{
  "problem_title": "short title",
  "problem_summary": "2-sentence summary",
  "discount_rate": 0.035,
  "time_horizon": 30,
  "currency": "USD",
  "currency_unit": "millions",
  "specialist_type": null,
  "measures": [
    {
      "name": "Measure name",
      "description": "What it involves",
      "category": "Infrastructure|Nature-based|Policy|Technology",
      "capex": 10.0,
      "annual_opex": 0.5,
      "benefit_types": ["flood protection", "tourism"],
      "co_benefits": "Non-monetary benefits",
      "lifetime_years": 30,
      "feasibility": "High|Medium|Low",
      "uncertainty": "Low|Medium|High",
      "data_source": "Literature estimate / User provided"
    }
  ],
  "sensitivity_vars": [
    {"name": "Discount Rate", "base": 0.035, "low": 0.02, "high": 0.07, "unit": "%"},
    {"name": "Annual Benefit", "base": 1.0, "low": 0.5, "high": 1.5, "unit": "multiplier"}
  ],
  "key_assumptions": [
    {"text": "Main assumption or limitation", "source": "Source or 'Methodology default'"},
    {"text": "Another assumption", "source": "Literature reference or user-provided"}
  ],
  "data_gaps": "What additional data would improve this"
}

If user data is missing for some fields, use reasonable literature-based estimates and note them in data_source."""

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
for k, v in {
    "messages": [], "stage": "problem", "analysis_data": None,
    "api_key": "", "problem_text": "", "specialist_type": None
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
badge_html = f'<div class="status-badge">{stages.get(st.session_state.stage, "")}</div>'
if st.session_state.specialist_type == "natural_shading":
    badge_html += '<div class="specialist-badge">🌿 Natural Shading Mode</div>'
elif st.session_state.specialist_type == "green_roof":
    badge_html += '<div class="specialist-badge">🏠 Green Roof Mode</div>'
st.markdown(badge_html, unsafe_allow_html=True)

# ── Chat display ───────────────────────────────────────────────────────────────
chat_container = st.container()
with chat_container:
    if not st.session_state.messages:
        st.markdown("""
        <div class="chat-msg-ai">
        Hello! I'm here to help you build a cost-benefit analysis for climate adaptation measures.<br><br>
        <b>Describe the climate problem</b> you're working on — include location, population, climate hazard, and any economic context you have.<br><br>
        <i>Tip: mention "shaded boulevard", "urban trees", or "green roof" to activate specialist VSL/CDD methodology.</i>
        </div>
        """, unsafe_allow_html=True)
    for msg in st.session_state.messages:
        if msg["role"] == "user":
            st.markdown(f'<div class="chat-msg-user">{msg["content"]}</div>', unsafe_allow_html=True)
        else:
            html_content = md_lib.markdown(msg["content"], extensions=["tables", "nl2br"])
            st.markdown(f'<div class="chat-msg-ai">{html_content}</div>', unsafe_allow_html=True)

# ── Input ──────────────────────────────────────────────────────────────────────
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
            # Detect specialist type from the user's raw problem description
            st.session_state.specialist_type = detect_specialist_type(st.session_state.problem_text)
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
            stype = st.session_state.specialist_type
            if stype == "natural_shading":
                prompt_text = NATURAL_SHADING_DATA_PROMPT.replace("USER_INPUT_PLACEHOLDER", user_input)
                max_tok = 6000
            elif stype == "green_roof":
                prompt_text = GREEN_ROOF_DATA_PROMPT.replace("USER_INPUT_PLACEHOLDER", user_input)
                max_tok = 6000
            else:
                prompt_text = GENERIC_DATA_PROMPT.replace("USER_INPUT_PLACEHOLDER", user_input)
                max_tok = 4000

            with st.spinner("Building analysis..."):
                history = [{"role": m["role"], "content": m["content"]} for m in st.session_state.messages]
                history.append({"role": "user", "content": prompt_text})
                resp = client.messages.create(model="claude-opus-4-6", max_tokens=max_tok, messages=history)

            raw = resp.content[0].text.strip()
            fence = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw)
            if fence:
                raw = fence.group(1).strip()

            try:
                data = json.loads(raw)
                st.session_state.analysis_data = data
                specialist_note = ""
                if data.get("specialist_type") == "natural_shading":
                    specialist_note = " using specialist VSL/CDD Natural Shading methodology (50-year horizon, 8-year maturity ramp)"
                elif data.get("specialist_type") == "green_roof":
                    specialist_note = " using specialist VSL/CDD Green Roof methodology (50-year horizon, property value uplift included)"
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": f"✅ Analysis complete for **{data['problem_title']}**. I've analyzed {len(data['measures'])} measure(s){specialist_note}. Click **Download Excel** below to get your full CBA model."
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
        bullets = (
            "- 📋 Inputs & Assumptions sheet (blue = editable)\n"
            "- 📊 CBA Results with NPV & BCR\n"
            "- 📈 Sensitivity Analysis (4-pillar)\n"
            "- 📁 Summary sheet"
        )
        if st.session_state.analysis_data.get("specialist_type"):
            bullets += "\n- 🔬 Specialist Detail (VSL derivation, year-by-year benefit table)"
        st.markdown(bullets)

    if st.button("🔄 Start New Analysis"):
        for k in ["messages", "stage", "analysis_data", "problem_text", "specialist_type"]:
            if k == "messages":
                st.session_state[k] = []
            elif k == "stage":
                st.session_state[k] = "problem"
            elif k == "specialist_type":
                st.session_state[k] = None
            else:
                st.session_state[k] = None if k == "analysis_data" else ""
        st.rerun()
