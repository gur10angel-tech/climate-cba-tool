import streamlit as st
import anthropic
import boto3
import json
import re
from excel_builder import build_excel
import tempfile, os
import markdown as md_lib


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
    {"name": "Discount Rate", "base": 0.035, "low": 0.02, "high": 0.07, "unit": "%"},
    {"name": "Annual Benefit Multiplier", "base": 1.0, "low": 0.6, "high": 1.4, "unit": "multiplier"},
    {"name": "CAPEX Variation", "base": 1.0, "low": 0.8, "high": 1.3, "unit": "multiplier"}
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
for k, v in {
    "messages": [], "stage": "problem", "analysis_data": None,
    "api_key": "", "problem_text": "", "specialist_type": None,
    "kb_id": "", "aws_key": "", "aws_secret": "", "aws_region": "eu-north-1"
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("## 🌍 Climate Adaptation CBA Tool")

# API key
with st.expander("⚙️ API Key", expanded=not st.session_state.api_key):
    k = st.text_input("Anthropic API Key", type="password", value=st.session_state.api_key)
    if k:
        sanitized = k.strip().encode("ascii", errors="ignore").decode("ascii")
        if sanitized != k.strip():
            st.warning("\u26a0\ufe0f Non-ASCII characters were removed from your API key. Please re-check it.")
        st.session_state.api_key = sanitized

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
                        kb_section += f"\n\n--- {label} ---\n{kb_text}"

                if kb_section:
                    kb_instruction = f"""
The following passages come DIRECTLY from peer-reviewed academic papers in the knowledge base.
Your proposed measures MUST be drawn from these papers — do not invent measures that are not mentioned.
Cite the paper name when proposing each measure.

{kb_section}

---"""
                else:
                    kb_instruction = ""

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

Be concise and cite sources."""
                    }]
                )
            reply = resp.content[0].text
            st.session_state.messages.append({"role": "assistant", "content": reply})
            st.session_state.specialist_type = detect_specialist_type(st.session_state.problem_text)
            st.session_state.stage = "measures"
            st.rerun()

        # ── Stage: user selects measures ───────────────────────────────────────
        elif st.session_state.stage == "measures":
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
                        kb_section += f"\n\n--- {label} ---\n{kb_text}"

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

                history = [{"role": m["role"], "content": m["content"]} for m in st.session_state.messages]
                history.append({
                    "role": "user",
                    "content": f"""The user selected these measures: "{user_input}"
{kb_block}
For each selected measure:

1. State the benefit formula clearly (e.g. "Avoided mortality = Population × mortality_rate × heat_mortality_factor × efficiency × VSL").
   Cite the source paper for each parameter value you will use as a default.

2. Split data needs into two explicit groups:
   **I need from you (local data):** e.g. population size, project area, local electricity price
   **I will use from literature (with source):** e.g. VSL = €3.8M (OECD 2012), heat_mortality_factor = 0.00083 (Gasparrini 2017)

3. Ask the user ONLY for the local data. Keep questions short and grouped by measure.

If the user has already said "use defaults", skip the questions and confirm you will use literature defaults for everything."""
                })
                resp = client.messages.create(model="claude-opus-4-6", max_tokens=1500, messages=history)

            reply = resp.content[0].text
            st.session_state.messages.append({"role": "assistant", "content": reply})
            st.session_state.stage = "data"
            st.rerun()

        # ── Stage: user provides data → build analysis ─────────────────────────
        elif st.session_state.stage == "data":
            stype = st.session_state.specialist_type
            if stype == "natural_shading":
                prompt_text = NATURAL_SHADING_DATA_PROMPT.replace("USER_INPUT_PLACEHOLDER", user_input)
                max_tok = 7000
            elif stype == "green_roof":
                prompt_text = GREEN_ROOF_DATA_PROMPT.replace("USER_INPUT_PLACEHOLDER", user_input)
                max_tok = 7000
            else:
                prompt_text = GENERIC_DATA_PROMPT.replace("USER_INPUT_PLACEHOLDER", user_input)
                max_tok = 6000

            # Query KB for formula parameters from literature
            kb_cba = query_bedrock_kb(
                f"VSL mortality rate heat reduction efficiency cost per unit {st.session_state.problem_text}"
            )
            if kb_cba and not kb_cba.startswith("[KB unavailable"):
                prompt_text += f"""

LITERATURE PARAMETER VALUES — use these to populate formula fields and their _source fields:

{kb_cba}

---"""

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

            with st.spinner("Building analysis from literature + your data..."):
                history = [{"role": m["role"], "content": m["content"]} for m in st.session_state.messages]
                history.append({"role": "user", "content": prompt_text})
                resp = client.messages.create(
                    model="claude-opus-4-6", max_tokens=max_tok, messages=history
                )

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
                    resp2 = client.messages.create(
                        model="claude-opus-4-6", max_tokens=max_tok, messages=retry_messages
                    )
                    data = _extract_json(resp2.content[0].text)

            if data:
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
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": (
                        "I wasn't able to generate a valid analysis. "
                        "Please try again or type **'use defaults'** to let me fill everything from literature."
                    )
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
