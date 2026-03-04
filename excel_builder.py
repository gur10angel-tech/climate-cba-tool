from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

C_DARK   = "1A1A2E"
C_MID    = "16213E"
C_ACCENT = "3B82F6"
C_GREEN  = "059669"
C_RED    = "DC2626"
C_AMBER  = "D97706"
C_LIGHT  = "EFF6FF"
C_WHITE  = "FFFFFF"
C_BORDER = "E2E8F0"
BLUE     = "0000FF"
BLACK    = "000000"
GREEN_LK = "008000"


def _bd():
    s = Side(style="thin", color=C_BORDER)
    return Border(left=s, right=s, top=s, bottom=s)

def _hdr(ws, r, c, v, bg=C_DARK, fg=C_WHITE, bold=True, sz=11, wrap=False, span=1):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(name="Arial", bold=bold, color=fg, size=sz)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    if span > 1:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c+span-1)
    return cell

def _cell(ws, r, c, v, bold=False, color=BLACK, fmt=None, bg=None, align="left"):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(name="Arial", bold=bold, color=color, size=10)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = _bd()
    if fmt: cell.number_format = fmt
    if bg:  cell.fill = PatternFill("solid", fgColor=bg)
    return cell

def _sec(ws, r, c, label, span=6):
    cell = ws.cell(row=r, column=c, value=label)
    cell.font = Font(name="Arial", bold=True, color=C_WHITE, size=10)
    cell.fill = PatternFill("solid", fgColor=C_MID)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c+span-1)
    ws.row_dimensions[r].height = 18

def _widths(ws, d):
    for col, w in d.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def build_excel(data: dict, path: str):
    wb = Workbook()
    rm = _inputs(wb, data)
    _results(wb, data, rm)
    _sensitivity(wb, data, rm)
    _summary(wb, data, rm)
    if data.get("specialist_type") in ("natural_shading", "green_roof"):
        _specialist_detail(wb, data, rm)
        _benefit_breakdown(wb, data, rm)
    wb.save(path)


# ── SHEET 1: INPUTS ────────────────────────────────────────────────────────────
def _inputs(wb, data):
    ws = wb.active
    ws.title = "Inputs"
    ws.sheet_view.showGridLines = False

    measures = data["measures"]
    n = len(measures)
    cur = f"{data['currency']} ({data['currency_unit']})"

    _hdr(ws, 1, 1, f"CLIMATE ADAPTATION CBA — {data['problem_title'].upper()}", sz=13, span=n+4)
    _hdr(ws, 2, 1, data["problem_summary"], bg=C_MID, fg="CBD5E1", bold=False, sz=9, wrap=True, span=n+4)
    ws.row_dimensions[2].height = 36

    row = 4
    _sec(ws, row, 1, "GLOBAL PARAMETERS", span=4); row += 1

    _cell(ws, row, 1, "Discount Rate", bold=True)
    c = ws.cell(row=row, column=2, value=data["discount_rate"])
    c.font = Font(name="Arial", bold=True, color=BLUE, size=10)
    c.number_format = "0.00%"; c.border = _bd()
    c.fill = PatternFill("solid", fgColor=C_LIGHT)
    c.alignment = Alignment(horizontal="right")
    _cell(ws, row, 3, "← change here to update all calculations", color="94A3B8")
    DR_ROW = row; row += 1

    _cell(ws, row, 1, "Time Horizon (years)", bold=True)
    c = ws.cell(row=row, column=2, value=data["time_horizon"])
    c.font = Font(name="Arial", bold=True, color=BLUE, size=10)
    c.number_format = "#,##0"; c.border = _bd()
    c.fill = PatternFill("solid", fgColor=C_LIGHT)
    c.alignment = Alignment(horizontal="right")
    _cell(ws, row, 3, "← change here to update all calculations", color="94A3B8")
    YR_ROW = row; row += 1

    _cell(ws, row, 1, "Currency", bold=True)
    _cell(ws, row, 2, cur)
    row += 2

    _sec(ws, row, 1, f"MEASURE INPUTS  [{cur}]  —  Blue cells are editable", span=n+3); row += 1

    _hdr(ws, row, 1, "Parameter", bg=C_ACCENT, sz=10)
    for ci, m in enumerate(measures, 2):
        _hdr(ws, row, ci, m["name"], bg=C_ACCENT, sz=10)
    _hdr(ws, row, n+2, "Notes", bg=C_ACCENT, sz=10)
    row += 1

    CAPEX_ROW = row
    _cell(ws, row, 1, "Capital Cost / CAPEX", bold=True)
    for ci, m in enumerate(measures, 2):
        c = ws.cell(row=row, column=ci, value=m["capex"])
        c.font = Font(name="Arial", bold=True, color=BLUE, size=10)
        c.number_format = "#,##0.0"; c.border = _bd()
        c.fill = PatternFill("solid", fgColor=C_LIGHT)
        c.alignment = Alignment(horizontal="right")
    _cell(ws, row, n+2, "One-time capital cost"); row += 1

    OPEX_ROW = row
    _cell(ws, row, 1, "Annual O&M Cost / OPEX", bold=True)
    for ci, m in enumerate(measures, 2):
        c = ws.cell(row=row, column=ci, value=m["annual_opex"])
        c.font = Font(name="Arial", bold=True, color=BLUE, size=10)
        c.number_format = "#,##0.0"; c.border = _bd()
        c.fill = PatternFill("solid", fgColor=C_LIGHT)
        c.alignment = Alignment(horizontal="right")
    _cell(ws, row, n+2, "Recurring annual cost"); row += 1

    BENEFIT_ROW = row
    _cell(ws, row, 1, "Annual Benefit", bold=True)
    for ci, m in enumerate(measures, 2):
        c = ws.cell(row=row, column=ci, value=m["annual_benefit"])
        c.font = Font(name="Arial", bold=True, color=BLUE, size=10)
        c.number_format = "#,##0.0"; c.border = _bd()
        c.fill = PatternFill("solid", fgColor=C_LIGHT)
        c.alignment = Alignment(horizontal="right")
    _cell(ws, row, n+2, "Annual monetised benefit"); row += 1

    LIFE_ROW = row
    _cell(ws, row, 1, "Measure Lifetime (years)", bold=True)
    for ci, m in enumerate(measures, 2):
        c = ws.cell(row=row, column=ci, value=m["lifetime_years"])
        c.font = Font(name="Arial", bold=True, color=BLUE, size=10)
        c.number_format = "#,##0"; c.border = _bd()
        c.fill = PatternFill("solid", fgColor=C_LIGHT)
        c.alignment = Alignment(horizontal="right")
    _cell(ws, row, n+2, "Used for PV calculations"); row += 1

    for label, key in [("Category", "category"), ("Feasibility", "feasibility"),
                       ("Uncertainty", "uncertainty"), ("Data Source", "data_source")]:
        _cell(ws, row, 1, label, bold=True)
        for ci, m in enumerate(measures, 2):
            _cell(ws, row, ci, m[key], align="center")
        row += 1

    _cell(ws, row, 1, "Benefit Types", bold=True)
    for ci, m in enumerate(measures, 2):
        _cell(ws, row, ci, ", ".join(m["benefit_types"]), align="center")
    row += 1

    _cell(ws, row, 1, "Co-benefits", bold=True)
    for ci, m in enumerate(measures, 2):
        _cell(ws, row, ci, m["co_benefits"])
    row += 2

    _sec(ws, row, 1, "KEY ASSUMPTIONS & LIMITATIONS", span=n+3); row += 1
    c = ws.cell(row=row, column=1, value=data.get("key_assumptions", ""))
    c.font = Font(name="Arial", size=9, color="475569")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row+2, end_column=n+3)
    ws.row_dimensions[row].height = 48; row += 3

    _cell(ws, row, 1, "Data Gaps:", bold=True)
    c = ws.cell(row=row, column=2, value=data.get("data_gaps", ""))
    c.font = Font(name="Arial", size=9, color="475569")
    c.alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=n+3); row += 2

    _sec(ws, row, 1, "COLOUR LEGEND", span=4); row += 1
    for txt, col in [("Blue = hardcoded input (editable)", BLUE),
                     ("Black = formula (do not overwrite)", BLACK),
                     ("Green = linked from another sheet (do not overwrite)", GREEN_LK)]:
        c = ws.cell(row=row, column=1, value=txt)
        c.font = Font(name="Arial", color=col, size=9)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        row += 1

    _widths(ws, {1: 30, **{i+2: 22 for i in range(n)}, n+2: 28})

    return {"dr": DR_ROW, "yr": YR_ROW,
            "capex": CAPEX_ROW, "opex": OPEX_ROW,
            "benefit": BENEFIT_ROW, "life": LIFE_ROW, "n": n}


# ── SHEET 2: CBA RESULTS ───────────────────────────────────────────────────────
def _results(wb, data, rm):
    ws = wb.create_sheet("CBA Results")
    ws.sheet_view.showGridLines = False
    measures = data["measures"]
    n = rm["n"]

    _hdr(ws, 1, 1, "COST-BENEFIT ANALYSIS — RESULTS", sz=13, span=n+3)
    _hdr(ws, 2, 1,
         f"All values in {data['currency']} {data['currency_unit']}  |  Change blue cells in Inputs sheet to update",
         bg=C_MID, fg="CBD5E1", bold=False, sz=9, span=n+3)

    row = 4
    _hdr(ws, row, 1, "Metric", bg=C_ACCENT, sz=10)
    for ci, m in enumerate(measures, 2):
        _hdr(ws, row, ci, m["name"], bg=C_ACCENT, sz=10)
    _hdr(ws, row, n+2, "Best", bg=C_ACCENT, sz=10)
    row += 1

    dr   = f"Inputs!$B${rm['dr']}"
    # Helper references
    def inp(key, ci):
        return f"Inputs!{get_column_letter(ci)}${rm[key]}"

    rr = {}  # track row numbers

    def add_row(label, key, f_fn, fmt, is_key=False):
        nonlocal row
        bg = C_LIGHT if is_key else None
        _cell(ws, row, 1, label, bold=is_key, bg=bg)
        for ci in range(2, n+2):
            f = f_fn(ci)
            c = ws.cell(row=row, column=ci, value=f)
            is_inp_link = f.startswith("=Inputs")
            c.font = Font(name="Arial", color=GREEN_LK if is_inp_link else BLACK,
                          bold=is_key, size=10)
            c.number_format = fmt; c.border = _bd()
            c.alignment = Alignment(horizontal="right", vertical="center")
            if bg: c.fill = PatternFill("solid", fgColor=C_LIGHT)
        rr[key] = row; row += 1

    add_row(f"CAPEX ({data['currency']}M)", "capex",
            lambda ci: f"=Inputs!{get_column_letter(ci)}${rm['capex']}", "#,##0.0")
    add_row(f"Annual Benefit ({data['currency']}M)", "ann_ben",
            lambda ci: f"=Inputs!{get_column_letter(ci)}${rm['benefit']}", "#,##0.0")
    add_row(f"Annual OPEX ({data['currency']}M)", "ann_opex",
            lambda ci: f"=Inputs!{get_column_letter(ci)}${rm['opex']}", "#,##0.0")

    def pv_ben_f(ci):
        b = inp("benefit", ci); l = inp("life", ci)
        return f"={b}/{dr}*(1-(1+{dr})^(-{l}))"

    def pv_cost_f(ci):
        k = inp("capex", ci); o = inp("opex", ci); l = inp("life", ci)
        return f"={k}+{o}/{dr}*(1-(1+{dr})^(-{l}))"

    add_row(f"PV of Benefits ({data['currency']}M)", "pv_ben", pv_ben_f, "#,##0.0")
    PVB = rr["pv_ben"]
    add_row(f"PV of Costs ({data['currency']}M)", "pv_cost", pv_cost_f, "#,##0.0")
    PVC = rr["pv_cost"]

    def npv_f(ci):
        col = get_column_letter(ci)
        return f"={col}{PVB}-{col}{PVC}"

    add_row(f"NPV ({data['currency']}M)", "npv", npv_f,
            '#,##0.0;(#,##0.0);"-"', is_key=True)
    NPV_ROW = rr["npv"]

    def bcr_f(ci):
        col = get_column_letter(ci)
        return f"=IF({col}{PVC}>0,{col}{PVB}/{col}{PVC},0)"

    add_row("BCR", "bcr", bcr_f, "0.00", is_key=True)
    BCR_ROW = rr["bcr"]

    def payback_f(ci):
        k = inp("capex", ci); b = inp("benefit", ci); o = inp("opex", ci)
        return f'=IF(({b}-{o})>0,{k}/({b}-{o}),"N/A")'

    add_row("Simple Payback (years)", "payback", payback_f, "0.0")

    def ann_net_f(ci):
        b = inp("benefit", ci); o = inp("opex", ci)
        return f"={b}-{o}"

    add_row(f"Annual Net Benefit ({data['currency']}M)", "ann_net", ann_net_f, "#,##0.0")

    # Best column formulas
    bcr_range = f"{get_column_letter(2)}{BCR_ROW}:{get_column_letter(n+1)}{BCR_ROW}"
    npv_range = f"{get_column_letter(2)}{NPV_ROW}:{get_column_letter(n+1)}{NPV_ROW}"

    for key_r, rng, fmt in [(BCR_ROW, bcr_range, "0.00"), (NPV_ROW, npv_range, "#,##0.0")]:
        c = ws.cell(row=key_r, column=n+2, value=f"=MAX({rng})")
        c.font = Font(name="Arial", color=C_GREEN, bold=True, size=10)
        c.number_format = fmt; c.border = _bd()
        c.fill = PatternFill("solid", fgColor=C_LIGHT)
        c.alignment = Alignment(horizontal="right")

    # Best measure name (above BCR row)
    name_range = f"{get_column_letter(2)}4:{get_column_letter(n+1)}4"
    c = ws.cell(row=4, column=n+2,
                value=f'=INDEX({name_range},MATCH(MAX({bcr_range}),{bcr_range},0))')
    c.font = Font(name="Arial", color=C_GREEN, bold=True, size=10)
    c.fill = PatternFill("solid", fgColor=C_DARK)
    c.alignment = Alignment(horizontal="center")

    # Ranking & Viability
    row += 1
    _sec(ws, row, 1, "RANKING & VIABILITY", span=n+2); row += 1

    _cell(ws, row, 1, "BCR Rank (1 = best)", bold=True)
    for ci in range(2, n+2):
        col = get_column_letter(ci)
        c = ws.cell(row=row, column=ci,
                    value=f"=RANK({col}{BCR_ROW},{bcr_range},0)")
        c.font = Font(name="Arial", color=BLACK, bold=True, size=10)
        c.number_format = "0"; c.border = _bd()
        c.alignment = Alignment(horizontal="center")
    row += 1

    _cell(ws, row, 1, "Viable? (BCR ≥ 1)", bold=True)
    for ci in range(2, n+2):
        col = get_column_letter(ci)
        c = ws.cell(row=row, column=ci,
                    value=f'=IF({col}{BCR_ROW}>=1,"✓ Yes","✗ No")')
        c.font = Font(name="Arial", color=BLACK, size=10)
        c.border = _bd()
        c.alignment = Alignment(horizontal="center")
    row += 2

    # Advanced benefit breakdown (specialist modes only)
    BENEFIT_LABELS = {
        "avoided_mortality_npv":      "Avoided Mortality — NPV",
        "morbidity_savings_npv":      "Morbidity Savings — NPV",
        "skin_cancer_prevention_npv": "Skin Cancer Prevention — NPV",
        "carbon_sequestration_npv":   "Carbon Sequestration — NPV",
        "runoff_reduction_npv":       "Runoff Reduction — NPV",
        "air_quality_npv":            "Air Quality — NPV",
        "habitat_creation_npv":       "Habitat Creation — NPV",
        "property_value_uplift_npv":  "Property Value Uplift — NPV",
        "roof_longevity_npv":        "Roof Longevity Extension — NPV",
    }
    adv_measures = [m for m in measures if m.get("advanced_benefits")]
    if adv_measures:
        _sec(ws, row, 1, "ADVANCED BENEFIT BREAKDOWN  (Specialist Methodology)", span=n+2); row += 1
        all_adv_keys = {k for m in adv_measures for k in m.get("advanced_benefits", {})}
        for key, label in BENEFIT_LABELS.items():
            if key not in all_adv_keys:
                continue
            _cell(ws, row, 1, label, bold=False, bg="F0FDF4")
            for ci, m in enumerate(measures, 2):
                val = (m.get("advanced_benefits") or {}).get(key, 0) or 0
                c = ws.cell(row=row, column=ci, value=val)
                c.font = Font(name="Arial", color=BLACK, size=10)
                c.number_format = "#,##0.0"; c.border = _bd()
                c.alignment = Alignment(horizontal="right")
                c.fill = PatternFill("solid", fgColor="F0FDF4")
            row += 1
        row += 1

    # Chart data area
    chart_r = row
    ws.cell(row=chart_r, column=1, value="Measure")
    ws.cell(row=chart_r, column=2, value="BCR")
    for mi, m in enumerate(measures, 1):
        ws.cell(row=chart_r+mi, column=1, value=m["name"])
        ws.cell(row=chart_r+mi, column=2,
                value=f"={get_column_letter(mi+1)}{BCR_ROW}")

    chart = BarChart()
    chart.type = "col"; chart.title = "BCR by Measure"
    chart.y_axis.title = "Benefit-Cost Ratio"
    chart.style = 10; chart.width = 18; chart.height = 12

    data_ref = Reference(ws, min_col=2, max_col=2, min_row=chart_r, max_row=chart_r+n)
    cats_ref = Reference(ws, min_col=1, min_row=chart_r+1, max_row=chart_r+n)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws.add_chart(chart, f"A{chart_r+n+3}")

    _widths(ws, {1: 32, **{i+2: 22 for i in range(n)}, n+2: 20})

    # Expose BCR_ROW for summary sheet
    rm["results_bcr_row"]   = BCR_ROW
    rm["results_capex_row"] = rr["capex"]
    rm["results_npv_row"]   = NPV_ROW


def _build_specialist_sensitivity_vars(data):
    """Return the 4-pillar sensitivity variable list for specialist measure types."""
    sp  = data.get("specialist_params", {})
    dr  = data.get("discount_rate", 0.035)
    eff = sp.get("heat_reduction_efficiency", 0.4)
    return [
        {"name": "Discount Rate",            "low": 0.02,             "base": dr,  "high": 0.10,            "unit": "%"},
        {"name": "Heat Reduction Efficiency","low": round(eff*0.7,3), "base": eff, "high": round(eff*1.3,3),"unit": "multiplier"},
        {"name": "CAPEX Variation",          "low": 0.80,             "base": 1.00,"high": 1.30,            "unit": "multiplier"},
        {"name": "Activity Level",           "low": 0.70,             "base": 1.00,"high": 1.30,            "unit": "multiplier"},
    ]


# ── SHEET 3: SENSITIVITY ──────────────────────────────────────────────────────
def _sensitivity(wb, data, rm):
    ws = wb.create_sheet("Sensitivity")
    ws.sheet_view.showGridLines = False
    measures = data["measures"]
    n = rm["n"]
    sens_vars = data.get("sensitivity_vars", [])
    if data.get("specialist_type") in ("natural_shading", "green_roof"):
        sens_vars = _build_specialist_sensitivity_vars(data)

    _hdr(ws, 1, 1, "SENSITIVITY ANALYSIS", sz=13, span=n+5)
    _hdr(ws, 2, 1,
         "Edit Low / Base / High values (blue cells). BCR table below recalculates automatically via formulas.",
         bg=C_MID, fg="CBD5E1", bold=False, sz=9, span=n+5)

    row = 4
    _sec(ws, row, 1, "PARAMETER RANGES  (edit blue cells)", span=5); row += 1
    for ci, h in enumerate(["Variable", "Low", "Base", "High", "Unit"], 1):
        _hdr(ws, row, ci, h, bg=C_ACCENT, sz=10)
    row += 1

    sv_rows = {}
    for sv in sens_vars:
        _cell(ws, row, 1, sv["name"], bold=True)
        for col_i, val in enumerate([sv["low"], sv["base"], sv["high"]], 2):
            c = ws.cell(row=row, column=col_i, value=val)
            c.font = Font(name="Arial", bold=True, color=BLUE, size=10)
            c.number_format = "0.00%" if sv["unit"] == "%" else "0.00"
            c.border = _bd()
            c.fill = PatternFill("solid", fgColor=C_LIGHT)
            c.alignment = Alignment(horizontal="right")
        _cell(ws, row, 5, sv["unit"])
        sv_rows[sv["name"]] = {
            "low":  f"$B${row}",
            "base": f"$C${row}",
            "high": f"$D${row}"
        }
        row += 1

    row += 1
    _sec(ws, row, 1, "BCR SENSITIVITY TABLE  (formulas — updates automatically)", span=n+3); row += 1
    _hdr(ws, row, 1, "Variable", bg=C_ACCENT, sz=10)
    _hdr(ws, row, 2, "Scenario", bg=C_ACCENT, sz=10)
    for ci, m in enumerate(measures, 3):
        _hdr(ws, row, ci, m["name"], bg=C_ACCENT, sz=10)
    row += 1

    dr_inp = f"Inputs!$B${rm['dr']}"

    for sv in sens_vars:
        sv_name = sv["name"]
        sv_r = sv_rows.get(sv_name, {})

        for scenario, col_key, color in [("Low", "low", C_RED),
                                          ("Base", "base", BLACK),
                                          ("High", "high", C_GREEN)]:
            param_ref = sv_r.get(col_key, "$C$7")
            bg = C_LIGHT if scenario == "Base" else None

            _cell(ws, row, 1, sv_name if scenario == "Low" else "", bold=(scenario == "Low"))
            _cell(ws, row, 2, scenario, color=color, bg=bg)

            for mi in range(n):
                ci = mi + 2  # column in Inputs sheet (measures start at col 2)
                capex = f"Inputs!{get_column_letter(ci)}${rm['capex']}"
                opex  = f"Inputs!{get_column_letter(ci)}${rm['opex']}"
                ben   = f"Inputs!{get_column_letter(ci)}${rm['benefit']}"
                life  = f"Inputs!{get_column_letter(ci)}${rm['life']}"

                pv_cost_override = None
                if sv_name == "Discount Rate":
                    dr_use = param_ref
                    ben_term = ben
                elif sv_name == "Annual Benefit":
                    dr_use = dr_inp
                    ben_term = f"({ben}*{param_ref})"
                elif sv_name == "CAPEX Variation":
                    dr_use = dr_inp
                    ben_term = ben
                    pv_cost_override = f"(({capex}*{param_ref})+{opex}/{dr_use}*(1-(1+{dr_use})^(-{life})))"
                elif sv_name == "Activity Level":
                    dr_use = dr_inp
                    ben_term = f"({ben}*{param_ref})"
                elif sv_name == "Heat Reduction Efficiency":
                    dr_use = dr_inp
                    base_eff = data.get("specialist_params", {}).get("heat_reduction_efficiency", 0.5)
                    ben_term = f"({ben}*{param_ref}/{base_eff})"
                else:
                    dr_use = dr_inp
                    ben_term = ben

                pv_ben  = f"({ben_term}/{dr_use}*(1-(1+{dr_use})^(-{life})))"
                pv_cost = pv_cost_override if pv_cost_override else \
                          f"({capex}+{opex}/{dr_use}*(1-(1+{dr_use})^(-{life})))"
                formula = f"=IF({dr_use}>0,IF({pv_cost}>0,{pv_ben}/{pv_cost},0),0)"

                c = ws.cell(row=row, column=mi+3, value=formula)
                c.font = Font(name="Arial", color=color, bold=(scenario == "Base"), size=10)
                c.number_format = "0.00"; c.border = _bd()
                c.alignment = Alignment(horizontal="right")
                if bg: c.fill = PatternFill("solid", fgColor=C_LIGHT)
            row += 1
        row += 1  # blank between variables

    row += 1
    _sec(ws, row, 1, "INSTRUCTIONS", span=6); row += 1
    for inst in [
        "1. Edit Low / Base / High values in the table above (blue cells only)",
        "2. BCR values in the table update automatically — no manual recalculation needed",
        "3. 'Discount Rate': enter as decimal (e.g. 0.02 = 2%)",
        "4. 'Annual Benefit' multiplier: 1.0 = base, 0.5 = 50% of base benefits, 1.5 = 150%",
        "5. Add more sensitivity variables by editing the app and rerunning the analysis",
    ]:
        c = ws.cell(row=row, column=1, value=inst)
        c.font = Font(name="Arial", size=9, color="475569")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1

    _widths(ws, {1: 26, 2: 12, 3: 14, 4: 14, 5: 12, **{i+5: 22 for i in range(n)}})


# ── SHEET 4: SUMMARY ──────────────────────────────────────────────────────────
def _summary(wb, data, rm):
    ws = wb.create_sheet("Summary")
    ws.sheet_view.showGridLines = False
    measures = data["measures"]
    n = rm["n"]
    BCR_ROW   = rm["results_bcr_row"]
    CAPEX_ROW = rm["results_capex_row"]
    NPV_ROW   = rm["results_npv_row"]

    _hdr(ws, 1, 1, "EXECUTIVE SUMMARY", sz=13, span=5)
    _hdr(ws, 2, 1, data["problem_title"], bg=C_MID, fg="CBD5E1", sz=11, span=5)

    row = 4
    _sec(ws, row, 1, "PROBLEM OVERVIEW", span=5); row += 1
    c = ws.cell(row=row, column=1, value=data["problem_summary"])
    c.font = Font(name="Arial", size=10)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row+1, end_column=5)
    ws.row_dimensions[row].height = 40; row += 3

    _sec(ws, row, 1, "MEASURES AT A GLANCE  (values link live from CBA Results)", span=5); row += 1
    for ci, h in enumerate(["Measure", "Category", f"CAPEX ({data['currency']}M)",
                             "BCR (base)", "Viability"], 1):
        _hdr(ws, row, ci, h, bg=C_ACCENT, sz=10)
    row += 1

    for mi, m in enumerate(measures):
        ci = mi + 2  # Excel column in Results sheet
        col = get_column_letter(ci)

        _cell(ws, row, 1, m["name"], bold=True)
        _cell(ws, row, 2, m["category"], align="center")

        c = ws.cell(row=row, column=3, value=f"='CBA Results'!{col}{CAPEX_ROW}")
        c.font = Font(name="Arial", color=GREEN_LK, size=10)
        c.number_format = "#,##0.0"; c.border = _bd()
        c.alignment = Alignment(horizontal="right")

        c = ws.cell(row=row, column=4, value=f"='CBA Results'!{col}{BCR_ROW}")
        c.font = Font(name="Arial", color=GREEN_LK, bold=True, size=10)
        c.number_format = "0.00"; c.border = _bd()
        c.fill = PatternFill("solid", fgColor=C_LIGHT)
        c.alignment = Alignment(horizontal="center")

        bcr_here = f"D{row}"
        c = ws.cell(row=row, column=5,
                    value=f'=IF({bcr_here}>=1.5,"✓ Recommended",IF({bcr_here}>=1.0,"○ Consider","✗ Review"))')
        c.font = Font(name="Arial", color=BLACK, size=10)
        c.border = _bd(); c.alignment = Alignment(horizontal="center")
        row += 1

    row += 1
    _sec(ws, row, 1, "HOW TO USE THIS WORKBOOK", span=5); row += 1
    for step in [
        "1.  INPUTS sheet — Edit blue cells (CAPEX, OPEX, Benefits, Discount Rate, Time Horizon)",
        "2.  CBA RESULTS sheet — NPV and BCR recalculate automatically from Inputs",
        "3.  BENEFIT BREAKDOWN sheet — Shows NPV of each benefit category (mortality, morbidity, skin cancer, ecosystem services, property uplift, roof longevity) and their shares of total benefits",
        "4.  SENSITIVITY sheet — Edit Low/Base/High ranges; BCR table updates automatically",
        "5.  SPECIALIST DETAIL sheet — Inspect VSL derivation, CDD parameters, specialist parameters, and year-by-year benefit projections",
        "6.  SUMMARY sheet (this sheet) — BCR and CAPEX values link live from CBA Results; use together with Benefit Breakdown and Specialist Detail when presenting results",
        "7.  Never overwrite black or green cells — they contain formulas",
    ]:
        c = ws.cell(row=row, column=1, value=step)
        c.font = Font(name="Arial", size=10)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        row += 1

    _widths(ws, {1: 30, 2: 18, 3: 18, 4: 16, 5: 18})


# ── SHEET 5: SPECIALIST DETAIL ────────────────────────────────────────────────
def _specialist_detail(wb, data, rm):
    ws = wb.create_sheet("Specialist Detail")
    ws.sheet_view.showGridLines = False

    stype    = data.get("specialist_type", "")
    vsl      = data.get("vsl_params", {})
    cdd      = data.get("cdd_params", {})
    sp       = data.get("specialist_params", {})
    measures = data["measures"]
    n        = rm["n"]
    cur      = data.get("currency", "NIS")

    stype_label = "Natural Shading (Boulevard Trees)" if stype == "natural_shading" else "Green Roof"

    # ── Row 1-2: Header ────────────────────────────────────────────────────────
    _hdr(ws, 1, 1, f"SPECIALIST METHODOLOGY DETAIL — {stype_label.upper()}", sz=13, span=10)
    _hdr(ws, 2, 1,
         f"Functional unit: {sp.get('functional_unit','—')}  |  {data.get('problem_title','')}",
         bg=C_MID, fg="CBD5E1", bold=False, sz=9, span=10)

    row = 4

    # ── Section: VSL Derivation ────────────────────────────────────────────────
    _sec(ws, row, 1, "VALUE OF STATISTICAL LIFE — STEP-BY-STEP DERIVATION", span=4); row += 1
    for ci, h in enumerate(["Step", "Parameter", "Value", "Notes"], 1):
        _hdr(ws, row, ci, h, bg=C_ACCENT, sz=10)
    row += 1

    vsl_base     = vsl.get("base_vsl_usd_2005", 3_000_000)
    cpi_mult     = vsl.get("cpi_multiplier", 1.68)
    ppp_ratio    = vsl.get("gdp_ppp_ratio", 0.89)
    income_el    = vsl.get("income_elasticity", 1.0)
    fx           = vsl.get("usd_to_local_currency", 3.7)
    life_exp     = vsl.get("life_expectancy_remaining", 35)
    vsl_local    = vsl.get("computed_vsl_local", 12_800_000)
    vsly_local   = vsl.get("computed_vsly_local", 365_714)
    cpi_adj_vsl  = round(vsl_base * cpi_mult)
    ppp_adj_vsl_usd = round(cpi_adj_vsl * ppp_ratio * income_el)

    VSL_ROWS = [
        ("1", "Base VSL (OECD, 2005 USD)",       vsl_base,           "#,##0",   "OECD meta-study baseline"),
        ("2", f"× CPI Multiplier (2005→2024)",    cpi_mult,           "0.00",    "US Bureau of Labor Statistics"),
        ("3", "= CPI-Adjusted VSL (2024 USD)",    cpi_adj_vsl,        "#,##0",   ""),
        ("4", f"× GDP PPP Ratio (Israel / OECD)", ppp_ratio,          "0.000",   "World Bank WDI"),
        ("5", "× Income Elasticity",              income_el,          "0.0",     "Standard for developed economies"),
        ("6", "= PPP-Adjusted VSL (USD)",         ppp_adj_vsl_usd,    "#,##0",   ""),
        ("7", f"× Exchange Rate ({cur}/USD)",     fx,                 "0.00",    ""),
        ("8", f"= VSL in {cur}",                  vsl_local,          "#,##0",   "Used in benefit calculations"),
        ("9", "÷ Remaining Life Expectancy (yrs)",life_exp,           "0",       "Affected demographic"),
        ("10",f"= VSLY in {cur}",                 vsly_local,         "#,##0",   "Value per Statistical Life Year"),
    ]
    for step, param, val, fmt, note in VSL_ROWS:
        _cell(ws, row, 1, step, align="center")
        _cell(ws, row, 2, param, bold=(step in ("8","10")))
        c = ws.cell(row=row, column=3, value=val)
        c.font = Font(name="Arial",
                      bold=(step in ("8","10")),
                      color=C_GREEN if step in ("8","10") else BLACK, size=10)
        c.number_format = fmt; c.border = _bd()
        c.alignment = Alignment(horizontal="right")
        if step in ("8","10"):
            c.fill = PatternFill("solid", fgColor=C_LIGHT)
        _cell(ws, row, 4, note, color="94A3B8")
        row += 1

    row += 1

    # ── Section: CDD Parameters ────────────────────────────────────────────────
    _sec(ws, row, 1, "COOLING DEGREE DAY PARAMETERS", span=4); row += 1
    for ci, h in enumerate(["Parameter", "Value", "Unit", "Notes"], 1):
        _hdr(ws, row, ci, h, bg=C_ACCENT, sz=10)
    row += 1

    pop_label = "Pedestrians per Hour" if stype == "natural_shading" else "Population Density (people/km²)"
    CDD_ROWS = [
        ("Annual Cooling Degree Days",   cdd.get("annual_cdd", 735),          "CDD",        "Tel Aviv baseline (21°C base)"),
        ("Base Temperature",             cdd.get("base_temp_celsius", 21),     "°C",         "Threshold for heat health risk"),
        ("Heat-Mortality Factor",        cdd.get("heat_mortality_factor", 0.00083), "deaths/person", "Per CDD above threshold"),
        (pop_label,                      cdd.get("population_density_or_pedestrians", "—"), "people", "Population at risk driver"),
    ]
    for param, val, unit, note in CDD_ROWS:
        _cell(ws, row, 1, param, bold=True)
        c = ws.cell(row=row, column=2, value=val)
        c.font = Font(name="Arial", color=BLACK, size=10)
        c.number_format = "0.00000" if isinstance(val, float) and val < 0.01 else "#,##0.##"
        c.border = _bd(); c.alignment = Alignment(horizontal="right")
        _cell(ws, row, 3, unit, align="center")
        _cell(ws, row, 4, note, color="94A3B8")
        row += 1

    row += 1

    # ── Section: Specialist Parameters ────────────────────────────────────────
    _sec(ws, row, 1, "SPECIALIST PARAMETERS", span=4); row += 1
    for ci, h in enumerate(["Parameter", "Value", "Unit"], 1):
        _hdr(ws, row, ci, h, bg=C_ACCENT, sz=10)
    row += 1

    if stype == "natural_shading":
        SP_ROWS = [
            ("Heat Reduction Efficiency", sp.get("heat_reduction_efficiency", 0.50), "ratio", True),
            ("UV Reduction Factor",       sp.get("uv_reduction_factor", 0.75),      "ratio", False),
            ("Pedestrians per Hour",      sp.get("pedestrians_per_hour", 1200),      "pax/hr", False),
            ("Functional Unit",           sp.get("functional_unit", "linear meter"), "—",    False),
        ]
    else:
        SP_ROWS = [
            ("Heat Reduction Efficiency", sp.get("heat_reduction_efficiency", 0.28), "ratio",  True),
            ("Property Value Uplift %",   sp.get("property_value_uplift_pct", 0.03), "%",      False),
            ("Roof Longevity Extension",  sp.get("roof_longevity_extension_years", 15), "years", False),
            ("Roof Area",                 sp.get("roof_area_m2", 1000),               "m²",    False),
            ("Functional Unit",           sp.get("functional_unit", "sq meter"),      "—",     False),
        ]

    for param, val, unit, is_editable in SP_ROWS:
        _cell(ws, row, 1, param, bold=True)
        c = ws.cell(row=row, column=2, value=val)
        if is_editable:
            c.font = Font(name="Arial", bold=True, color=BLUE, size=10)
            c.fill = PatternFill("solid", fgColor=C_LIGHT)
        else:
            c.font = Font(name="Arial", color=BLACK, size=10)
        c.number_format = "0.00%" if unit in ("%", "ratio") and isinstance(val, float) and val < 1 else "#,##0.##"
        c.border = _bd(); c.alignment = Alignment(horizontal="right")
        _cell(ws, row, 3, unit, align="center")
        row += 1

    row += 1

    # ── Section: Formula Drivers ───────────────────────────────────────────────
    _sec(ws, row, 1, "YEAR-BY-YEAR FORMULA DRIVERS  (blue = editable, green = linked from Inputs)", span=4); row += 1

    if stype == "natural_shading":
        mat_val = sp.get("maturity_years", 8)
        _cell(ws, row, 1, "Maturity Period (years)", bold=True)
        c = ws.cell(row=row, column=2, value=mat_val)
        c.font = Font(name="Arial", bold=True, color=BLUE, size=10)
        c.number_format = "0"; c.border = _bd()
        c.fill = PatternFill("solid", fgColor=C_LIGHT)
        c.alignment = Alignment(horizontal="right")
        _cell(ws, row, 3, "years")
        _cell(ws, row, 4, "← edit to change maturity ramp", color="94A3B8")
        MAT_ROW = row; row += 1
    else:
        MAT_ROW = None

    _cell(ws, row, 1, "Discount Rate (from Inputs)", bold=True)
    dr_formula = f"=Inputs!$B${rm['dr']}"
    c = ws.cell(row=row, column=2, value=dr_formula)
    c.font = Font(name="Arial", bold=True, color=GREEN_LK, size=10)
    c.number_format = "0.00%"; c.border = _bd()
    c.alignment = Alignment(horizontal="right")
    _cell(ws, row, 3, "%")
    _cell(ws, row, 4, "← change in Inputs sheet", color="94A3B8")
    DR_CELL_ROW = row; row += 2

    # ── Section: Year-by-Year Benefit Tables (one per measure) ────────────────
    _sec(ws, row, 1, "YEAR-BY-YEAR BENEFIT PROJECTION  (formulas — updates automatically)", span=10); row += 1

    # Benefit type definitions per specialist type
    if stype == "natural_shading":
        BEN_TYPES = [
            ("avoided_mortality_npv",      "Avoided Mortality"),
            ("morbidity_savings_npv",      "Morbidity"),
            ("skin_cancer_prevention_npv", "Skin Cancer Prev."),
            ("carbon_sequestration_npv",   "Carbon Seq."),
            ("runoff_reduction_npv",       "Runoff Reduction"),
            ("air_quality_npv",            "Air Quality"),
            ("habitat_creation_npv",       "Habitat"),
        ]
    else:
        BEN_TYPES = [
            ("avoided_mortality_npv",      "Avoided Mortality"),
            ("morbidity_savings_npv",      "Morbidity"),
            ("property_value_uplift_npv",  "Prop. Value Uplift"),
            ("roof_longevity_npv",        "Roof Longevity"),
            ("carbon_sequestration_npv",   "Carbon Seq."),
            ("runoff_reduction_npv",       "Runoff Reduction"),
            ("air_quality_npv",            "Air Quality"),
            ("habitat_creation_npv",       "Habitat"),
        ]

    n_ben = len(BEN_TYPES)

    for mi, m in enumerate(measures):
        inp_col = get_column_letter(mi + 2)  # Inputs sheet column for this measure

        # Compute benefit fractions from advanced_benefits
        adv = m.get("advanced_benefits") or {}
        total_adv = sum((adv.get(k, 0) or 0) for k, _ in BEN_TYPES)
        if total_adv > 0:
            fracs = [(adv.get(k, 0) or 0) / total_adv for k, _ in BEN_TYPES]
        else:
            equal = 1.0 / n_ben
            fracs = [equal] * n_ben

        # Sub-table header
        _sec(ws, row, 1, f"Measure: {m['name']}", span=6 + n_ben); row += 1

        # Column headers (Year | Maturity | Base Benefit | Effective | Disc Factor | PV | benefit cols...)
        TABLE_HDRS = ["Year", "Maturity\nFactor", f"Base Benefit\n({cur}M/yr)",
                      f"Eff. Benefit\n({cur}M/yr)", "Disc.\nFactor", f"PV of Benefit\n({cur}M)"]
        TABLE_HDRS += [label for _, label in BEN_TYPES]
        for ci, h in enumerate(TABLE_HDRS, 1):
            _hdr(ws, row, ci, h, bg=C_ACCENT, sz=9, wrap=True)
        ws.row_dimensions[row].height = 30
        HDR_ROW = row; row += 1
        DATA_START = row

        mat_ref  = f"$B${MAT_ROW}" if MAT_ROW else None
        dr_ref   = f"$B${DR_CELL_ROW}"
        ben_ref  = f"Inputs!{inp_col}${rm['benefit']}"

        for yr in range(1, 51):
            r = row
            # Year
            c = ws.cell(row=r, column=1, value=yr)
            c.font = Font(name="Arial", size=9); c.border = _bd()
            c.alignment = Alignment(horizontal="center")

            # Maturity Factor
            if stype == "natural_shading" and mat_ref:
                mat_formula = f"=IF(A{r}<={mat_ref},A{r}/{mat_ref},1)"
            else:
                mat_formula = "=1"
            c = ws.cell(row=r, column=2, value=mat_formula)
            c.font = Font(name="Arial", color=BLACK, size=9); c.border = _bd()
            c.number_format = "0.00%"; c.alignment = Alignment(horizontal="right")

            # Base Annual Benefit
            c = ws.cell(row=r, column=3, value=f"={ben_ref}")
            c.font = Font(name="Arial", color=GREEN_LK, size=9); c.border = _bd()
            c.number_format = "#,##0.0"; c.alignment = Alignment(horizontal="right")

            # Effective Benefit
            c = ws.cell(row=r, column=4, value=f"=B{r}*C{r}")
            c.font = Font(name="Arial", color=BLACK, size=9); c.border = _bd()
            c.number_format = "#,##0.0"; c.alignment = Alignment(horizontal="right")

            # Discount Factor
            c = ws.cell(row=r, column=5, value=f"=1/(1+{dr_ref})^A{r}")
            c.font = Font(name="Arial", color=BLACK, size=9); c.border = _bd()
            c.number_format = "0.0000"; c.alignment = Alignment(horizontal="right")

            # PV of Benefit
            c = ws.cell(row=r, column=6, value=f"=D{r}*E{r}")
            c.font = Font(name="Arial", color=C_GREEN, bold=True, size=9); c.border = _bd()
            c.number_format = "#,##0.0"; c.alignment = Alignment(horizontal="right")
            c.fill = PatternFill("solid", fgColor="F0FDF4")

            # Benefit breakdown columns
            for bi, (frac, (key, _)) in enumerate(zip(fracs, BEN_TYPES)):
                c = ws.cell(row=r, column=7+bi, value=f"=D{r}*{frac:.6f}")
                c.font = Font(name="Arial", color=BLACK, size=9); c.border = _bd()
                c.number_format = "#,##0.0"; c.alignment = Alignment(horizontal="right")

            row += 1

        DATA_END = row - 1

        # Summary rows
        row += 1
        _cell(ws, row, 1, "Total NPV of Benefits", bold=True, bg=C_LIGHT)
        c = ws.cell(row=row, column=6,
                    value=f"=SUM(F{DATA_START}:F{DATA_END})")
        c.font = Font(name="Arial", bold=True, color=C_GREEN, size=10)
        c.number_format = "#,##0.0"; c.border = _bd()
        c.fill = PatternFill("solid", fgColor=C_LIGHT)
        c.alignment = Alignment(horizontal="right")
        # Track this total so Benefit Breakdown can reference it
        rm.setdefault("spec_benefit_totals", []).append(
            {"measure_name": m["name"], "row": row, "col": 6}
        )
        for bi in range(n_ben):
            c = ws.cell(row=row, column=7+bi,
                        value=f"=SUM({get_column_letter(7+bi)}{DATA_START}:{get_column_letter(7+bi)}{DATA_END})")
            c.font = Font(name="Arial", bold=True, color=BLACK, size=10)
            c.number_format = "#,##0.0"; c.border = _bd()
            c.fill = PatternFill("solid", fgColor=C_LIGHT)
            c.alignment = Alignment(horizontal="right")
        row += 1

        _cell(ws, row, 1, "Total Undiscounted Benefit", bold=False, bg=None)
        c = ws.cell(row=row, column=4,
                    value=f"=SUM(D{DATA_START}:D{DATA_END})")
        c.font = Font(name="Arial", color=BLACK, size=10)
        c.number_format = "#,##0.0"; c.border = _bd()
        c.alignment = Alignment(horizontal="right")
        row += 2

    # ── Explanatory benefit formulas (text only) ──────────────────────────────
    _sec(ws, row, 1, "BENEFIT FORMULAS (CONCEPTUAL OVERVIEW)", span=10); row += 1
    if stype == "natural_shading":
        expl_lines = [
            "Avoided mortality: vulnerable_population × base_mortality_rate × heat_mortality_factor × heat_reduction_efficiency × maturity_factor(year) × VSL, aggregated over 50 years and discounted.",
            "Morbidity savings: daily_hospital_cost (3,928 NIS) × average_length_of_stay (5.2 days) × heat_attributable_cases_avoided × efficiency × maturity_factor.",
            "Skin cancer prevention: pedestrians_per_hour × operating_hours (≈8) × UV_reduction_factor (0.75) × skin_cancer_incidence_rate × (treatment_cost + VSLY_loss) × maturity_factor.",
            "Ecosystem services: carbon sequestration (~450 NIS/tree/year × tree_density), runoff reduction (avoided drainage and treatment costs), air quality (PM2.5 reduction × health cost per unit), and habitat creation (200–500 NIS/m²/year biodiversity value).",
            "All benefit streams apply an 8-year linear biological maturity ramp (year/8 up to year 8, then 1.0) and are discounted over a 50-year horizon."
        ]
    else:
        expl_lines = [
            "Avoided mortality: catchment_population × base_mortality_rate × heat_mortality_factor × heat_reduction_efficiency × VSL, aggregated over 50 years and discounted.",
            "Morbidity savings: daily_hospital_cost (3,928 NIS) × average_length_of_stay (5.2 days) × heat_attributable_cases_avoided × heat_reduction_efficiency (≈0.28).",
            "Property value uplift: roof_area_m² × property_value_per_m² × uplift_pct (≈3%), treated as a one-time capital benefit near Year 1 and discounted.",
            "Roof longevity extension: (roof_replacement_cost / conventional_roof_lifetime) × roof_longevity_extension_years (≈15), treated as avoided replacement expenditure at end-of-life.",
            "Ecosystem services: carbon sequestration (~350 NIS/m²/year × green_roof_area), runoff reduction (stormwater_infrastructure_cost_avoided × runoff_reduction_coefficient ≈0.65), air quality (PM2.5 reduction × health cost per unit), and habitat creation (~300 NIS/m²/year × green_roof_area).",
            "Green roof benefits are assumed to operate at full capacity from Year 1 (no biological maturity ramp) and are discounted over a 50-year horizon."
        ]
    for line in expl_lines:
        c = ws.cell(row=row, column=1, value=line)
        c.font = Font(name="Arial", size=9, color="475569")
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
        row += 1

    # ── Column widths ──────────────────────────────────────────────────────────
    col_w = {1: 10, 2: 12, 3: 18, 4: 18, 5: 12, 6: 18}
    for bi in range(n_ben):
        col_w[7+bi] = 16
    col_w[4] = 22  # notes / params col
    _widths(ws, col_w)
    _widths(ws, {1: 10, 2: 14, 3: 20, 4: 20, 5: 13, 6: 20,
                 **{7+i: 16 for i in range(n_ben)}})


# ── SHEET 6: BENEFIT BREAKDOWN ────────────────────────────────────────────────
def _benefit_breakdown(wb, data, rm):
    """Show NPV of each benefit category per measure, plus totals and shares."""
    ws = wb.create_sheet("Benefit Breakdown")
    ws.sheet_view.showGridLines = False

    measures = data["measures"]
    n = rm["n"]
    cur = f"{data['currency']} ({data['currency_unit']})"

    _hdr(ws, 1, 1, "BENEFIT BREAKDOWN BY CATEGORY", sz=13, span=n+2)
    _hdr(
        ws,
        2,
        1,
        f"All benefit values in {cur}  |  Uses advanced_benefits NPVs from specialist analysis",
        bg=C_MID,
        fg="CBD5E1",
        bold=False,
        sz=9,
        span=n+2,
    )

    row = 4
    _sec(ws, row, 1, f"BENEFIT NPVs BY CATEGORY  [{cur}]", span=n+1); row += 1

    # Header row
    _hdr(ws, row, 1, "Benefit Category", bg=C_ACCENT, sz=10)
    for ci, m in enumerate(measures, 2):
        _hdr(ws, row, ci, m["name"], bg=C_ACCENT, sz=10)
    _hdr(ws, row, n+2, "Total Across Measures", bg=C_ACCENT, sz=10)
    row += 1

    BENEFIT_LABELS = [
        ("avoided_mortality_npv",      "Avoided Mortality — NPV"),
        ("morbidity_savings_npv",      "Morbidity Savings — NPV"),
        ("skin_cancer_prevention_npv", "Skin Cancer Prevention — NPV"),
        ("carbon_sequestration_npv",   "Carbon Sequestration — NPV"),
        ("runoff_reduction_npv",       "Runoff Reduction — NPV"),
        ("air_quality_npv",            "Air Quality — NPV"),
        ("habitat_creation_npv",       "Habitat Creation — NPV"),
        ("property_value_uplift_npv",  "Property Value Uplift — NPV"),
        ("roof_longevity_npv",         "Roof Longevity Extension — NPV"),
    ]

    first_benefit_row = row
    for key, label in BENEFIT_LABELS:
        # Only show rows that are used by at least one measure
        if not any((m.get("advanced_benefits") or {}).get(key) for m in measures):
            continue
        _cell(ws, row, 1, label, bold=False, bg="F0FDF4")
        row_total_cells = []
        for ci, m in enumerate(measures, 2):
            val = (m.get("advanced_benefits") or {}).get(key, 0) or 0
            c = ws.cell(row=row, column=ci, value=val)
            c.font = Font(name="Arial", color=BLACK, size=10)
            c.number_format = "#,##0.0"; c.border = _bd()
            c.alignment = Alignment(horizontal="right")
            c.fill = PatternFill("solid", fgColor="F0FDF4")
            row_total_cells.append(c.coordinate)
        if row_total_cells:
            total_formula = f"=SUM({row_total_cells[0]}:{row_total_cells[-1]})"
            c = ws.cell(row=row, column=n+2, value=total_formula)
            c.font = Font(name="Arial", bold=False, color=BLACK, size=10)
            c.number_format = "#,##0.0"; c.border = _bd()
            c.alignment = Alignment(horizontal="right")
            c.fill = PatternFill("solid", fgColor="F0FDF4")
        row += 1
    last_benefit_row = row - 1

    # Total PV of benefits per measure
    row += 1
    _cell(ws, row, 1, "Total PV of Benefits (all categories)", bold=True, bg=C_LIGHT)
    for ci in range(2, n+2):
        col_letter = get_column_letter(ci)
        c = ws.cell(
            row=row,
            column=ci,
            value=f"=SUM({col_letter}{first_benefit_row}:{col_letter}{last_benefit_row})",
        )
        c.font = Font(name="Arial", bold=True, color=C_GREEN, size=10)
        c.number_format = "#,##0.0"; c.border = _bd()
        c.fill = PatternFill("solid", fgColor=C_LIGHT)
        c.alignment = Alignment(horizontal="right")
    row += 2

    # Share of total benefits by category (per measure)
    _sec(ws, row, 1, "SHARE OF TOTAL BENEFITS BY CATEGORY  (per measure)", span=n+1); row += 1
    _hdr(ws, row, 1, "Benefit Category", bg=C_ACCENT, sz=10)
    for ci, m in enumerate(measures, 2):
        _hdr(ws, row, ci, m["name"], bg=C_ACCENT, sz=10)
    row += 1

    total_row = last_benefit_row + 2  # row where total PV of benefits per measure was written
    share_start_row = row
    for r in range(first_benefit_row, last_benefit_row + 1):
        label = ws.cell(row=r, column=1).value
        _cell(ws, row, 1, label, bold=False)
        for ci in range(2, n+2):
            num = f"{get_column_letter(ci)}{r}"
            den = f"{get_column_letter(ci)}{total_row}"
            c = ws.cell(row=row, column=ci, value=f"=IF({den}>0,{num}/{den},0)")
            c.font = Font(name="Arial", color=BLACK, size=10)
            c.number_format = "0.0%"; c.border = _bd()
            c.alignment = Alignment(horizontal="right")
        row += 1

    row += 2
    _sec(ws, row, 1, "LINKS TO YEAR-BY-YEAR BENEFITS (Specialist Detail)", span=n+1); row += 1
    _hdr(ws, row, 1, "Measure", bg=C_ACCENT, sz=10)
    _hdr(ws, row, 2, "Total NPV of Benefits (Specialist Detail, col F)", bg=C_ACCENT, sz=10)
    _hdr(ws, row, 3, "Total PV of Benefits (from table above)", bg=C_ACCENT, sz=10)
    row += 1

    spec_totals = rm.get("spec_benefit_totals", [])
    for mi, m in enumerate(measures):
        _cell(ws, row, 1, m["name"], bold=True)
        # If we have recorded a Specialist Detail total row for this measure, link to it
        link_cell_val = ""
        if mi < len(spec_totals):
            info = spec_totals[mi]
            spec_row = info["row"]
            spec_col = info["col"]
            spec_coord = f"$F${spec_row}" if spec_col == 6 else f"{get_column_letter(spec_col)}{spec_row}"
            link_cell_val = f"='Specialist Detail'!{spec_coord}"
        c = ws.cell(row=row, column=2, value=link_cell_val or None)
        if link_cell_val:
            c.font = Font(name="Arial", color=GREEN_LK, size=10)
            c.number_format = "#,##0.0"; c.border = _bd()
            c.alignment = Alignment(horizontal="right")

        total_col_letter = get_column_letter(mi + 2)
        c2 = ws.cell(row=row, column=3, value=f"={total_col_letter}{total_row}")
        c2.font = Font(name="Arial", color=BLACK, size=10)
        c2.number_format = "#,##0.0"; c2.border = _bd()
        c2.alignment = Alignment(horizontal="right")

        row += 1

    _widths(ws, {1: 34, **{i+2: 18 for i in range(n)}, n+2: 22})
