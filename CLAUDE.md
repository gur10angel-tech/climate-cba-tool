# Climate CBA Tool — Project Rules

## Excel Calculation Transparency Standard

**Every calculated value in this project's Excel output MUST be visible and auditable in the worksheet. No calculation may happen silently in Python and appear as a static number in a cell.**

---

### Rule 1: No Python-computed static values for methodology calculations

Never write:
```python
ws.cell(value=some_python_float)   # bad — hides the calculation
```

Always write an Excel formula instead:
```python
ws.cell(value=f"=B{row1}*B{row2}")  # good — user can audit it
```

**Exception:** Pure data fields (project name, year numbers, static labels, currency codes) may be static strings/numbers.

---

### Rule 2: Use the step-by-step block pattern for every calculation

All arithmetic must use the `_block_title()` / `_step()` / `_final()` helper pattern:

```python
_block_title("CALCULATION NAME", row); row += 1
r_a = row; _step("Input A (label)",        f"=ref_to_input",    "unit",  row, is_ref=True);      row += 1
r_b = row; _step("× Input B",              f"=ref_to_input",    "unit",  row, is_ref=True);      row += 1
r_c = row; _step("= A × B (intermediate)", f"=B{r_a}*B{r_b}",  "",      row, is_computed=True); row += 1
_final(key_or_label, "► Final Result",      f"=B{r_c}/1000000", "#,##0.000", "unit", row); row += 2
```

- `is_ref=True` → green font (referencing a parameter defined elsewhere)
- `is_computed=True` → grey fill (intermediate computed step)
- `_final()` → green highlighted result row (the ► output used by downstream sheets)

Every new benefit, cost component, financial metric, or sensitivity scenario MUST follow this pattern.

---

### Rule 3: Each calculation type has a designated sheet

| Calculation type | Sheet |
|-----------------|-------|
| VSL derivation chain (8 steps) | ASSUMPTIONS — Section 1 |
| Benefit sub-calculations (annual base value per functional unit) | ASSUMPTIONS — Section 5 |
| NPV, BCR, Annuity Factor, Payback Period | CALCULATIONS — Section 2 |
| Sensitivity analysis BCR recalculations | CALCULATIONS — Section 3 |
| Year-by-year discounting and maturity ramp | Specialist Detail |

---

### Rule 4: Cross-sheet links, never re-computation

If a value is already computed in ASSUMPTIONS or CALCULATIONS, reference it:
```python
f"=ASSUMPTIONS!$B${row}"    # good
f"={some_python_value}"     # bad — re-computes hidden from the user
```

---

### Rule 5: All `am` dict entries must point to Excel formula cells

When `_assumptions()` returns the `am` dict and `_calculations()` returns calc row numbers, those references must point to `►` final rows — cells that contain live Excel formulas, not static values.

---

## Helper Functions Reference

| Helper | Sheet it lives in | Purpose |
|--------|-------------------|---------|
| `_block_title(label, row)` | ASSUMPTIONS / CALCULATIONS | Dark header bar |
| `_sub_title(label, row)` | CALCULATIONS | Blue sub-section header |
| `_step(label, val, note, row, is_ref, is_computed)` | ASSUMPTIONS / CALCULATIONS | One labeled calculation row |
| `_final(label, val, fmt, note, row)` | ASSUMPTIONS / CALCULATIONS | Highlighted ► result row |
| `_inp_row(label, val, fmt, note, row)` | ASSUMPTIONS | Blue editable input row |
| `_frm_row(label, formula, fmt, note, row, highlight)` | ASSUMPTIONS | Black formula row |
