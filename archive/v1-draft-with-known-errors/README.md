# v1 draft — kept deliberately, not a usable deliverable

These four files are the **first draft** of this project. They are kept here
on purpose, as documentation of the correction process — not because they're
usable.

**Known errors in this draft, since fixed in `/model/`:**

1. MLI's standalone financials were fabricated. The draft shows MLI at a
   **negative** EBITDA margin (-1.4%); MLI's actual FY2025 results (per its
   10-K) were $4,178.5M revenue and $1,027.1M EBITDA (24.6% margin) — the
   best year in the company's history.
2. The pro forma EBITDA bridge in `Mueller_Industries_MWA_Forecasting_Report.md`
   doesn't reconcile with its own component figures (sums to $318M per the
   table's own line items; the table reports $203M).
3. The copper-pricing formula in `generate_mli_mwa_model.py` is
   self-referential: `=B4-(B4-4.50)*(B5/365)*0.70`, where B4 is the same
   cell that holds the value 4.50 — so the pricing-lag mechanic silently
   does nothing unless the input is moved away from its own default.

See the main `README.md` and `MLI_MWA_Project_Report.pdf` for how each of
these was corrected.
