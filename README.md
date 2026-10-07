# Mueller Industries / Mueller Water Products — Horizontal Consolidation Model

> ⚠️ **HYPOTHETICAL / INDEPENDENT ANALYSIS — NOT AN ANNOUNCED TRANSACTION.**
> This is a self-directed MBA project exploring a hypothetical combination
> of two real, unrelated public companies. It is not affiliated with,
> endorsed by, or based on any non-public information from either company,
> and is not investment advice or a signal of any pending deal.

## What this is

A forecasting and credit model exploring a hypothetical acquisition of
**Mueller Water Products (NYSE: MWA)** by **Mueller Industries (NYSE: MLI)** —
two real, unrelated companies that happen to share a name but serve
overlapping water-infrastructure end markets (MLI: copper tube, fittings,
brass rod; MWA: iron gate valves, fire hydrants, smart water metering).

All standalone financials are sourced from each company's FY2025 10-K and
earnings releases (filed with the SEC). All synergy, growth, and margin
assumptions used to build the forward forecast are explicitly labeled as
illustrative and are **not** disclosed guidance from either company.

## What's in the model

| Tab | Contents |
|---|---|
| `Executive_Summary` | Headline pro forma figures, linked (not hardcoded) |
| `MLI_Standalone` | FY2025 actuals: revenue, COGS, SG&A, EBITDA, net income, debt, cash, CapEx — sourced to the 10-K, each figure cell-commented with its source |
| `MWA_Standalone` | Same, for MWA |
| `Copper_Cost_Engine` | Models how a lag between COMEX copper pricing and MLI's realized selling price affects fabrication margin |
| `Pro_Forma_Consolidated` | Combined FY2025 EBITDA bridge, built from each company's reported headline EBITDA plus an illustrative synergy assumption |
| `Five_Year_Forecast` | FY2026E–FY2030E driver-based revenue (volume × price/mix, not a flat blended rate), EBITDA tied to the Copper_Cost_Engine's base-case price assumption, CapEx anchored on MLI's disclosed FY2026 guidance, a full EBITDA→Free Cash Flow bridge (D&A, interest, tax, ΔNWC, dividends, cost-to-achieve), and a credit panel that rolls cash forward with actual FCF rather than holding it static |
| `Sensitivity_Matrix` | 2D grid: copper price shock × synergy realization, impact on pro forma EBITDA |
| `Audit_Checks` | Formula checks validating that the pro forma and forecast tie to their components |

## Key assumptions (see cell comments in the workbook for full detail)

- **Revenue is driver-based, not a flat rate:** each company's growth is built as volume growth × price/mix growth, not one undifferentiated percentage. MLI: ~2.0% volume + ~2.0% price/mix (price/mix held modest, consistent with the copper price being held flat in the base case — see below). MWA: ~2.5% volume + ~2.4% price. Both components are illustrative — neither company discloses unit volume at this granularity.
- **CapEx uses disclosed guidance where it exists:** MLI's FY2026E CapEx is set to $85M, the midpoint of the company's own disclosed $80–90M FY2026 guidance, not a historical ratio. Later years use the ratio implied by that guidance. MWA CapEx, where no guidance was found, uses its historical capex/revenue ratio.
- **The copper engine and the forecast are consistent with each other:** the base case explicitly holds the COMEX copper price flat at the Current-Quarter average ($5.13/lb), and MLI's EBITDA margin assumption is stated as consistent with that. The `Sensitivity_Matrix` tab shocks the forecast's FY2026E figures, so a copper price move visibly flows into the forecast.
- **Synergies ramp in, and cost money to get:** $14M illustrative run-rate synergies phase in over four years (25%/50%/75%/100%), and a cost-to-achieve of 1.2x run-rate synergies ($16.8M) is modeled as a one-time cash outlay in Years 1–2 — not an EBITDA reduction, but a real cash cost, consistent with how integrations actually work.
- **A full EBITDA-to-Free-Cash-Flow bridge:** EBITDA → less D&A → EBIT → less interest (on MWA's existing 4.0% Senior Notes only — MLI carries zero debt) → pretax income → less tax (24% illustrative blended rate) → net income → add back D&A → less CapEx → less incremental net working capital → less dividends (each company's own trailing payout ratio, applied to its EBITDA-weighted share of combined net income) → less cost-to-achieve → Free Cash Flow.
- **The credit panel rolls cash forward, not flat:** combined debt stays at $451.6M because MWA's Senior Notes are a bullet structure with no maturity until 2029 and no scheduled amortization — but cash compounds using actual Free Cash Flow each year, so the leverage ratio reflects real cash generation, dividends, and reinvestment rather than an assumption held artificially static.

## Known data-quality caveat

MLI's disclosed COGS + D&A + SG&A line items sum to $3,283.3M, but the 10-K's
own "Operating expenses" subtotal is $3,220.0M — a $63.4M gap, most likely
unallocated corporate costs not broken out in the excerpt available. The
model uses MLI's **headline reported** Operating Income and EBITDA (not a
bottom-up reconstruction from the three line items) as the authoritative
figures, and flags the gap on the `MLI_Standalone` and `Audit_Checks` tabs
rather than papering over it.

## Model Governance

The model was reviewed for formula integrity, source consistency, forecast
linkage, cash-flow reconciliation, and sensitivity functionality. Key checks
are documented in the `Audit_Checks` tab.

## Sources

- Mueller Industries FY2025 10-K (filed Feb 2026), SEC EDGAR
- Mueller Industries Q4/FY2025 earnings release
- Mueller Water Products FY2025 10-K and Q4/FY2025 earnings release, SEC EDGAR

## Files

```
/
├── README.md                     — this file
├── MLI_MWA_Project_Report.pdf    — narrative write-up: approach, assumptions, findings
└── model/
    ├── MLI_MWA_Corrected_Model.xlsx   — the model
    └── build_model.py                 — generator script
```
