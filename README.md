# Inventory Sale — Preliminary Work Package (Aug 2026)

Preliminary work for the bulk sale of slab, equipment, appliance and furniture inventory (US/Canada) and marketing of the Ultima mining license portfolio. Prepared from source files in Google Drive on 2026-08-27.

## Contents

| Path | What it is |
|---|---|
| `output/MASTER_INVENTORY.xlsx` | Master workbook — all categories, quantities, costs, and Quick/Mid/Slow price tiers (formula-driven; totals compute when opened in Excel/Google Sheets) |
| `data/*.csv` | Clean source data extracted from Drive files (slabs, equipment, appliances, furniture, mining sites, buyer prospects) |
| `catalogs/*.html` | Print-ready sale catalogs (slabs, equipment, appliances) — open in a browser, print to PDF; add photos before sending |
| `docs/SALES_PLAN_30_DAYS.md` | 30-day plan, price strategy, valuation summary, **legal guardrails (read first)**, data gaps |
| `docs/BUYER_PROSPECTS.md` | US & Canada buyer list with contact channels |
| `docs/EMAIL_TEMPLATES.md` | Outreach email drafts (EN) — **do not send before counsel sign-off** |
| `scripts/` | Reproducible builders: `build_workbook.py`, `build_catalogs.py` |

## Headline numbers (priced categories)

- Quick exit (7–21 days): **≈ $897k**
- Recommended blend (30–45 days): **≈ $1.25M ceiling**
- Slow/retail (60–90d+): ≈ $1.68M ceiling

## Structure

The inventory was sold in 2025 to **Milhurst Mills** (exact legal entity name to confirm); the principal markets and sells it as Milhurst's **authorized sales agent**. All catalogs, quotes and invoices issue in Milhurst's name; all proceeds go to Milhurst's account. Mining licenses (Ultima Ltd. Şti., Türkiye) are outside the Milhurst mandate.

## Before anything is offered or sold

1. Signed marketing/agency agreement with Milhurst (asset schedule + commission terms).
2. Bankruptcy counsel clearance (Chapter 11 — In re Cakmak, 26-11521-VFP); agent compensation disclosed in MORs.
3. Physical verification of quantities (equipment/furniture/Ausavina especially).
4. Mining license fee status verification.

Note: LibreOffice in this build environment cannot recalculate, so the workbook ships with formulas whose cached values are empty — Excel/Google Sheets computes them on open. Expected totals were independently verified in Python (see git history).
