#!/usr/bin/env python3
"""Generate print-ready HTML catalogs (English) from the CSV data files."""
import csv
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUT = ROOT / "catalogs"
OUT.mkdir(exist_ok=True)

CSS = """
body{font-family:Arial,Helvetica,sans-serif;color:#1a1a1a;margin:40px auto;max-width:960px;background:#fff}
h1{font-size:26px;border-bottom:3px solid #1F4E78;padding-bottom:8px;color:#1F4E78}
h2{font-size:18px;color:#1F4E78;margin-top:28px}
.sub{color:#555;font-size:13px;margin-bottom:20px}
table{border-collapse:collapse;width:100%;font-size:12px;margin:10px 0 24px}
th{background:#1F4E78;color:#fff;text-align:left;padding:6px 8px}
td{border-bottom:1px solid #ddd;padding:5px 8px}
tr:nth-child(even) td{background:#f5f8fb}
.num{text-align:right;white-space:nowrap}
.note{background:#fff8e1;border:1px solid #e0c060;padding:10px 14px;font-size:12px;margin:18px 0}
.footer{margin-top:30px;font-size:11px;color:#777;border-top:1px solid #ccc;padding-top:10px}
@media print{body{margin:10mm}}
"""

DISCLAIMER = ("Offered by Milhurst Mills (seller of record); presented by its authorized sales agent. "
              "Preliminary draft for internal review. Quantities subject to physical verification. "
              "Prices subject to change and final confirmation; nothing herein constitutes a binding offer. "
              "All payments to Milhurst Mills only. All transactions subject to seller's counsel approval.")


def read(name):
    with open(DATA / name, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def money(v, dec=0):
    if v in (None, ""):
        return ""
    return f"${float(v):,.{dec}f}"


def page(title, subtitle, body):
    return (f"<!doctype html><html><head><meta charset='utf-8'><title>{title}</title>"
            f"<style>{CSS}</style></head><body><h1>{title}</h1>"
            f"<div class='sub'>{subtitle}</div>{body}"
            f"<div class='footer'>{DISCLAIMER}<br>Prepared August 2026.</div></body></html>")


# ------------------------------------------------------------------ slabs
rows = read("slabs_inventory.csv")
groups = {}
for r in rows:
    groups.setdefault(r["Category"], []).append(r)
body = ["<div class='note'><b>Lot summary:</b> ~2,541 slabs (porcelain, quartz, marble) + 100,000 ceramic tiles + 80 steel A-frames. "
        "Origins: Spain, Italy, Germany (Quartzforms), Vietnam, India. Bulk, pallet-lot and container-lot pricing available. "
        "Photos and batch labels available on request.</div>"]
grand_pcs = 0
for g, items in groups.items():
    body.append(f"<h2>{g}</h2><table><tr><th>Product</th><th>Origin</th><th>Size (mm)</th>"
                "<th class='num'>Pcs</th><th class='num'>Bulk Price / pc</th><th class='num'>List Price / pc</th></tr>")
    for r in items:
        la, lb = r["ListA_Jan2025"], r["ListB_Jun2025"]
        pcs = int(float(r["Pcs"]))
        grand_pcs += pcs
        dec = 2 if la and float(la) < 10 else 0
        body.append(f"<tr><td>{r['Product']}</td><td>{r['Origin']}</td><td>{r['Size']}</td>"
                    f"<td class='num'>{pcs:,}</td><td class='num'>{money(la, dec)}</td>"
                    f"<td class='num'>{money(lb or la, dec)}</td></tr>")
    body.append("</table>")
body.append(f"<p><b>Total pieces: {grand_pcs:,}</b> &nbsp;|&nbsp; Full-lot and category-lot offers welcome — ask for bulk pricing.</p>")
(OUT / "slab_catalog.html").write_text(
    page("Porcelain, Quartz & Marble Slab Inventory — Bulk Sale Catalog",
         "Offered by Milhurst Mills. Premium European porcelain and quartz slabs — Spain, Italy, Germany (Quartzforms), Vietnam, India. New, first-quality, warehouse-stored.",
         "".join(body)))

# ------------------------------------------------------------------ equipment
rows = read("equipment_inventory.csv")
body = ["<div class='note'><b>Condition:</b> Late-model units acquired 2024, low/no hours, warehouse-stored. "
        "'Available' items confirmed in the May 2026 inventory; others pending verification. "
        "Also available: Ausavina stone-handling equipment (clamps, vacuum lifters, A-frames, trolleys, ladders & scaffolding) — separate Ausavina US price list available on request.</div>",
        "<table><tr><th>Category</th><th>Item</th><th>Brand</th><th class='num'>Qty</th>"
        "<th class='num'>Asking (each)</th><th>Status</th></tr>"]
for r in rows:
    ask = float(r["UnitCost"]) * 0.75
    status = "Available" if r["OnHand_May2026"] == "Yes" else "Verify"
    body.append(f"<tr><td>{r['Category']}</td><td>{r['Item']}</td><td>{r['Brand']}</td>"
                f"<td class='num'>{r['Qty_2024']}</td><td class='num'>{money(ask)}</td><td>{status}</td></tr>")
body.append("</table><p>Volume discounts for multi-unit and full-lot purchases. Inspection welcome by appointment (Northern NJ).</p>")
(OUT / "equipment_catalog.html").write_text(
    page("Forklift & Warehouse Equipment — Sale Catalog",
         "Offered by Milhurst Mills. Forklifts (diesel / LPG / electric), electric pallet trucks, stackers, reach trucks, VNA trucks and towing tractors — Lonking, Zowell, Huaya, RY.",
         "".join(body)))

# ------------------------------------------------------------------ appliances
rows = read("appliances_inventory.csv")
body = ["<div class='note'><b>Condition:</b> New / unused surplus; several units factory-boxed. Model tags verified before invoicing. "
        "Package and volume pricing available. Pickup or delivery (Northern NJ).</div>"]
for grp, label in (("Luxury", "Luxury European Appliances — Gaggenau · Miele · Thermador · Bosch"),
                   ("Mainstream", "Mainstream Kitchen & Laundry — Frigidaire · Electrolux"),
                   ("HVAC", "HVAC — Greenheck · Aspen")):
    items = [r for r in rows if r["Group"] == grp]
    if not items:
        continue
    body.append(f"<h2>{label}</h2><table><tr><th>Brand</th><th>Model</th><th>Description</th>"
                "<th class='num'>MSRP</th><th class='num'>Asking</th><th>Notes</th></tr>")
    for r in items:
        msrp = r["MSRP"]
        ask = float(msrp) * 0.55 if msrp else None
        body.append(f"<tr><td>{r['Brand']}</td><td>{r['Model']}</td><td>{r['Description']}</td>"
                    f"<td class='num'>{money(msrp)}</td><td class='num'>{money(ask)}</td><td>{r['Notes']}</td></tr>")
    body.append("</table>")
(OUT / "appliances_catalog.html").write_text(
    page("Premium Appliance Inventory — Sale Catalog",
         "Offered by Milhurst Mills. New/unused surplus luxury and mainstream appliances at 45–55% off MSRP. Ideal for dealers, builders, kitchen contractors and multifamily projects.",
         "".join(body)))

print("catalogs written:", [p.name for p in OUT.glob("*.html")])
