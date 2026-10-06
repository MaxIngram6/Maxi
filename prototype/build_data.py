"""Build the prototype's seed data from the existing Eurogold spreadsheets.

Usage:
  python3 build_data.py LEDGER.xlsx RATES.xlsx PHOTO1.jpg [PHOTO2.jpg ...] > data.js
"""
import base64, io, json, re, sys
import openpyxl
from PIL import Image

ledger_path, rates_path, *photo_paths = sys.argv[1:]

# Booking items, in the order they appear on the processing sheet.
ITEMS = [
    ("t355", "Trench block 355mm", "No"),
    ("t300", "Trench block 300mm", "No"),
    ("b140", "140mm block", "No"),
    ("b100", "100mm block", "No"),
    ("commons", "Commons", "No"),
    ("splits", "40mm splits", "No"),
    ("stone", "Stone block", "Tonne"),
    ("boe", "Brick on edge (BOE)", "m"),
    ("facings", "Facings / engineering", "No"),
    ("bnb", "Block and beam", "m2"),
    ("lintels", "Lintels", "No"),
    ("pipes", "Service pipes", "No"),
    ("vents", "Vents", "No"),
]
# Ledger booked columns Q..AC -> item keys
LEDGER_BOOKED = {17: "t355", 18: "t300", 19: "b100", 20: "b140", 21: "commons", 22: "facings",
                 23: "splits", 24: "stone", 25: "boe", 26: "lintels", 27: "pipes", 28: "vents", 29: "bnb"}
# Rates sheet columns C..O -> item keys
RATE_COLS = {3: "t355", 4: "t300", 5: "b140", 6: "b100", 7: "commons", 8: "splits", 9: "stone",
             10: "boe", 11: "facings", 12: "lintels", 13: "pipes", 14: "vents", 15: "bnb", 16: "daywork"}


def num(v):
    return v if isinstance(v, (int, float)) else 0


def parse_booked(v):
    """Ledger cells are sometimes typed as text ('236+1210', '6-588', '1 vent - 11').
    Excel's SUM ignores those, so we recover the intended number and flag it."""
    if v is None or v == "":
        return 0, None
    if isinstance(v, (int, float)):
        return float(v), None
    s = str(v).strip()
    nums = [float(n) for n in re.findall(r"\d+(?:\.\d+)?", s)]
    if not nums:
        return 0, s
    if "+" in s:
        return sum(nums), s
    return nums[-1], s  # "courses-quantity" style: quantity is the last number


wb = openpyxl.load_workbook(ledger_path, data_only=True)
ws = wb["Ledger"]
plots, fixes = [], []
for r in range(17, ws.max_row + 1):
    no = ws.cell(r, 1).value
    if not isinstance(no, (int, float)):
        continue
    c = lambda col: num(ws.cell(r, col).value)
    allowance = {
        "trench": c(5),               # E: 2 courses trench block
        "b100": c(6) + c(14),         # F + N
        "b140": c(7) + c(15),         # G + O
        "commons": c(8) + c(13),      # H + M
        "facings": c(9),              # I: 6 courses engineering
    }
    lv = c(11) + c(12)                # K/L vents, pipes & lintels (empty on this tender)
    opening = {}
    for col, key in LEDGER_BOOKED.items():
        q, raw = parse_booked(ws.cell(r, col).value)
        if q:
            opening[key] = round(q, 2)
        if raw is not None:
            fixes.append({"plot": int(no), "item": key, "text": raw, "value": q})
    history = []
    for bc, dc in ((31, 32), (34, 35), (37, 38)):
        who, when = ws.cell(r, bc).value, ws.cell(r, dc).value
        if who or when:
            history.append({"by": (who or "").strip(),
                            "date": when.strftime("%Y-%m-%d") if hasattr(when, "strftime") else (str(when) if when else "")})
    plots.append({
        "no": int(no), "type": (ws.cell(r, 2).value or "").strip(), "foundation": ws.cell(r, 3).value or "",
        "allowance": allowance, "extraCourse": c(10), "lintelsPipesVents": lv,
        "opening": opening, "openingClaims": history,
    })

rwb = openpyxl.load_workbook(rates_path, data_only=True)["Sheet1"]
rates = {}
for r in range(2, rwb.max_row + 1):
    name = rwb.cell(r, 2).value
    if not name or name in ("Name", "AVERAGE", "MAX", "RISK"):
        continue
    name = name.strip()
    rates[name] = {k: rwb.cell(r, col).value or 0 for col, k in RATE_COLS.items()}

photos = []
for p in photo_paths:
    im = Image.open(p).convert("RGB")
    im.thumbnail((900, 900))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=62)
    photos.append("data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode())

data = {
    "site": {"code": "EG587", "name": "Old Hall Road, Bromborough", "client": "Northstone Developments",
             "ledgerRef": "EG582 Brickwork Ledger"},
    "items": [{"key": k, "label": l, "unit": u} for k, l, u in ITEMS],
    "plots": plots, "rates": rates, "ledgerTextFixes": fixes, "demoPhotos": photos,
}
print("window.SEED = " + json.dumps(data, separators=(",", ":")) + ";")
