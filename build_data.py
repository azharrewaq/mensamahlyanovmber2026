"""Builds data.json from the Excel file. Usage: python3 build_data.py applicants.xlsx
National IDs are stored only as SHA-256 hashes (not in plain text) so the public repo doesn't expose them."""
import sys, json, hashlib, openpyxl
SALT = "azhar-rawaq-2026-11:"
ws = openpyxl.load_workbook(sys.argv[1], data_only=True).worksheets[0]
dates, slots, rows = [], [], []
def idx(lst, v):
    if v not in lst: lst.append(v)
    return lst.index(v)
for r in ws.iter_rows(min_row=2, values_only=True):
    if not r or not r[0]: continue
    name = " ".join(str(r[0]).split())
    nid = str(r[1] or "").strip()
    d, t = (r[2] or ""), (r[3] or "")
    rows.append([name,
                 hashlib.sha256((SALT + nid).encode()).hexdigest()[:20] if nid else "",
                 idx(dates, str(d)), idx(slots, str(t))])
json.dump({"salt": SALT, "dates": dates, "slots": slots, "rows": rows},
          open("data.json", "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
print(len(rows), "rows")
