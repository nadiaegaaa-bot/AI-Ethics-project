"""Convert the Excel data package into js/data.js for the dashboard.

Usage (from the "Final Project" folder):
    python3 tools/excel_to_data.py data/Bulk_Energy_Intelligence_Data_and_Competitor_Benchmarks.xlsx

Uses only the Python standard library. Reads the cached cell values, so formulas
must have been calculated (saved by Excel) before running.
"""
import datetime
import json
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
HEADER_ROW = 5


def col_index(ref):
    letters = "".join(ch for ch in ref if ch.isalpha())
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_workbook(path):
    z = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(M + "si"):
            shared.append("".join(t.text or "" for t in si.iter(M + "t")))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    sheets = {}
    for s in ET.fromstring(z.read("xl/workbook.xml")).find(M + "sheets"):
        target = rels[s.get(R)].lstrip("/")
        target = target if target.startswith("xl/") else "xl/" + target
        rows = {}
        for row in ET.fromstring(z.read(target)).iter(M + "row"):
            cells = {}
            for c in row.findall(M + "c"):
                v, inline, t = c.find(M + "v"), c.find(M + "is"), c.get("t")
                if t == "s" and v is not None:
                    val = shared[int(v.text)]
                elif inline is not None:
                    val = "".join(x.text or "" for x in inline.iter(M + "t"))
                elif v is not None:
                    val = v.text
                    if t not in ("str", "b"):
                        try:
                            val = float(val)
                        except ValueError:
                            pass
                else:
                    val = None
                if val not in (None, ""):
                    cells[col_index(c.get("r"))] = val
            rows[int(row.get("r"))] = cells
        sheets[s.get("name")] = rows
    return sheets


def table(rows, keys):
    """Rows below the header row, mapped to the given keys by column order."""
    out = []
    for r in sorted(k for k in rows if k > HEADER_ROW):
        cells = rows[r]
        if not cells:
            continue
        out.append({k: cells.get(i) for i, k in enumerate(keys) if k})
    return out


def excel_month(serial):
    d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(serial))
    return f"{d.year}-{d.month:02d}"


def rnd(x, n=3):
    return round(x, n) if isinstance(x, float) else x


def main(xlsx, out_js):
    s = read_workbook(xlsx)

    monthly = table(s["Bulk Monthly Synthetic"], ["month", "site", "zone", "days", "hours", "itLoadMW", "itMWh",
                    "coolingMWh", "otherMWh", "totalMWh", "pue", "tempC", "priceNOKMWh", "costNOK",
                    "expectedCoolingMWh", "residualMWh", "residualPct", "alert", "basis", "provenance", "year"])
    for r in monthly:
        r["month"] = excel_month(r["month"])
        for k, v in list(r.items()):
            r[k] = rnd(v, 5 if k in ("pue", "residualPct") else 3)
        r["year"] = int(r["year"])
        r.pop("basis", None)

    peer_monthly = table(s["Peer Synthetic Scenarios"], ["month", "peer", "refYear", "refPUE", "pue", "totalMWh", "itMWh", "provenance"])
    for r in peer_monthly:
        r["month"] = excel_month(r["month"])
        r["refYear"] = int(r["refYear"])
        r["itMWh"] = rnd(r["itMWh"], 1)

    benchmarks = table(s["Competitor Benchmarks"], ["operator", "facility", "year", "pue", "pueType", "energyMWh",
                       "energyScope", "status", "relevance", "source", "cautions"])
    for r in benchmarks:
        if isinstance(r.get("year"), float):
            r["year"] = int(r["year"])

    facts = table(s["Real Company Facts"], ["entity", "period", "metric", "value", "unit", "source", "classification", "caveat"])
    sources = table(s["Sources & Methods"], ["id", "organization", "document", "period", "type", "use"])
    dictionary = table(s["Data Dictionary"], ["sheet", "fields", "unit", "method", "notes"])
    guide = [{"step": r.get(0), "text": r.get(1)} for k, r in sorted(s["START HERE"].items()) if k >= 5 and r.get(1)]

    # Dashboard sheet: annual 2025 summary used as a cross-check.
    dash = s["Dashboard"]
    sites = [dash[5][i] for i in (1, 2, 3)]
    summary = {}
    for rnum in range(6, 14):
        label = dash[rnum][0]
        summary[label] = {sites[i]: rnd(dash[rnum][i + 1], 4) for i in range(3)}
        summary[label]["scope"] = dash[rnum].get(4)

    data = {
        "source": Path(xlsx).name,
        "generated": datetime.date.today().isoformat(),
        "monthly": monthly,
        "peerMonthly": peer_monthly,
        "benchmarks": benchmarks,
        "facts": facts,
        "sources": sources,
        "dictionary": dictionary,
        "guide": guide,
        "dashboardSummary2025": summary,
        "reviewThreshold": 0.12,
    }
    js = ("// Generated by tools/excel_to_data.py from " + Path(xlsx).name + ". Do not edit by hand.\n"
          "window.BULK_DATA = " + json.dumps(data, indent=1, ensure_ascii=False) + ";\n")
    Path(out_js).write_text(js, encoding="utf-8")
    print(f"monthly={len(monthly)} peerMonthly={len(peer_monthly)} benchmarks={len(benchmarks)} "
          f"facts={len(facts)} sources={len(sources)} dictionary={len(dictionary)} -> {out_js}")


if __name__ == "__main__":
    here = Path(__file__).resolve().parent.parent
    xlsx = sys.argv[1] if len(sys.argv) > 1 else here / "data" / "Bulk_Energy_Intelligence_Data_and_Competitor_Benchmarks.xlsx"
    main(str(xlsx), str(here / "js" / "data.js"))
