#!/usr/bin/env python3
"""
critcareiq_update.py — regenerate CritCareIQ's data blocks from the source spreadsheets.

Reads
  • Vituity "E&M Billing" monthly reports (physicians only)  → cases, CC charts (99291), Level 5 charts (99285), raw CC %
  • the weekly "pts per hour physicians unblinded" workbook   → triage ESI mix and hours per physician-week

Writes (into index.html, between the /* @@CC_DATA START */ and /* @@CC_DATA END */ markers)
  • CC_MONTH_KEYS, REPL_RAW (kept from the existing file), CC_COUNTS
  • ESI_COMPLETE, ESI_PARTIAL, ESI_COUNTS
and, with --add-months, appends any billing month that isn't yet in the raw-rate tables
  (MONTHS, MONTH_KEYS, GROUP_RATES, each PHYSICIANS row).

Usage
  python3 critcareiq_update.py --billing "Critical Care Data" --pph "pts per hour physicians unblinded 2026.xlsx" --index index.html --write --add-months
  (omit --write to preview; add --out other.html to write elsewhere)

Requires: python3 and openpyxl  (pip3 install openpyxl)

Things the script does NOT touch — update by hand after a run:
  PARTIAL_KEY, LATEST_FINAL, STATUS_KEYS, CODED, DATA_ASOF, the header date badge, the footer, CC_NIGHTS.
"""
import argparse, datetime as dt, glob, json, os, re, sys
from collections import defaultdict

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is required:  pip3 install openpyxl")

MONTH_NAMES = ["january","february","march","april","may","june","july","august","september","october","november","december"]
ABBR = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

def mkey(year, month):            # 2026, 9 -> "sep26"
    return f"{ABBR[month-1].lower()}{year%100:02d}"
def mlabel(year, month):          # -> "Sep '26"
    return f"{ABBR[month-1]} '{year%100:02d}"

# ───────────────────────── billing ─────────────────────────
def read_billing(paths):
    """Return {mkey: {name: (cases, cc, l5, raw_pct)}} — newest file wins when a month appears twice."""
    out, seen = {}, {}
    for f in sorted(paths, key=os.path.getmtime):
        ws = openpyxl.load_workbook(f, data_only=True, read_only=True).active
        rows = list(ws.iter_rows(values_only=True))
        if not rows: continue
        hdr = [str(c).strip() if c is not None else "" for c in rows[0]]
        try:
            ci = next(i for i,h in enumerate(hdr) if h.startswith("Provider ("))
            mi = hdr.index("txt_Prior_month"); ni = hdr.index("Cases"); li = hdr.index("Level 5"); cci = hdr.index("Critical %")
        except (StopIteration, ValueError):
            print(f"  skip (not a billing report): {os.path.basename(f)}"); continue
        ti = hdr.index("Provider Type") if "Provider Type" in hdr else None
        month_key = None; data = {}
        for r in rows[1:]:
            name = r[ci]
            if not name or "[" not in str(name): continue
            if ti is not None and r[ti] and str(r[ti]).strip() != "MD/DO": continue
            cases = r[ni]
            if not isinstance(cases, (int, float)) or cases <= 0: continue
            if month_key is None:
                m = re.match(r"\s*([A-Za-z]+)\s+(\d{4})", str(r[mi]))
                month_key = mkey(int(m.group(2)), MONTH_NAMES.index(m.group(1).lower())+1)
            nm = str(name).split(" [")[0].strip()
            l5 = r[li] or 0; cc = r[cci] or 0
            data[nm] = (int(cases), round(cc*cases), round(l5*cases), round(cc*100, 1))
        if month_key:
            out[month_key] = data; seen[month_key] = os.path.basename(f)
            print(f"  billing {month_key}: {len(data)} physicians  ← {os.path.basename(f)}")
    return out

# ───────────────────────── weekly ESI ─────────────────────────
def parse_week_start(title, prev_start):
    """Sheet titles look like '8.30-9.5.26', '9.6-9.12.26', '.2-8.8.26', '7.26-8.' — fall back to prev + 7 days."""
    m = re.match(r"^\s*(\d{1,2})\.(\d{1,2})", title)
    if m and prev_start is not None:
        yr = prev_start.year
        cand = dt.date(yr, int(m.group(1)), int(m.group(2)))
        if cand < prev_start: cand = dt.date(yr+1, int(m.group(1)), int(m.group(2)))
        return cand
    if m:
        yr = 2000 + int(re.search(r"\.(\d{2})\s*$", title).group(1)) if re.search(r"\.(\d{2})\s*$", title) else dt.date.today().year
        return dt.date(yr, int(m.group(1)), int(m.group(2)))
    return prev_start + dt.timedelta(days=7) if prev_start else None

def read_pph(path):
    """Return ({mkey: {name: [pts,e1,e2,e3,e4,e5,hrs]}}, complete_keys, partial_notes)."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    agg = defaultdict(lambda: defaultdict(lambda: [0,0,0,0,0,0,0.0]))
    weeks = []   # (start_date, has_data)
    prev = None
    for ws in wb.worksheets:
        if "cumulative" in ws.title.lower(): continue
        start = parse_week_start(ws.title, prev)
        if start is None: continue
        prev = start
        rows = list(ws.iter_rows(values_only=True))
        hdr = None
        for r in rows:
            if r and any(c == "Provider Name" for c in r): hdr = list(r); break
        has = False
        if hdr:
            ci = hdr.index("Provider Name"); hi = hdr.index("Hours Worked")
            pi = next((i for i,h in enumerate(hdr) if h and str(h).startswith("#")), None)
            ei = [hdr.index(f"ESI {k}") for k in range(1,6)] if "ESI 1" in hdr else None
            mid = start + dt.timedelta(days=3)                      # the month holding most of the week
            key = mkey(mid.year, mid.month)
            for r in rows:
                nm = r[ci] if ci < len(r) else None
                if nm in (None, "Provider Name", "AVG") or not isinstance(r[hi], (int, float)): continue
                a = agg[key][str(nm).strip()]
                hrs = r[hi] or 0; pts = (r[pi] or 0) if pi is not None and isinstance(r[pi], (int,float)) else 0
                a[0] += pts; a[6] += hrs
                if ei:
                    for k in range(5):
                        v = r[ei[k]]; a[1+k] += v if isinstance(v, (int, float)) else 0
                if hrs > 0: has = True
        weeks.append((start, has))
    # completeness: a month is complete if every week whose majority falls in it has data
    # and the sheet series reaches past the month's end.
    last_covered = max((s + dt.timedelta(days=6) for s, h in weeks if h), default=None)
    months = sorted({mkey((s+dt.timedelta(days=3)).year, (s+dt.timedelta(days=3)).month) for s,_ in weeks})
    complete, partial = [], {}
    for key in months:
        wk = [(s,h) for s,h in weeks if mkey((s+dt.timedelta(days=3)).year, (s+dt.timedelta(days=3)).month) == key]
        y = int("20"+key[-2:]); mo = ABBR.index(key[:3].capitalize())+1
        month_end = (dt.date(y, mo+1, 1) if mo < 12 else dt.date(y+1,1,1)) - dt.timedelta(days=1)
        filled = [s for s,h in wk if h]
        if filled and all(h for _,h in wk) and last_covered and last_covered >= month_end:
            complete.append(key)
        elif filled:
            partial[key] = "through " + (max(filled)+dt.timedelta(days=6)).strftime("%-m/%-d")
    # convert names "First Last" -> "Last, First" to match billing
    out = {}
    for key, d in agg.items():
        out[key] = {}
        for nm, a in d.items():
            parts = nm.split()
            last = " ".join(parts[1:]); first = parts[0]
            out[key][(last + ", " + first)] = [int(a[0]), int(a[1]), int(a[2]), int(a[3]), int(a[4]), int(a[5]), round(a[6], 1)]
    return out, complete, partial

def canon(name, known):
    """Match 'DiSandro, Daniel' / 'Krass, Sam' to the billing spelling."""
    n = name.lower().replace(" ", "")
    for k in known:
        kk = k.lower().replace(" ", "")
        if kk == n: return k
    last = n.split(",")[0]
    cands = [k for k in known if k.lower().replace(" ", "").split(",")[0] == last]
    if len(cands) == 1: return cands[0]
    first = n.split(",")[1] if "," in n else ""
    for k in cands:
        kf = k.lower().replace(" ", "").split(",")[1]
        if kf.startswith(first[:3]) or first.startswith(kf[:3]): return k
    return None

# ───────────────────────── emit ─────────────────────────
def js_num(x): return "null" if x is None else (str(int(x)) if float(x).is_integer() else str(x))

def build_block(billing, esi, complete, partial, repl_raw):
    keys = sorted(billing, key=lambda k: (int(k[-2:]), ABBR.index(k[:3].capitalize())))
    names = sorted({n for d in billing.values() for n in d})
    L = []
    L.append("/* @@CC_DATA START — generated by critcareiq_update.py; REPL_RAW is kept from the existing file, everything else is regenerated ── */")
    L.append("const CC_MONTH_KEYS = [" + ",".join(f'"{k}"' for k in keys) + "];")
    L.append(f"const REPL_RAW = {repl_raw};                        // replacement level = pre-series physician CC rate (Nov '25–Mar '26 avg)")
    L.append("const CC_COUNTS = [  // [cases, CC charts (99291), Level 5 charts (99285)] per month — from Vituity E&M report")
    for n in names:
        parts = []
        for k in keys:
            v = billing[k].get(n)
            parts.append(f"{k}:" + (f"[{v[0]},{v[1]},{v[2]}]" if v else "null"))
        L.append(f'  {{ name:"{n}", ' + ", ".join(parts) + " },")
    L[-1] = L[-1].rstrip(",")
    L.append("];")
    L.append('/* ── Triage ESI mix per physician per month, from the weekly "pts per hour" sheet: [pts, ESI1, ESI2, ESI3, ESI4, ESI5, hours].')
    L.append("   ESI is nurse-assigned at triage, so an ESI-based pool is not physician-controlled. Weeks are assigned to the month holding most of their days. ── */")
    ekeys = [k for k in keys if k in esi]
    L.append("const ESI_COMPLETE = [" + ",".join(f'"{k}"' for k in ekeys if k in complete) + "];   // months with every week of the sheet filled in")
    L.append("const ESI_PARTIAL  = {" + ", ".join(f'{k}:"{partial[k]}"' for k in ekeys if k in partial) + "};")
    L.append("const ESI_COUNTS = [")
    rows = []
    for n in names:
        parts = []
        for k in ekeys:
            match = next((en for en in esi[k] if canon(en, [n]) == n), None)
            if match: v = esi[k][match]; parts.append(f"{k}:[{v[0]},{v[1]},{v[2]},{v[3]},{v[4]},{v[5]},{v[6]}]")
        if parts: rows.append(f'  {{ name:"{n}", ' + ", ".join(parts) + " }")
    L.append(",\n".join(rows))
    L.append("];")
    L.append("/* @@CC_DATA END */")
    return "\n".join(L), keys

def add_months(src, billing, keys):
    """Append new billing months to MONTHS / MONTH_KEYS / GROUP_RATES / PHYSICIANS rows."""
    mk = re.search(r"const MONTH_KEYS = \[(.*?)\];", src).group(1)
    existing = re.findall(r'"([a-z]{3}\d{2})"', mk)
    new = [k for k in keys if k not in existing]
    if not new: return src, []
    for k in new:
        y = int("20"+k[-2:]); mo = ABBR.index(k[:3].capitalize())+1; label = mlabel(y, mo)
        src = re.sub(r'(const MONTHS = \[.*?)(\];)', lambda m: m.group(1).replace('*"', '"') + f',"{label}*"' + m.group(2), src, count=1)
        src = re.sub(r'(const MONTH_KEYS = \[.*?)(\];)', lambda m: m.group(1) + f',"{k}"' + m.group(2), src, count=1)
        d = billing[k]; N = sum(v[0] for v in d.values()); CC = sum(v[1] for v in d.values())
        src = re.sub(r'(const GROUP_RATES = \{.*?)\s*\};', lambda m: m.group(1).rstrip() + f", {k}:{round(CC/N*100,1)} }};", src, count=1)
        def fix_row(m):
            name = m.group(1); v = d.get(name)
            rate = js_num(v[3]) if v else "null"; cases = js_num(v[0]) if v else "null"
            row = m.group(0)
            row = re.sub(r"(, cases:\{)", f", {k}:{rate}\\1", row, count=1)
            row = re.sub(r"(cases:\{[^}]*)(\})", lambda mm: mm.group(1) + f",{k}:{cases}" + mm.group(2), row, count=1)
            return row
        src = re.sub(r'\{ name:"([^"]+)",\s+[^\n]*?cases:\{[^}]*\}[^\n]*\}', fix_row, src)
        # physicians in billing but not in PHYSICIANS table
        tbl = re.search(r"const PHYSICIANS = \[(.*?)\n\];", src, re.S).group(1)
        for name in d:
            if f'name:"{name}"' not in tbl:
                print(f"  NOTE: {name} is in {k} billing but not in the PHYSICIANS table — add a row by hand (lowVol if appropriate).")
    return src, new

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--billing", required=True, nargs="+", help="folder(s) or .xlsx file(s) with Vituity E&M Billing reports (searched recursively)")
    ap.add_argument("--pph", help="weekly 'pts per hour physicians unblinded' .xlsx")
    ap.add_argument("--index", required=True, help="index.html to update")
    ap.add_argument("--out", help="write result here instead of overwriting --index")
    ap.add_argument("--write", action="store_true", help="actually write (default is a dry run)")
    ap.add_argument("--add-months", action="store_true", help="also append new billing months to the raw-rate tables")
    ap.add_argument("--esi-complete", nargs="*", help="override auto-detected complete ESI months, e.g. --esi-complete may26 jun26")
    args = ap.parse_args()

    files = []
    for b in args.billing:
        files += [b] if os.path.isfile(b) else glob.glob(os.path.join(b, "**", "*.xlsx"), recursive=True)
    files = [f for f in files if "billing" in os.path.basename(f).lower() and not os.path.basename(f).startswith("~$")]
    print(f"Reading {len(files)} billing file(s)…")
    billing = read_billing(files)
    if not billing: sys.exit("no billing months found")

    esi, complete, partial = ({}, [], {})
    if args.pph:
        print("Reading weekly ESI sheet…")
        esi, complete, partial = read_pph(args.pph)
        if args.esi_complete is not None: complete = args.esi_complete
        print(f"  ESI months: {sorted(esi)}  complete: {complete}  partial: {partial}")

    src = open(args.index, encoding="utf-8").read()
    m = re.search(r"const REPL_RAW = ([0-9.]+)", src); repl = m.group(1) if m else "0.045"
    block, keys = build_block(billing, esi, complete, partial, repl)
    if "/* @@CC_DATA START" not in src or "/* @@CC_DATA END */" not in src:
        sys.exit("index.html has no @@CC_DATA markers — this script expects the 10/17/26 or later index")
    s = src.index("/* @@CC_DATA START"); e = src.index("/* @@CC_DATA END */") + len("/* @@CC_DATA END */")
    new_src = src[:s] + block + src[e:]
    added = []
    if args.add_months:
        new_src, added = add_months(new_src, billing, keys)

    # summary for the newest month
    newest = keys[-1]; d = billing[newest]; N = sum(v[0] for v in d.values()); CC = sum(v[1] for v in d.values()); L5 = sum(v[2] for v in d.values())
    print(f"\nNewest billing month {newest}: {N} cases, {CC} CC, raw {CC/N*100:.1f}%, conversion {CC/(CC+L5)*100:.1f}%, needed for 8%: {0.08*N/(CC+L5)*100:.1f}%")
    if added: print(f"Added to raw-rate tables: {added}")
    print("Hand-update after this run if the month changed: PARTIAL_KEY, LATEST_FINAL, STATUS_KEYS, CODED, DATA_ASOF, header badge, footer.")

    if new_src == src:
        print("\nNo changes."); return
    if not args.write:
        print("\nDry run — rerun with --write to save."); return
    out = args.out or args.index
    if out == args.index:
        bak = args.index + ".bak"; open(bak, "w", encoding="utf-8").write(src); print(f"backup → {bak}")
    open(out, "w", encoding="utf-8").write(new_src)
    print(f"wrote → {out}")

if __name__ == "__main__":
    main()
