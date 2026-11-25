"""Extract 'รายละเอียดตามคำอธิบาย' sheet -> data.js for the dashboard."""
import collections, json, re, sys, pathlib, openpyxl

SRC = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "../สรุปสัญญา_Printer.xlsx")
OUT = pathlib.Path(__file__).with_name("data.js")

wb = openpyxl.load_workbook(SRC, data_only=True, read_only=True)
# cost centre -> หน่วยงานหลัก
def norm_parent(p):
    # The workbook takes the text after the last "." in the unit name, so "ผจฟ.1กปบ." becomes
    # "1กปบ." and "...ผจฟ.1กปบ" becomes "1กปบ" -> both are really กปบ. Strip the leading
    # digit and restore the trailing dot.
    m = re.fullmatch(r"\d+(.+?)\.?", p or "")
    return m.group(1) + "." if m else p

parent = {r[1]: norm_parent(r[3]) for r in wb["สรุปจำนวนเครื่อง"].iter_rows(min_row=5, values_only=True)
          if isinstance(r[0], int)}

BRANDS = [("FUJI XEROX", "FUJI XEROX"), ("XEROX", "FUJI XEROX"), ("SAMSUNG", "SAMSUNG"), ("ML-45", "SAMSUNG"),
          ("LEXMARK", "LEXMARK"), ("MS521", "LEXMARK"), ("LASERJET", "HP"), ("DESIGNJET", "HP"), ('"HP"', "HP"),
          ("RICOH", "RICOH"), ("SP230", "RICOH"), ("EPSON", "EPSON"), ("LQ-", "EPSON"), ("TM-T", "EPSON"),
          ("BROTHER", "BROTHER"), ("FUJITSU", "FUJITSU"), ("DL3850", "FUJITSU"), ("FP-4", "FUJITSU"),
          ("KYOCERA", "KYOCERA"), ("CANON", "CANON"), ("CANNON", "CANON"), ("PLXMA", "CANON"),
          ("GENICOM", "TALLY GENICOM"), ("OKI", "OKI"), ("DATECS", "DATECS")]

def brand(d):
    u = d.upper()
    return next((b for k, b in BRANDS if k in u), "ไม่ระบุยี่ห้อ")

def ptype(d):
    u = d.upper()
    if "PLOTTER" in u or "DESIGNJET" in u: return "พล็อตเตอร์"
    if "บาร์โค้ด" in d: return "อื่นๆ"
    if any(k in u for k in ("THERMAL", "RFID", "POS", "TM-T")): return "เทอร์มอล / ฉลาก"
    if any(k in u for k in ("DOT", "MATRIX", "LINE PRINTER")): return "ดอทเมทริกซ์ / ไลน์"
    if "INKJET" in u: return "อิงค์เจ็ท"
    if "MULTI" in u: return "มัลติฟังก์ชัน"
    if "LASER" in u or "เลเซอร์" in d or "PRINTER" in u: return "เลเซอร์"
    return "อื่นๆ"

def contract(d):
    m = re.search(r"(บ|สต)\.?\s*(\d{1,3})\s*/\s*(\d{2,4})", d)
    if not m: return None, None
    y = int(m.group(3))
    if y < 100:
        if y < 40:  # e.g. "สต.034/20 กค.53" -> year written later
            m2 = re.search(r"\.\s*(\d{2})\s*$", d)
            y = int(m2.group(1)) if m2 else None
        y = 2500 + y if y else None
    return f"{m.group(1)}.{int(m.group(2))}/{y or m.group(3)}", y

# ---- abbreviations.csv: ตัวย่อ -> ชื่อเต็ม (see build_abbr.py) ----
import csv
key = lambda s: re.sub(r"[.\s]", "", s or "")
ABBR = {}
abbr_file = pathlib.Path(__file__).with_name("abbreviations.csv")
if abbr_file.exists():
    with abbr_file.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            ABBR[(key(row["ตัวย่อ"]), key(row["สังกัด"]))] = row["ชื่อเต็ม"].strip()
missing = collections.Counter()

def look(a, *ctxs):
    for c in ctxs:
        v = ABBR.get((key(a), key(c)))
        if v: return v
    return None

def office_of(name):
    m = re.search(r"(กฟ[สจย])\.([^.]+)\.?$", name)
    return (m.group(1), m.group(0)) if m else (None, None)

def expand_section(sec, ctx_kong, ctx_office):
    if not sec: return ""
    m = re.fullmatch(r"(สฟฟ\..+?)(ผจฟ\.\d)", sec)
    if m:  # สฟฟ.พล.1ผจฟ.1 -> station + section
        return f"{m.group(1)} {expand_section(m.group(2), ctx_kong, ctx_office)}"
    v = look(sec, ctx_kong or "", ctx_office or "", "")
    if not v: missing[sec] += 1
    return v or sec

def full_name(name, par):
    otype, otok = office_of(name)
    if "/" in name:  # ผชก.(น2)/สชก.(น2)
        return " / ".join(look(p, "") or p for p in name.split("/"))
    if otok:
        sec = name[: -len(otok)]
        ofull = look(otok, "") or otok
        return (expand_section(sec, None, otype) + " " + ofull).strip()
    tail = re.search(re.escape(par.rstrip(".")) + r"\.?$", name) if par else None
    if tail:
        # "ผจฟ.1กปบ." -> section "ผจฟ.1"; the digit belongs to the section, not the parent
        sec = name[: tail.start()]
        pfull = look(par, "") or par
        return (expand_section(sec, par, None) + " " + pfull).strip()
    return name

rows = []
for r in wb["รายละเอียดตามคำอธิบาย"].iter_rows(min_row=5, values_only=True):
    if not isinstance(r[0], int): continue
    _, cc, name, desc, qty, one, more, book = r
    c, y = contract(desc)
    par = parent.get(cc, "")
    rows.append([cc, name, par, desc, qty, one, more, round(book or 0, 2),
                 brand(desc), ptype(desc), c or "ไม่ระบุ", y, full_name(name, par)])

# full name of each หน่วยงานหลัก: กอง/ฝ่าย directly, otherwise the กฟจ./กฟส. it stands for
parent_full = {}
for p in {x[2] for x in rows}:
    kinds = {office_of(x[1])[0] for x in rows if x[2] == p}
    order = (["กฟจ."] if "กฟจ" in kinds else []) + ["กฟส.", "กฟจ."]
    parent_full[p] = look(p, "") or next((look(o + p, "") for o in order if look(o + p, "")), None) or look(p.rstrip(".") , "") or ""
for x in rows:
    x.append(parent_full[x[2]])
unresolved = sorted(p for p, v in parent_full.items() if not v)
if unresolved: print("หน่วยงานหลักที่ยังไม่มีชื่อเต็ม:", ", ".join(unresolved))
if missing: print("ตัวย่อแผนกที่ยังไม่มีชื่อเต็ม:", ", ".join(f"{k}({v})" for k, v in missing.most_common()))

total = sum(x[4] for x in rows)
assert total == 1742, total
OUT.write_text("window.PRINTER_DATA=" + json.dumps(
    {"cols": ["cc", "name", "parent", "desc", "qty", "one", "more", "book", "brand", "type", "contract", "year",
              "nameFull", "parentFull"],
     "rows": rows}, ensure_ascii=False, separators=(",", ":")) + ";\n", encoding="utf-8")
print(f"{len(rows)} rows, {total} units, book={sum(x[7] for x in rows):,.2f} -> {OUT}")
