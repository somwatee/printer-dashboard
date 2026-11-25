"""Build abbreviations.csv (ตัวย่อ -> ชื่อเต็ม) from 'Computer Nameสค69.xlsx'.

Only unit names and abbreviations are copied. IP addresses and network-device details
in the source workbook are never read into the output.

Columns:
  ตัวย่อ   abbreviation as used in the asset data, e.g. กดส.  ผคข.  กฟส.ตห.
  สังกัด    context the abbreviation is valid in:
            ""        = global (กอง/ฝ่าย/กฟส./กฟจ.)
            "กดส." etc = a section (แผนก) belonging to that กอง
            "กฟส"/"กฟจ" = a section inside a branch office
  ชื่อเต็ม  full Thai name
  ที่มา     sheet it came from, or "เพิ่มเอง" for manual rows

Rows marked ที่มา=เพิ่มเอง in an existing abbreviations.csv are kept when rebuilding,
so manual additions/corrections survive a re-run.
"""
import csv, re, sys, pathlib, collections, openpyxl

SRC = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "../Computer Nameสค69.xlsx")
OUT = pathlib.Path(__file__).with_name("abbreviations.csv")
MANUAL = "เพิ่มเอง"

def cl(s):
    return re.sub(r"\s+", " ", str(s)).strip() if s is not None else ""

def abbr(s):
    """'กฟส.รก (S)' -> 'กฟส.รก.'; 'ผจฟ.1' stays; ensure a trailing dot after Thai letters."""
    s = re.sub(r"\s*\((?:L|M|S|XS)\)\s*$", "", cl(s)).replace(" ", "")
    return s + "." if re.search(r"[ก-๙]$", s) else s

def strip_num(n):
    return re.sub(r"^\(?\d+\)\s*", "", n).rstrip("*").strip()

wb = openpyxl.load_workbook(SRC, data_only=True, read_only=True)
rows = {}  # (abbr, context) -> (name, source)

def put(a, ctx, name, src, override=False):
    if a and name and ((a, ctx) not in rows or override):
        rows[(a, ctx)] = (name, src)

# 1) ComName sheet: hierarchy in columns A-E, abbreviation in F
sh = "ComNameเขต(ใหม่)"
cur_kong = None
for r in wb[sh].iter_rows(min_row=2, values_only=True):
    lv = next((i for i in range(5) if r[i]), None)
    if lv is None: continue
    name, a = strip_num(cl(r[lv])), abbr(r[5])
    if not a or "...." in a or "…" in name or "...." in name:
        if lv <= 3: cur_kong = None
        continue
    if lv <= 3:
        put(a, "", re.sub(r"\s*\(ขึ้นตรงกับ.*\)$", "", name), sh)
        cur_kong = a if lv == 3 else None
    elif cur_kong:
        put(a, cur_kong, name, sh)

# "(น2)" used as a main unit in the asset data = the region office (กฟน.2)
for (a, ctx), (n, src) in list(rows.items()):
    m = re.fullmatch(r"กฟ([นฉกต])\.(\d)", a)
    if m and not ctx:
        put(f"({m.group(1)}{m.group(2)})", "", n, sh)

# 2) Branch offices: "การไฟฟ้าส่วนภูมิภาคสาขาเมืองแพร่ (กฟส.พร.)"
sh = "กฟส. (L, M, S, XS)"
for r in wb[sh].iter_rows(values_only=True):
    m = re.fullmatch(r"(การไฟฟ้าส่วนภูมิภาคสาขา(.+?))\s*\((กฟส\.[^)]+)\)(?:\s*พิเศษ)?", cl(r[2]))
    if not m: continue
    a = abbr(m.group(3))
    put(a, "", m.group(1), sh)
    # กฟจ. of a province shares the abbreviation of its "เมือง" branch
    if m.group(2).startswith("เมือง"):
        put(a.replace("กฟส.", "กฟจ.", 1), "", "การไฟฟ้าส่วนภูมิภาคจังหวัด" + m.group(2)[5:], sh)

# 3) Sections inside branch offices (most common name per abbreviation)
votes = collections.defaultdict(collections.Counter)
for sh in ("กฟฟ.สาขา (ขนาด M-L-S)", "กฟฟ.สาขา (ขนาด M-L-S) (2)"):
    for r in wb[sh].iter_rows(values_only=True):
        n, a = cl(r[2]), abbr(r[3])
        if re.match(r"^\d+\)\s*แผนก", n) and re.fullmatch(r"ผ[ก-๙]+\.", a):
            votes[a][re.sub(r"\s*\(ถ้ามี\)", "", strip_num(n))] += 1
for a, c in votes.items():
    for ctx in ("กฟส", "กฟจ"):
        put(a, ctx, c.most_common(1)[0][0], "กฟฟ.สาขา (ขนาด M-L-S)")
# old-structure sections still used in the asset data for S-size branches
for a, n in {"ผบง.": "แผนกบัญชีและการเงิน", "ผปร.": "แผนกปฏิบัติการระบบไฟฟ้า"}.items():
    for ctx in ("กฟส", "กฟจ"):
        put(a, ctx, n, "กฟฟ.สาขา (ขนาด M-L-S) (2)")

# keep manual rows from a previous file (they win over generated ones)
if OUT.exists():
    with OUT.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if row.get("ที่มา") == MANUAL:
                put(row["ตัวย่อ"], row["สังกัด"], row["ชื่อเต็ม"], MANUAL, override=True)

with OUT.open("w", encoding="utf-8-sig", newline="") as f:
    w = csv.writer(f)
    w.writerow(["ตัวย่อ", "สังกัด", "ชื่อเต็ม", "ที่มา"])
    for (a, ctx), (n, src) in sorted(rows.items(), key=lambda x: (x[0][1] != "", x[0][1], x[0][0])):
        w.writerow([a, ctx, n, src])
print(f"{len(rows)} abbreviations -> {OUT}")
