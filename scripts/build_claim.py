#!/usr/bin/env python3
"""Build one AI4Health CDT expense claim: the filled E1 workbook, the receipts ZIP and the email text.

    python3 build_claim.py claim.json    # writes ./out next to claim.json

Reads ~/.config/cdt-expense/me.json and the blank ~/.config/cdt-expense/E1.xlsx.
The workbook is edited at XML level so the form's formulas, styles and validation
survive; cached formula results are dropped so Excel recalculates on open.
Needs `pdftoppm` (poppler) for PDF receipts.
"""
import datetime as dt
import html
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

CONFIG = Path("~/.config/cdt-expense").expanduser()
ADMIN = "Britta Ross <b.ross@imperial.ac.uk>"
PURPOSE = "{surname}- AI4Health CDT-Conference expenses"  # rule 19
SHEET = "xl/worksheets/sheet1.xml"

# Imperial E1 (non-employees) v1.2 cell map.
PERSONAL = {"cid": "D4", "date_of_birth": "I5", "postcode": "J10", "bank_name": "M5",
            "bank_branch": "M6", "sort_code": "M7", "account_number": "M9"}
NAME, ADDRESS, PURPOSE_CELL, SIGN_NAME, SIGN_DATE = "D3", ["G7", "G8", "G9", "G10"], "C6", "J63", "N63"
ADMIN_CELLS = ["E20", "J20", "C62", "F62", "C66", "F66", "J63", "N63"]  # cost centre, approvals, declaration
SECTIONS = {
    "travel": {"rows": range(13, 19),
               "columns": {"date": "C", "mode": "D", "return": "E", "from": "F", "to": "G", "foreign": "I", "gbp": "J"},
               "formula": 'IF(ISERROR(VLOOKUP(D{r}&K{r},S$12:U$23,2,FALSE)),"",VLOOKUP(D{r}&K{r},S$12:U$23,2,FALSE))'},
    "other": {"rows": range(41, 48),
              "columns": {"date": "C", "type": "D", "description": "E", "foreign": "I", "gbp": "J"},
              "formula": 'IF(AND(J{r}<>"",D{r}<>""),VLOOKUP(IF(OR(D{r}=W$24,D{r}=W$21),D{r}&K{r},D{r}),S$37:T$103,2,FALSE),"")'},
}


def safe(text):
    """Filename part: letters and digits only, words joined CamelCase (rule 27)."""
    return "".join(w[:1].upper() + w[1:] for w in re.findall(r"[A-Za-z0-9]+", str(text)))


def day(row):
    return dt.date.fromisoformat(row["date"])


def excel_serial(d):
    return (d - dt.date(1899, 12, 30)).days


def col_index(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def full_name(me):
    return f"{me['title']} {me['first_name']} {me['surname']}"


def plan_cells(claim, me):
    """Return {cell_ref: (kind, value)} for every cell the build touches."""
    edits = {}
    # Clear every input row and admin cell; restore analysis-code formulas (rules 18, 21).
    for sec in SECTIONS.values():
        for r in sec["rows"]:
            for col in sec["columns"].values():
                edits[f"{col}{r}"] = ("clear", None)
            edits[f"L{r}"] = ("formula", sec["formula"].format(r=r))
    for ref in ADMIN_CELLS:
        edits[ref] = ("clear", None)

    edits[NAME] = ("str", full_name(me))
    cid = str(me["cid"])
    edits[PERSONAL["cid"]] = ("num", int(cid)) if cid.isdigit() and not cid.startswith("0") else ("str", cid)
    edits[PERSONAL["date_of_birth"]] = ("num", excel_serial(dt.date.fromisoformat(me["date_of_birth"])))
    for ref, line in zip(ADDRESS, me["address"]):
        edits[ref] = ("str", line)
    for key in ("postcode", "bank_name", "bank_branch", "sort_code", "account_number"):
        edits[PERSONAL[key]] = ("str", str(me[key]))  # text keeps leading zeros
    edits[PURPOSE_CELL] = ("str", PURPOSE.format(surname=me["surname"]))
    if claim.get("sign"):
        edits[SIGN_NAME] = ("str", full_name(me))
        edits[SIGN_DATE] = ("num", excel_serial(dt.date.today()))

    used = {name: 0 for name in SECTIONS}
    for row in claim["rows"]:
        name = row.get("section", "other")
        sec = SECTIONS[name]
        if used[name] == len(sec["rows"]):
            raise SystemExit(f"Too many '{name}' rows for one form ({len(sec['rows'])} max). "
                             "Split into a second claim: one form, one receipts ZIP, one email.")
        r = sec["rows"][used[name]]
        used[name] += 1
        values = dict(row, date=day(row).strftime("%d/%m/%Y"),
                      foreign=None if row["currency"] == "GBP" else f"{float(row['amount']):.2f} {row['currency']}")
        for key, col in sec["columns"].items():
            if key == "gbp":
                edits[f"{col}{r}"] = ("num", float(row["gbp"]))
            elif values.get(key):
                edits[f"{col}{r}"] = ("str", str(values[key]))
    return edits


def output_names(me, claim):
    s, c = safe(me["surname"]), "_".join(re.findall(r"[A-Za-z0-9]+", claim["claim"]))
    return f"{s}_E1_{c}.xlsx", f"{s}_{c}_receipts.zip"


# ---------------------------------------------------------------- workbook

CELL = r'<c\b(?=[^>]*\br="{ref}")([^>]*?)(?:\s*/>|>(.*?)</c>)'


def render_cell(attrs, kind, value):
    attrs = re.sub(r'\s+t="[^"]*"', "", attrs).rstrip()
    if kind == "clear":
        return f"<c{attrs}/>"
    if kind == "str":
        return f'<c{attrs} t="inlineStr"><is><t xml:space="preserve">{html.escape(value)}</t></is></c>'
    if kind == "num":
        return f"<c{attrs}><v>{value:g}</v></c>"
    return f"<c{attrs}><f>{html.escape(value, quote=False)}</f></c>"


def insert_cell(xml, ref, kind, value):
    """Add a cell the template never materialised, in column order."""
    col, r = re.fullmatch(r"([A-Z]+)(\d+)", ref).groups()
    m = re.search(rf'<row\b[^>]*\br="{r}"[^>]*?(?:/>|>(.*?)</row>)', xml, re.S)
    new = render_cell(f' r="{ref}"', kind, value)
    if m.group(0).endswith("/>"):
        return xml[: m.start()] + m.group(0)[:-2] + ">" + new + "</row>" + xml[m.end():]
    body, pos = m.group(1), len(m.group(1))
    for c in re.finditer(r'<c\b[^>]*\br="([A-Z]+)\d+"', body):
        if col_index(c.group(1)) > col_index(col):
            pos = c.start()
            break
    at = m.start(1) + pos
    return xml[:at] + new + xml[at:]


def edit_sheet(xml, edits):
    for ref, (kind, value) in edits.items():
        pat = re.compile(CELL.format(ref=re.escape(ref)), re.S)
        xml, n = pat.subn(lambda m: render_cell(m.group(1), kind, value), xml, count=1)
        if n == 0 and kind != "clear":  # a missing cell is already clear
            xml = insert_cell(xml, ref, kind, value)

    def drop_cache(m):  # so no stale total survives
        c = m.group(0)
        return re.sub(r"<v(?:\s*/>|>.*?</v>)", "", c, flags=re.S) if "<f" in c else c
    return re.sub(r"<c\b[^>]*>(?:(?!</c>).)*</c>", drop_cache, xml, flags=re.S)


def force_recalc(xml):
    def fix(m):
        attrs = re.sub(r'\s+(?:calcMode|fullCalcOnLoad|forceFullCalc)="[^"]*"', "", m.group(1)).rstrip()
        return f'<calcPr{attrs} calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
    return re.sub(r"<calcPr\b([^>]*?)(?:\s*/>|></calcPr>)", fix, xml, count=1)


def build_workbook(template, out, edits):
    if not zipfile.is_zipfile(template):
        raise SystemExit(f"{template} is not an .xlsx. Open the E1 .xls in Microsoft Excel and "
                         "Save As 'Excel Workbook (.xlsx)' to that path.")
    with zipfile.ZipFile(template) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == SHEET:
                data = edit_sheet(data.decode("utf-8"), edits).encode("utf-8")
            elif info.filename == "xl/workbook.xml":
                data = force_recalc(data.decode("utf-8")).encode("utf-8")
            zout.writestr(info, data)


# ---------------------------------------------------------------- receipts

def to_jpgs(src, stem, dest):
    """Render a receipt as JPGs named stem[_pN].jpg, one per page (rule 25)."""
    if src.suffix.lower() == ".pdf":
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run(["pdftoppm", "-jpeg", "-r", "150", "-jpegopt", "quality=90",
                            str(src), f"{tmp}/page"], check=True)
            pages = sorted(Path(tmp).glob("page*.jpg"), key=lambda p: int(re.findall(r"\d+", p.stem)[-1]))
            for i, page in enumerate(pages, 1):
                shutil.move(str(page), dest / (f"{stem}.jpg" if len(pages) == 1 else f"{stem}_p{i}.jpg"))
    elif src.suffix.lower() in (".jpg", ".jpeg"):
        shutil.copyfile(src, dest / f"{stem}.jpg")
    else:
        subprocess.run(["sips", "-s", "format", "jpeg", str(src), "--out", str(dest / f"{stem}.jpg")],
                       check=True, stdout=subprocess.DEVNULL)


def pack_receipts(me, claim, zip_path, base):
    surname = safe(me["surname"])
    with tempfile.TemporaryDirectory() as tmp:
        staging = Path(tmp)
        for row in claim["rows"]:
            stem = f"{surname}_{day(row):%Y%m%d}_{safe(row['vendor'])}"
            to_jpgs(base / Path(row["receipt"]).expanduser(), f"{stem}_{float(row['amount']):.2f}{row['currency']}",
                    staging)
            if row.get("translation"):  # rule 24
                (staging / f"{stem}_translation.txt").write_text(row["translation"].strip() + "\n", encoding="utf-8")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for p in sorted(staging.iterdir()):
                z.write(p, p.name)


# ---------------------------------------------------------------- email

def email(me, claim, xlsx, zip_name):
    rows = claim["rows"]
    fx = [r for r in rows if r["currency"] != "GBP"]
    notes = [r["note"] for r in rows if r.get("note")]
    out = [f"To: {ADMIN}", f"Subject: {claim['claim']} expense claim", "",
           "Dear Britta,", "",
           f"Please find attached my expense claim for {claim['claim']}: the E1 form ({xlsx}) "
           f"and a ZIP of the matching receipts ({zip_name}).", "",
           f"Total claimed: £{sum(float(r['gbp']) for r in rows):.2f}"]
    if fx:  # rule 22
        out += ["", f"Exchange rates for each transaction date, from {claim['fx_source']}:"]
        out += [f"- {day(r):%d/%m/%Y} {r['vendor']}: {r['currency']} {float(r['amount']):.2f} = £{float(r['gbp']):.2f}"
                for r in fx]
    if notes:
        out += ["", "Notes:"] + [f"- {n}" for n in notes]
    out += ["", "Best wishes,", me["first_name"]]
    return "\n".join(out) + "\n"


def main(argv=None):
    claim_path = Path((argv or sys.argv[1:])[0]).resolve()
    claim = json.loads(claim_path.read_text(encoding="utf-8"))
    me = json.loads((CONFIG / "me.json").read_text(encoding="utf-8"))
    xlsx_name, zip_name = output_names(me, claim)
    edits = plan_cells(claim, me)
    body = email(me, claim, xlsx_name, zip_name)

    out = claim_path.parent / "out"
    out.mkdir(exist_ok=True)
    build_workbook(CONFIG / "E1.xlsx", out / xlsx_name, edits)
    pack_receipts(me, claim, out / zip_name, claim_path.parent)
    (out / "email.txt").write_text(body, encoding="utf-8")

    print(out / xlsx_name)
    with zipfile.ZipFile(out / zip_name) as z:
        print(f"{out / zip_name}:", *z.namelist(), sep="\n   ")
    print("\n" + body)


if __name__ == "__main__":
    main()
