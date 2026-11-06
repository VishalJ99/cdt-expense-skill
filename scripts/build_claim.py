#!/usr/bin/env python3
"""Build one AI4Health CDT expense claim: the filled E1 workbook, the receipts ZIP and the email text.

    python3 build_claim.py claim.json            # writes ./out next to claim.json
    python3 build_claim.py claim.json --dry-run  # prints the plan and the email, writes nothing

Reads the claimant's details from ~/.config/cdt-expense/me.json and the blank E1
from ~/.config/cdt-expense/E1.xlsx.

The workbook is edited at XML level so the form's formulas, styles, validation and
namespace declarations survive. Only mapped cells change; every cached formula
result is dropped and Excel recalculates on open.

Standard library only, plus `pdftoppm` (poppler) for PDF receipts.
"""
import argparse
import datetime as dt
import html
import os
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
LAYOUT = HERE.parent / "layouts" / "imperial-e1.json"
CONFIG = Path(os.environ.get("CDT_EXPENSE_HOME", "~/.config/cdt-expense")).expanduser()
PURPOSE = "{surname}- AI4Health CDT-Conference expenses"  # guidance, How to claim expenses
ADMIN = "Britta Ross <b.ross@imperial.ac.uk>"  # AI4Health CDT administrator; every claim goes to her
ME_FIELDS = ["title", "first_name", "surname", "cid", "date_of_birth",
             "address", "postcode", "bank_name", "bank_branch", "sort_code", "account_number"]
JUNK = re.compile(r"(^|/)(__MACOSX|\.DS_Store|\._[^/]*)($|/)")


# ---------------------------------------------------------------- helpers

def safe(text):
    """Filename part: letters and digits only, words joined CamelCase."""
    words = re.findall(r"[A-Za-z0-9]+", str(text))
    return "".join(w[:1].upper() + w[1:] for w in words)


def iso_date(value):
    return dt.date.fromisoformat(value)


def excel_serial(day):
    return (day - dt.date(1899, 12, 30)).days


def load_me(path):
    if not path.exists():
        raise SystemExit(f"{path} not found. First run: copy examples/me.example.json there and fill it in.")
    me = json.loads(path.read_text(encoding="utf-8"))
    missing = [k for k in ME_FIELDS if not me.get(k)]
    if missing:
        raise SystemExit(f"{path} is missing: {', '.join(missing)}")
    return me


def full_name(me):
    return f"{me['title']} {me['first_name']} {me['surname']}"


def col_index(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def split_ref(ref):
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    if not m:
        raise ValueError("bad cell ref " + ref)
    return m.group(1), int(m.group(2))


# ---------------------------------------------------------------- plan

def plan_cells(claim, me, layout):
    """Return {cell_ref: (kind, value)} for every cell the build touches."""
    edits = {}
    sections = layout["sections"]

    # 1. Clear every input cell in every section, plus admin/approval cells.
    for sec in sections.values():
        for r in range(sec["first_row"], sec["last_row"] + 1):
            for col in sec["columns"].values():
                edits[f"{col}{r}"] = ("clear", None)
            if "formula" in sec:  # restore template formulas an admin hard-coded
                edits[f"{sec['formula_column']}{r}"] = ("formula", sec["formula"].format(r=r))
    for ref in layout["clear"]:
        edits[ref] = ("clear", None)

    # 2. Claimant details and purpose, from me.json.
    p = layout["personal"]
    edits[p["full_name"]] = ("str", full_name(me))
    cid = str(me["cid"])
    edits[p["cid"]] = ("num", int(cid)) if cid.isdigit() and not cid.startswith("0") else ("str", cid)
    edits[p["date_of_birth"]] = ("num", excel_serial(iso_date(me["date_of_birth"])))
    lines = me["address"] if isinstance(me["address"], list) else [me["address"]]
    if len(lines) > len(p["address"]):
        raise SystemExit(f"Address has {len(lines)} lines; the form has room for {len(p['address'])}.")
    for ref, line in zip(p["address"], lines):
        edits[ref] = ("str", line)
    for key in ("postcode", "bank_name", "bank_branch", "sort_code", "account_number"):
        edits[p[key]] = ("str", str(me[key]))  # text keeps leading zeros
    edits[layout["purpose_cell"]] = ("str", PURPOSE.format(surname=me["surname"]))
    if claim.get("sign"):  # only when the claimant has said to sign
        edits[layout["declaration"]["name"]] = ("str", full_name(me))
        edits[layout["declaration"]["date"]] = ("num", excel_serial(dt.date.today()))

    # 3. Rows, in order, per section.
    used = {name: 0 for name in sections}
    for row in claim["rows"]:
        name = row.get("section", "other")
        sec = sections[name]
        r = sec["first_row"] + used[name]
        if r > sec["last_row"]:
            raise SystemExit(
                f"Too many '{name}' rows for one form ({sec['last_row'] - sec['first_row'] + 1} max). "
                "Split into a second claim: one form, one receipts ZIP, one email."
            )
        used[name] += 1
        cols = sec["columns"]
        values = {
            "date": iso_date(row["date"]).strftime("%d/%m/%Y"),
            "type": row.get("type"),
            "description": row.get("description"),
            "mode": row.get("mode"),
            "return": row.get("return"),
            "from": row.get("from"),
            "to": row.get("to"),
            "foreign": None if row.get("currency", "GBP") == "GBP"
            else f"{float(row['amount']):.2f} {row['currency']}",
        }
        for key, col in cols.items():
            if key == "gbp":
                edits[f"{col}{r}"] = ("num", float(row["gbp"]))
            elif values.get(key) not in (None, ""):
                edits[f"{col}{r}"] = ("str", str(values[key]))
    return edits


def receipt_name(surname, row, suffix=""):
    d = iso_date(row["date"]).strftime("%Y%m%d")
    amount = f"{float(row['amount']):.2f}{row.get('currency', 'GBP')}"
    return f"{surname}_{d}_{safe(row['vendor'])}_{amount}{suffix}"


def output_names(me, claim):
    s = safe(me["surname"])
    c = "_".join(re.findall(r"[A-Za-z0-9]+", claim["claim"]))  # no dashes or symbols
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
    if kind == "formula":
        return f"<c{attrs}><f>{html.escape(value, quote=False)}</f></c>"
    raise ValueError(kind)


def insert_cell(xml, ref, kind, value):
    """Add a cell the template never materialised, in column order."""
    col, r = split_ref(ref)
    m = re.search(rf'<row\b[^>]*\br="{r}"[^>]*?(?:/>|>(.*?)</row>)', xml, re.S)
    if not m:
        raise SystemExit(f"Template has no row {r}; is the layout right for this form?")
    new = render_cell(f' r="{ref}"', kind, value)
    if m.group(0).endswith("/>"):
        row = m.group(0)[:-2] + ">" + new + "</row>"
        return xml[: m.start()] + row + xml[m.end():]
    body = m.group(1)
    pos = len(body)
    for c in re.finditer(r'<c\b[^>]*\br="([A-Z]+)\d+"', body):
        if col_index(c.group(1)) > col_index(col):
            pos = c.start()
            break
    start = m.start(1)
    return xml[: start + pos] + new + xml[start + pos:]


def edit_sheet(xml, edits):
    for ref, (kind, value) in edits.items():
        pat = re.compile(CELL.format(ref=re.escape(ref)), re.S)
        xml, n = pat.subn(lambda m: render_cell(m.group(1), kind, value), xml, count=1)
        if n == 0 and kind != "clear":  # a missing cell is already clear
            xml = insert_cell(xml, ref, kind, value)
    # Drop every cached formula result so old totals cannot survive.
    def drop_cache(m):
        c = m.group(0)
        return re.sub(r"<v(?:\s*/>|>.*?</v>)", "", c, flags=re.S) if "<f" in c else c
    return re.sub(r"<c\b[^>]*>(?:(?!</c>).)*</c>", drop_cache, xml, flags=re.S)


def force_recalc(xml):
    def fix(m):
        attrs = re.sub(r'\s+(?:calcMode|fullCalcOnLoad|forceFullCalc)="[^"]*"', "", m.group(1)).rstrip()
        return f'<calcPr{attrs} calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
    xml, n = re.subn(r"<calcPr\b([^>]*?)(?:\s*/>|></calcPr>)", fix, xml, count=1)
    if n == 0:
        xml, n = re.subn(r"</workbook>", '<calcPr fullCalcOnLoad="1" forceFullCalc="1"/></workbook>', xml)
    return xml


def build_workbook(template, out, edits, sheet):
    if not zipfile.is_zipfile(template):
        raise SystemExit(f"{template} is not an .xlsx. Open the E1 .xls in Microsoft Excel and "
                         "Save As 'Excel Workbook (.xlsx)' to that path.")
    with zipfile.ZipFile(template) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        if sheet not in names:
            raise SystemExit(f"{sheet} not in {template}; is it the E1 form?")
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == sheet:
                data = edit_sheet(data.decode("utf-8"), edits).encode("utf-8")
            elif info.filename == "xl/workbook.xml":
                data = force_recalc(data.decode("utf-8")).encode("utf-8")
            zout.writestr(info, data)


# ---------------------------------------------------------------- receipts

def to_jpgs(src, stem, dest, base=Path(".")):
    """Render a receipt as one or more JPGs named stem[_pN].jpg; return paths."""
    src = base / Path(src).expanduser()
    if not src.exists():
        raise SystemExit(f"Receipt not found: {src}")
    ext = src.suffix.lower()
    if ext == ".pdf":
        if not shutil.which("pdftoppm"):
            raise SystemExit("pdftoppm not found: install poppler (brew install poppler).")
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run(["pdftoppm", "-jpeg", "-r", "150", "-jpegopt", "quality=90",
                            str(src), f"{tmp}/page"], check=True)
            pages = sorted(Path(tmp).glob("page*.jpg"), key=lambda p: int(re.findall(r"\d+", p.stem)[-1]))
            out = []
            for i, page in enumerate(pages, 1):
                target = dest / (f"{stem}.jpg" if len(pages) == 1 else f"{stem}_p{i}.jpg")
                shutil.move(str(page), target)
                out.append(target)
            return out
    target = dest / f"{stem}.jpg"
    if ext in (".jpg", ".jpeg"):
        shutil.copyfile(src, target)
    elif shutil.which("sips"):
        subprocess.run(["sips", "-s", "format", "jpeg", str(src), "--out", str(target)],
                       check=True, stdout=subprocess.DEVNULL)
    else:
        try:
            from PIL import Image
        except ImportError:
            raise SystemExit(f"Cannot convert {src.name} to JPG: install Pillow or use macOS.")
        Image.open(src).convert("RGB").save(target, "JPEG", quality=90)
    return [target]


def pack_receipts(surname, claim, zip_path, workdir, base=Path(".")):
    surname = safe(surname)
    staging = workdir / "receipts"
    staging.mkdir(parents=True, exist_ok=True)
    for row in claim["rows"]:
        receipts = row.get("receipts") or ([row["receipt"]] if row.get("receipt") else [])
        if not receipts:
            raise SystemExit(f"Row {row['date']} {row.get('vendor')} has no receipt; itemised receipts are required.")
        for i, path in enumerate(receipts, 1):
            to_jpgs(path, receipt_name(surname, row, f"_{i}" if len(receipts) > 1 else ""), staging, base)
        if row.get("translation"):
            d = iso_date(row["date"]).strftime("%Y%m%d")
            (staging / f"{surname}_{d}_{safe(row['vendor'])}_translation.txt").write_text(
                row["translation"].strip() + "\n", encoding="utf-8")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(staging.iterdir()):
            if not JUNK.search(p.name):
                z.write(p, p.name)


# ---------------------------------------------------------------- email

def email(me, claim, xlsx, zip_name):
    rows = claim["rows"]
    total = sum(float(r["gbp"]) for r in rows)
    fx = [r for r in rows if r.get("currency", "GBP") != "GBP"]
    notes = [r["note"] for r in rows if r.get("note")]
    out = [f"To: {ADMIN}", f"Subject: {claim['claim']} expense claim", "",
           "Dear Britta,", "",
           f"Please find attached my expense claim for {claim['claim']}: the E1 form ({xlsx}) "
           f"and a ZIP of the matching receipts ({zip_name}).", "",
           f"Total claimed: £{total:.2f}"]
    if fx:
        out += ["", f"Exchange rates for each transaction date, from {claim.get('fx_source') or 'MISSING: set fx_source'}:"]
        out += [f"- {iso_date(r['date']).strftime('%d/%m/%Y')} {r['vendor']}: "
                f"{r['currency']} {float(r['amount']):.2f} = £{float(r['gbp']):.2f}" for r in fx]
    if notes:
        out += ["", "Notes:"] + [f"- {n}" for n in notes]
    out += ["", "Best wishes,", me["first_name"]]
    return "\n".join(out) + "\n"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("claim", help="claim JSON (see examples/claim.example.json)")
    ap.add_argument("--dry-run", action="store_true", help="print the plan and email, write nothing")
    args = ap.parse_args(argv)

    claim_path = Path(args.claim).resolve()
    claim = json.loads(claim_path.read_text(encoding="utf-8"))
    me = load_me(CONFIG / "me.json")
    layout = json.loads(LAYOUT.read_text(encoding="utf-8"))
    edits = plan_cells(claim, me, layout)
    xlsx_name, zip_name = output_names(me, claim)
    body = email(me, claim, xlsx_name, zip_name)

    if args.dry_run:
        for ref, (kind, value) in sorted(edits.items(), key=lambda kv: (split_ref(kv[0])[1], col_index(split_ref(kv[0])[0]))):
            if kind not in ("clear", "formula"):
                print(f"{ref:>5} {value}")
        for row in claim["rows"]:
            print("receipt:", receipt_name(safe(me["surname"]), row) + ".jpg")
        print("\n" + body)
        return 0

    out = claim_path.parent / "out"
    out.mkdir(exist_ok=True)
    build_workbook(CONFIG / "E1.xlsx", out / xlsx_name, edits, layout["sheet"])
    with tempfile.TemporaryDirectory() as tmp:
        pack_receipts(me["surname"], claim, out / zip_name, Path(tmp), claim_path.parent)
    (out / "email.txt").write_text(body, encoding="utf-8")

    with zipfile.ZipFile(out / xlsx_name) as z:
        assert z.testzip() is None
    with zipfile.ZipFile(out / zip_name) as z:
        assert z.testzip() is None
        print(f"{out / zip_name}:")
        for n in z.namelist():
            print("  ", n)
    print(out / xlsx_name)
    print("\n" + body)
    return 0


if __name__ == "__main__":
    sys.exit(main())
