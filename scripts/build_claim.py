#!/usr/bin/env python3
"""Build one expense claim: a filled copy of the E1 workbook plus a receipts ZIP.

    python3 build_claim.py --template E1.xlsx --claim claim.json --out-dir out/
    python3 build_claim.py --template E1.xlsx --claim claim.json --dry-run

The workbook is edited at XML level so the template's formulas, styles, data
validation and namespace declarations survive untouched. Only the cells named in
the layout file are changed. Every formula cache is dropped and Excel is told to
recalculate on open, so no total from a previous claim can survive.

Standard library only. Receipts need `pdftoppm` (poppler) for PDFs and either
`sips` (macOS) or Pillow for non-JPG images.
"""
import argparse
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

HERE = Path(__file__).resolve().parent
DEFAULT_LAYOUT = HERE.parent / "layouts" / "imperial-e1.json"
JUNK = re.compile(r"(^|/)(__MACOSX|\.DS_Store|\._[^/]*)($|/)")


# ---------------------------------------------------------------- helpers

def safe(text):
    """Filename part: letters and digits only, words joined CamelCase."""
    words = re.findall(r"[A-Za-z0-9]+", str(text))
    return "".join(w[:1].upper() + w[1:] for w in words)


def iso_date(value):
    return dt.date.fromisoformat(value)


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

def plan_cells(claim, layout):
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

    # 2. Header fields.
    edits[layout["purpose_cell"]] = ("str", claim["purpose"])
    decl = claim.get("declaration")
    if decl:  # only when the claimant has authorised signing
        edits[layout["declaration"]["name"]] = ("str", decl["name"])
        edits[layout["declaration"]["date"]] = ("str", decl["date"])

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


def output_names(claim):
    s = safe(claim["surname"])
    c = re.sub(r"[^A-Za-z0-9]+", "_", claim["claim_name"]).strip("_")  # no dashes or symbols
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
        raise SystemExit("Template is not an .xlsx. Open the .xls in Microsoft Excel and "
                         "Save As 'Excel Workbook (.xlsx)', then pass that file.")
    with zipfile.ZipFile(template) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        if sheet not in names:
            raise SystemExit(f"{sheet} not in template; set 'sheet' in the layout file.")
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


def pack_receipts(claim, zip_path, workdir, base=Path(".")):
    surname = safe(claim["surname"])
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


# ---------------------------------------------------------------- summary

def summary(claim, xlsx, zip_name):
    total = sum(float(r["gbp"]) for r in claim["rows"])
    lines = [f"Workbook: {xlsx}", f"Receipts: {zip_name}", f"Total: £{total:.2f}", ""]
    for r in claim["rows"]:
        line = f"- {r['date']} {r['vendor']}: £{float(r['gbp']):.2f}"
        if r.get("currency", "GBP") != "GBP":
            line += f" ({float(r['amount']):.2f} {r['currency']})"
        lines.append(line)
    fx = [r for r in claim["rows"] if r.get("currency", "GBP") != "GBP"]
    if fx:
        lines += ["", f"FX basis: {claim.get('fx_source', 'MISSING: add fx_source before sending')}"]
    return "\n".join(lines) + "\n"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--template", required=True, help="the original E1 .xlsx from your administrator")
    ap.add_argument("--claim", required=True, help="claim JSON (see examples/claim.example.json)")
    ap.add_argument("--layout", default=str(DEFAULT_LAYOUT), help="cell layout JSON for the form")
    ap.add_argument("--out-dir", default="out")
    ap.add_argument("--dry-run", action="store_true", help="print the plan, write nothing")
    args = ap.parse_args(argv)

    claim = json.loads(Path(args.claim).read_text(encoding="utf-8"))
    layout = json.loads(Path(args.layout).read_text(encoding="utf-8"))
    edits = plan_cells(claim, layout)
    xlsx_name, zip_name = output_names(claim)

    if args.dry_run:
        for ref, (kind, value) in sorted(edits.items(), key=lambda kv: (split_ref(kv[0])[1], col_index(split_ref(kv[0])[0]))):
            if kind != "clear":
                print(f"{ref:>5} {kind:7} {value}")
        for row in claim["rows"]:
            print("receipt:", receipt_name(safe(claim["surname"]), row) + ".jpg")
        print(summary(claim, xlsx_name, zip_name))
        return 0

    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    build_workbook(args.template, out / xlsx_name, edits, layout.get("sheet", "xl/worksheets/sheet1.xml"))
    with tempfile.TemporaryDirectory() as tmp:
        pack_receipts(claim, out / zip_name, Path(tmp), Path(args.claim).resolve().parent)
    (out / "summary.txt").write_text(summary(claim, xlsx_name, zip_name), encoding="utf-8")

    with zipfile.ZipFile(out / xlsx_name) as z:
        assert z.testzip() is None
    with zipfile.ZipFile(out / zip_name) as z:
        assert z.testzip() is None
        print(f"{out / zip_name}:")
        for n in z.namelist():
            print("  ", n)
    print(f"{out / xlsx_name}")
    print((out / "summary.txt").read_text())
    return 0


if __name__ == "__main__":
    sys.exit(main())
