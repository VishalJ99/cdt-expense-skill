"""Tests with a synthetic stand-in for the E1 form and dummy receipts only.

    python3 -m unittest discover -s tests -v
"""
import json
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import build_claim  # noqa: E402

NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac"'


def fake_template(path):
    """A tiny workbook shaped like the E1: stale values, a hard-coded analysis code, old totals."""
    rows = {
        6: '<c r="C6" t="inlineStr"><is><t>Old purpose</t></is></c>',
        13: '<c r="C13"><v>46266</v></c><c r="D13" t="inlineStr"><is><t>Rail</t></is></c>'
            '<c r="J13"><v>157.83</v></c><c r="L13"><f>IF(AND(J13&lt;&gt;"",D13&lt;&gt;""),999999,"")</f><v>999999</v></c>',
        19: '<c r="J19"><f>SUM(J13:J18)</f><v>157.83</v></c>',
        20: '<c r="E20" t="inlineStr"><is><t>OLDCC</t></is></c><c r="J20" t="inlineStr"><is><t>OLDPROJ</t></is></c>',
        41: '<c r="C41" s="3"/><c r="D41" s="3"/><c r="K41"><v>0</v></c>',
        42: '<c r="K42"><v>0</v></c>',
        48: '<c r="J48"><f>SUM(J41:J47)</f><v>55</v></c>',
        62: '<c r="C62" t="inlineStr"><is><t>Approver, A</t></is></c><c r="F62"><v>46267</v></c>',
        63: '<c r="J63" t="inlineStr"><is><t>Old Claimant</t></is></c><c r="N63"><v>46266</v></c>',
    }
    for r in [3, 4, 5, 7, 8, 9, 10] + list(range(14, 19)) + list(range(43, 48)) + [66]:
        rows.setdefault(r, "")
    body = "".join(f'<row r="{r}">{c}</row>' if c else f'<row r="{r}"/>' for r, c in sorted(rows.items()))
    sheet = f'<?xml version="1.0" encoding="UTF-8"?><worksheet {NS} {MC}><sheetData>{body}</sheetData></worksheet>'
    wb = f'<?xml version="1.0" encoding="UTF-8"?><workbook {NS}><sheets/><calcPr calcId="191029" calcMode="manual"/></workbook>'
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/worksheets/sheet1.xml", sheet)


def fake_pdf(path, pages=1):
    objs = ["<< /Type /Catalog /Pages 2 0 R >>"]
    kids = " ".join(f"{3 + i} 0 R" for i in range(pages))
    objs.append(f"<< /Type /Pages /Kids [{kids}] /Count {pages} >>")
    for _ in range(pages):
        objs.append("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] >>")
    out, offsets = b"%PDF-1.4\n", []
    for i, o in enumerate(objs, 1):
        offsets.append(len(out))
        out += f"{i} 0 obj {o} endobj\n".encode()
    xref = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    out += b"".join(f"{o:010} 00000 n \n".encode() for o in offsets)
    out += f"trailer << /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    Path(path).write_bytes(out)


@unittest.skipUnless(shutil.which("pdftoppm"), "needs poppler's pdftoppm")
class BuildClaimTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp)
        (self.tmp / "receipts").mkdir()
        config = self.tmp / "config"
        config.mkdir()
        fake_template(config / "E1.xlsx")
        shutil.copyfile(ROOT / "examples" / "me.example.json", config / "me.json")
        build_claim.CONFIG = config
        fake_pdf(self.tmp / "receipts" / "rail.pdf")
        fake_pdf(self.tmp / "receipts" / "hotel.pdf", pages=2)
        subprocess.run(["pdftoppm", "-jpeg", "-singlefile", str(self.tmp / "receipts" / "rail.pdf"),
                        str(self.tmp / "receipts" / "dinner")], check=True)
        claim = json.loads((ROOT / "examples" / "claim.example.json").read_text())
        claim["rows"].append({"date": "2026-09-27", "type": "HOTEL / SUBSISTENCE", "description": "Hotel, 1 night",
                              "vendor": "Hotel Example", "amount": "80", "currency": "EUR", "gbp": 68.9,
                              "receipt": "receipts/hotel.pdf"})
        (self.tmp / "claim.json").write_text(json.dumps(claim))
        build_claim.main([str(self.tmp / "claim.json")])
        self.out = self.tmp / "out"

    def sheet(self):
        with zipfile.ZipFile(self.out / "Doe_E1_ExampleConf_2026_post_trip.xlsx") as z:
            return z.read("xl/worksheets/sheet1.xml").decode(), z.read("xl/workbook.xml").decode()

    def cell(self, xml, ref):
        m = re.search(rf'<c\b[^>]*r="{ref}"[^>]*?(?:/>|>(.*?)</c>)', xml, re.S)
        return m.group(0) if m else ""

    def test_values_written_and_old_values_gone(self):
        xml, _ = self.sheet()
        self.assertIn("Doe- AI4Health CDT-Conference expenses", self.cell(xml, "C6"))
        self.assertIn("Ms Jane Doe", self.cell(xml, "D3"))
        self.assertIn("<v>36191</v>", self.cell(xml, "I5"))  # 1999-01-31 as an Excel date
        self.assertIn("00000000</t>", self.cell(xml, "M9"))  # account number kept as text
        self.assertIn("SW1A 0AA", self.cell(xml, "J10"))
        self.assertNotIn("Ms Jane Doe", self.cell(xml, "J63"))  # not signed unless asked
        self.assertIn("26/09/2026", self.cell(xml, "C41"))
        self.assertIn('s="3"', self.cell(xml, "C41"))  # style kept
        self.assertIn("<v>18.7</v>", self.cell(xml, "J41"))
        self.assertIn("21.80 EUR", self.cell(xml, "I41"))
        self.assertIn("<v>120.5</v>", self.cell(xml, "J13"))
        self.assertIn("Hotel, 1 night", self.cell(xml, "E42"))  # inserted cell
        for gone in ("Old purpose", "OLDCC", "OLDPROJ", "Approver", "Old Claimant", "157.83"):
            self.assertNotIn(gone, xml)
        self.assertIn('xmlns:mc=', xml)  # compatibility namespaces preserved

    def test_cell_order_within_row(self):
        xml, _ = self.sheet()
        row42 = re.search(r'<row r="42">(.*?)</row>', xml, re.S).group(1)
        cols = re.findall(r'<c r="([A-Z]+)42"', row42)
        self.assertEqual(cols, sorted(cols, key=build_claim.col_index))

    def test_formulas_restored_and_caches_dropped(self):
        xml, wb = self.sheet()
        self.assertNotIn("999999", xml)
        self.assertIn("VLOOKUP(D13&amp;K13", self.cell(xml, "L13"))
        self.assertIn("VLOOKUP(IF(OR(D41=W$24", self.cell(xml, "L41"))
        for c in re.findall(r"<c\b[^>]*>(?:(?!</c>).)*</c>", xml, re.S):
            if "<f>" in c:
                self.assertNotIn("<v>", c)
        self.assertIn('fullCalcOnLoad="1"', wb)
        self.assertNotIn('calcMode="manual"', wb)

    def test_receipts_zip(self):
        with zipfile.ZipFile(self.out / "Doe_ExampleConf_2026_post_trip_receipts.zip") as z:
            names = sorted(z.namelist())
        self.assertEqual(names, [
            "Doe_20260925_ExampleRail_120.50GBP.jpg",
            "Doe_20260926_CafeExample_21.80EUR.jpg",
            "Doe_20260926_CafeExample_translation.txt",
            "Doe_20260927_HotelExample_80.00EUR_p1.jpg",
            "Doe_20260927_HotelExample_80.00EUR_p2.jpg",
        ])

    def test_email(self):
        text = (self.out / "email.txt").read_text()
        self.assertIn("To: cdt-admin@example.ac.uk", text)
        self.assertIn("Subject: ExampleConf 2026 post trip expense claim", text)
        self.assertIn("Total claimed: £208.10", text)
        self.assertIn("EUR 21.80 = £18.70", text)
        self.assertIn("tip excluded", text)
        self.assertTrue(text.rstrip().endswith("Jane"))


class PlanTest(unittest.TestCase):
    def test_too_many_rows_says_split(self):
        layout = json.loads((ROOT / "layouts" / "imperial-e1.json").read_text())
        row = {"date": "2026-01-01", "type": "X", "vendor": "V", "amount": "1", "gbp": 1}
        with self.assertRaises(SystemExit) as e:
            me = json.loads((ROOT / "examples" / "me.example.json").read_text())
            build_claim.plan_cells({"claim": "c", "rows": [row] * 8}, me, layout)
        self.assertIn("second claim", str(e.exception))


if __name__ == "__main__":
    unittest.main()
