# cdt-expense-skill

An agent skill (Claude Code, Codex, OpenClaw or any Agent Skills runner) for UK doctoral training centre (CDT) students who claim conference and travel expenses on Imperial College's **E1 Excel form**. It covers the whole job:

1. finds the claims guidance and your administrator's latest instructions in your mail
2. downloads the receipts
3. fills a copy of the original E1 workbook without breaking it
4. renders receipts to JPG and names them `Surname_YYYYMMDD_Vendor_12.34EUR.jpg` in one ZIP per claim
5. drafts the email, checks it, and sends it only after you approve

## Install

Claude Code:

```sh
git clone https://github.com/VishalJ99/cdt-expense-skill ~/.claude/skills/cdt-expense
```

Other agents: clone it into that agent's skills directory (e.g. `~/.codex/skills/cdt-expense`), or use `npx skills add VishalJ99/cdt-expense-skill`.

Then ask your agent something like *"claim my conference expenses from the receipts in this thread"*.

## Requirements

- Python 3.9+ (standard library only)
- [poppler](https://poppler.freedesktop.org/) for `pdftoppm`: `brew install poppler` or `apt install poppler-utils`
- [Surface CLI](https://github.com/VishalJ99/surface-cli) to read and send Outlook or Gmail mail: `npm install -g surface-cli`. Any mail CLI with search, attachment download and send-with-attachments will also work.
- Microsoft Excel, optional but recommended for the final check
- **The E1 template itself is not included.** Get the current `.xlsx` from your CDT administrator or Imperial's Finance forms page.

## The builder on its own

```sh
python3 scripts/build_claim.py --template E1.xlsx --claim examples/claim.example.json --dry-run
python3 scripts/build_claim.py --template E1.xlsx --claim my-claim.json --out-dir out
```

It writes `out/Surname_E1_<claim>.xlsx`, `out/Surname_<claim>_receipts.zip` and `out/summary.txt` (total and FX lines for the email).

Why it edits XML rather than using openpyxl: re-saving the form drops data validation, compatibility namespaces and lookup formulas, and administrators have rejected forms because of it. The script changes only the mapped cells and restores the analysis-code formulas. It also clears every cached total and makes Excel recalculate on open.

Cell positions live in `layouts/imperial-e1.json`. Edit that file for a different form version.

## Tests

```sh
python3 -m unittest discover -s tests -v
```

The tests use a synthetic stand-in workbook and generated dummy receipts; no real data.

## Licence

MIT
