# cdt-expense-skill

An agent skill (Claude Code, Codex, OpenClaw or any Agent Skills runner) for students of Imperial's **UKRI AI4Health CDTs**: the CDT in AI for Healthcare and the AI CDT in Digital Healthcare. It handles conference approval requests and expense claims end to end, with the CDT's rules built in, so nobody has to dig out the guidance again.

**What's built in**

- `references/ai4health-guidance.md`: the 2025-26 AI4Health conference approval and expense guidance as 29 sourced rules. It covers approval, the £3,000 cap, eligible dates, taxis, no alcohol, itemised and translated receipts, JPGs in one ZIP per claim, surname-first filenames without dashes, purpose-of-claim wording, leaving cost centres to the administrator, FX, the post-trip report and UKRI acknowledgement. Each rule cites the guidance section or the administrator instruction it comes from.
- `assets/AI4Health_conference_cost_estimate_template.xlsx`: the CDT's blank cost-estimate sheet for approval requests.
- `scripts/build_claim.py`: fills Imperial's E1 form without breaking it, and packs receipts as `Surname_YYYYMMDD_Vendor_12.34EUR.jpg` in one ZIP.
- `SKILL.md`: the agent's procedure. It finds the trip's mail, reconciles, applies the rules, builds, checks in Excel, drafts, and sends only after you approve.

The guidance is reissued each year. If the CDT sends a newer edition, or the administrator tells you something different, that wins over the bundled file. PRs updating `references/ai4health-guidance.md` are welcome.

## Install

Claude Code:

```sh
git clone https://github.com/VishalJ99/cdt-expense-skill ~/.claude/skills/cdt-expense
```

Other agents: clone it into that agent's skills directory (e.g. `~/.codex/skills/cdt-expense`), or use `npx skills add VishalJ99/cdt-expense-skill`.

Then ask your agent something like *"claim my MICCAI expenses from the receipts in my mail"* or *"help me request approval for ICML"*.

## Requirements

- Python 3.9+ (standard library only)
- [poppler](https://poppler.freedesktop.org/) for `pdftoppm`: `brew install poppler` or `apt install poppler-utils`
- [Surface CLI](https://github.com/VishalJ99/surface-cli) to read and send Outlook or Gmail mail: `npm install -g surface-cli`. Any mail CLI with search, attachment download and send-with-attachments will also work.
- Microsoft Excel, for the Save As step and the final check
- **The E1 form is not bundled.** Use the `.xlsx` the CDT administrator last sent you. Otherwise download Imperial's [E1 for non-employees](https://www.imperial.ac.uk/media/imperial-college/administration-and-support-services/finance/public/E1-Non-employee-expense-form.xls) from the [Finance expenses page](https://www.imperial.ac.uk/finance/financial-services/expenses/), open it in Excel and *Save As .xlsx*.
- The administrator's email address, entered once at setup. The repo contains no staff names or addresses.

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
