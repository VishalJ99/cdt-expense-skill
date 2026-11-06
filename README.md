# cdt-expense-skill

Your agent handles AI4Health CDT conference approvals and expense claims end to end: it follows the CDT's rules, fills the E1 form, builds the receipts ZIP, drafts the email to Britta Ross (AI4Health CDT administrator, b.ross@imperial.ac.uk), and sends it once you approve.

## Install

```sh
git clone https://github.com/VishalJ99/cdt-expense-skill ~/.claude/skills/cdt-expense
```

## First run

Your agent asks once for your surname and the personal details the E1 needs: name, CID, date of birth, address and bank details.

These are saved to `~/.config/cdt-expense/me.json` on your machine only. The agent also:

- downloads the E1 for you to save once as `.xlsx` in Excel;
- installs the [Surface](https://github.com/VishalJ99/surface-cli) mail CLI and signs you in to your Imperial mail through the normal Imperial login in Chrome.

You need Python 3, Node, Microsoft Excel and poppler (`brew install poppler`).

## Use

Ask your agent, for example:

- "Claim my MICCAI 2026 expenses"
- "Request approval for ICML"
- "Can I claim a taxi from the airport?"

The rules it follows are in [`references/ai4health-guidance.md`](references/ai4health-guidance.md) (2025-26 guidance).
