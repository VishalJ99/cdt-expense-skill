---
name: cdt-expense
description: Use when an Imperial AI4Health CDT student requests conference approval or claims expenses; applies the CDT's rules, builds the E1 workbook and receipts ZIP, drafts and sends.
---

# AI4Health CDT expense claim

For students of the UKRI AI4Health CDTs at Imperial (CDT in AI for Healthcare; AI CDT in Digital Healthcare). The CDT's rules are already in `references/ai4health-guidance.md`; read it at the start of every run and apply it by default. Rule numbers below refer to that file. Rules marked [P] are practice from accepted claims, not written rules. The administrator's latest email to the student, or a newer guidance edition, overrides the file; say so when it does.

One claim = one E1 workbook + one ZIP of that workbook's receipts + one email.

## Setup (once)

Record in agent memory or a local note:

- `SURNAME`, and `ADMIN_EMAIL`, the address of the AI4Health CDT administrator / programme manager (from the guidance email or the CDT team).
- `ACCOUNT`: the Surface account name for the student's Imperial mail (`uni` if you follow Step 0).
- **E1 template.** Use the `.xlsx` the administrator last sent you if there is one. Otherwise download Imperial's [E1 for non-employees](https://www.imperial.ac.uk/media/imperial-college/administration-and-support-services/finance/public/E1-Non-employee-expense-form.xls) (listed on https://www.imperial.ac.uk/finance/financial-services/expenses/). It is `.xls`, so open it in Microsoft Excel and *Save As → Excel Workbook (.xlsx)*. Do not convert it with LibreOffice or Python; the form must stay original.
- `python3`, `pdftoppm` (poppler) and Microsoft Excel.

Paths below are relative to this skill's directory.

## Step 0: mail setup (once)

This skill reads and sends mail with [Surface](https://github.com/VishalJ99/surface-cli).
1. Run `surface auth status uni` (or the student's `ACCOUNT`). If it reports the account as authenticated, skip to the claim steps.
2. Otherwise:
   ```sh
   npm install -g surface-cli
   surface skill install claude-code      # or codex | all; on OpenClaw: openclaw skills install surface-cli
   surface account add uni --provider outlook --email <you>@imperial.ac.uk
   surface auth login uni                 # opens Chrome: the student completes normal Imperial SSO
   ```
   Outlook uses the browser SSO session, so no Azure app registration or app password is needed. Wait for the student to finish signing in, then rerun `surface auth status uni`.
3. For anything else (sender identity, `surface auth check`, re-login, Gmail/IMAP), follow the installed Surface skill or the Surface README's Setup section; don't improvise.

*Check:* `surface auth status uni` shows it authenticated, and `surface mail search --account uni --text "expenses" --limit 1` returns a result.

## Before the trip (approval)

If the student is planning a trip rather than claiming, fill in a copy of `assets/AI4Health_conference_cost_estimate_template.xlsx` with them, following rules 1–6 and 9–13, and draft the approval request to `ADMIN_EMAIL`. Stop there; booking needs CDT approval.

## Claim steps

1. **Collect the trip's mail.** `surface mail search --account ACCOUNT --text "<conference> expenses" --limit 50`, then `surface mail read <message_ref>` / `surface mail thread get <thread_ref>`. Look for the approval, earlier claims for this trip, and any newer instruction from the administrator.
   *Check:* you know what was approved and the cap (rule 8), and any instruction that overrides the reference file.

2. **Reconcile.** List what was already claimed; a "submitted for payment" reply means that item is done. An approved estimate is not a claim (rule 5). Read every page of each receipt PDF, since one file can hold several receipts.
   *Check:* every expense appears once across all claims, and the trip total is within £3,000.

3. **Collect receipts and amounts.** `surface attachment download <message_ref> <attachment_id>`. For each expense record date, vendor, currency, amount and items bought. Then apply the rules:
   - Itemised receipts only (rule 23). Remove alcohol and tips (rules 14–15). Keep to the eligible dates and nights (rule 9). Note a taxi justification (rule 10). Translate foreign receipts (rule 24).
   - Foreign amounts: use the actual GBP card debit if you have it, otherwise a public reference rate for the transaction date (e.g. ECB `EXR.D.GBP.EUR.SP00.A`). Record the source, and never invent a rate (rule 22).
   - If something is missing, gather everything else first, then ask once for the whole list.
   *Check:* every row has an itemised receipt file, and every foreign amount has a GBP figure with a source.

4. **Write `claim.json`** (copy `examples/claim.example.json`):
   - `purpose`: `SURNAME- AI4Health CDT-Conference expenses` (rule 19).
   - Rows with `"section": "travel"` go in Travel; all others go in Subsistence / Hotels / Others, including registration (rule 20).
   - `type`: the Expense Type dropdown value exactly, e.g. `CONFERENCE FEES`, `HOTEL / SUBSISTENCE`.
   - If a section overflows, split into a second claim instead of adding rows (rules 18, 26).
   - Add `"declaration": {"name", "date"}` only if the student has authorised signing.
   *Check:* `python3 scripts/build_claim.py --template E1.xlsx --claim claim.json --dry-run` shows the right cells, receipt names and total.

5. **Build.** `python3 scripts/build_claim.py --template E1.xlsx --claim claim.json --out-dir out`
   The script:
   - edits the template's XML in place, keeping styles, validation and namespaces;
   - clears old rows and the cost-centre and approval cells, and restores the analysis-code formulas (rule 21);
   - drops cached totals and forces recalculation;
   - names files with letters, digits and underscores only (rule 27): `SURNAME_E1_<claim>.xlsx` plus `SURNAME_<claim>_receipts.zip`. The ZIP holds one `SURNAME_YYYYMMDD_Vendor_12.34CCY.jpg` per receipt page (PDFs rendered at 150 dpi) and a `_translation.txt` per foreign receipt (rules 24–25).
   Never re-save the workbook with openpyxl or LibreOffice.
   *Check:* the printed ZIP listing has one JPG per receipt page, and `summary.txt` shows the expected total.

6. **Verify in Excel.** Open the workbook in Microsoft Excel. Confirm there is no repair prompt, the total matches `summary.txt`, the Analysis Code cells show numbers, and no previous claim's values remain. LibreOffice shows `#N/A` in that column even when Excel resolves it, so only Excel counts.
   *Check:* report what Excel showed, labelled as an Excel check.

7. **Draft.** Write the email in the student's voice: purpose, total, one line per foreign amount with rate and source, any taxi justification, and what was excluded (alcohol, tips). Mention only attachments that are actually included. Save it as a draft:
   `surface mail reply <message_ref> --body "$(cat body.txt)" --attach out/<workbook>.xlsx --attach out/<receipts>.zip --draft`
   (or `surface mail send --account ACCOUNT --to ADMIN_EMAIL --subject "<Conference> <Year> expense claim" ... --draft`).
   *Check:* read the draft back. The attachment names and body should match the plan. Do this yourself even if Surface verifies on send: the npm release (0.4.6) doesn't yet check that the sent message reads back exactly.

8. **Send only on the student's explicit approval.** Approving a self-test is not approval to send to the administrator. Run the same command without `--draft`. After any change, rebuild the workbook and ZIP together and redraft; never patch one file alone.
   *Check:* find the message in Sent (`surface mail search --account ACCOUNT --mailbox "<sent folder>" --subject "..."`), download its attachments, and confirm their SHA-256 matches `out/`.

9. **After the trip:** remind the student to email the CDT a short conference report (rule 28) and to cite the UKRI grant in outputs (rule 29).
