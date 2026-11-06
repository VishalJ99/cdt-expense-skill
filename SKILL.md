---
name: cdt-expense
description: AI4Health CDT conference approval requests and expense claims; applies the CDT rules, builds the E1 workbook, receipts ZIP and email, and sends after approval.
---

# AI4Health CDT expenses

The rules are in `references/ai4health-guidance.md`; rule numbers below refer to it. Apply them as written. Only a newer guidance edition or a later instruction from the administrator in the student's mail overrides them; name the override in your reply.

A claim is one E1 workbook, one ZIP of that workbook's receipts, and one email to **Britta Ross, AI4Health CDT administrator, `b.ross@imperial.ac.uk`**. She is the recipient for every approval request and claim. `SKILL` below means this skill's directory.

## First run (once)

Personal files live in `~/.config/cdt-expense/`.

1. **`me.json`.** If it is missing, ask the student in one message for their surname and the E1 personal details:
   - title and first name;
   - CID and date of birth;
   - home address (up to 4 lines) and postcode;
   - bank name, branch, sort code and account number.
   Write it in the shape of `SKILL/examples/me.example.json` and `chmod 600` it. Never copy it elsewhere or send it anywhere.
2. **`E1.xlsx`.** Download Imperial's [E1 for non-employees](https://www.imperial.ac.uk/media/imperial-college/administration-and-support-services/finance/public/E1-Non-employee-expense-form.xls) into that folder. Ask the student to open it in Microsoft Excel and *Save As → Excel Workbook* named `E1.xlsx` in the same folder. Excel must do the conversion.
3. **Mail.** Run `surface auth status uni`. If it is not authenticated:
   ```sh
   npm install -g surface-cli
   surface skill install claude-code      # codex | all; on OpenClaw: openclaw skills install surface-cli
   surface account add uni --provider outlook --email <student's Imperial email>
   surface auth login uni                 # Chrome opens; the student signs in with Imperial SSO
   ```
   Anything else about Surface: follow the Surface skill.

*Check:* both files exist, and `surface mail search --account uni --text "expenses" --limit 1` returns a result.

## Approval (before booking)

Copy `SKILL/assets/AI4Health_conference_cost_estimate_template.xlsx` and fill it in with the student (rules 1–6, 9–13). Draft the request to Britta with `--draft`, as in claim step 7. Booking waits for approval.

## Claim

1. **Collect.** `surface mail search --account uni --text "<conference>" --limit 50`, then `surface mail read <message_ref>` and `surface mail thread get <thread_ref>`. Find the approval, earlier claims for this trip, and any later administrator instruction.
   *Check:* you know what was approved and what's already claimed.
2. **Reconcile.** A "submitted for payment" reply means that item is done. An estimate is not a claim (rule 5). Read every page of each receipt PDF, since one file can hold several receipts.
   *Check:* every expense appears once across all claims, and the trip total is within £3,000 (rule 8).
3. **Receipts and amounts.** Download with `surface attachment download <message_ref> <attachment_id>`. For each expense, record date, vendor, currency, amount and items.
   - Itemised receipts only (rule 23). Leave out alcohol and tips (rules 14–15). Keep to eligible dates and nights (rule 9). Write a taxi justification (rule 10). Translate foreign receipts (rule 24).
   - Foreign amounts: use the GBP card debit if the student has it, otherwise the ECB reference rate for the transaction date (`EXR.D.GBP.EUR.SP00.A` for EUR). Never invent a rate (rule 22).
   - Ask for any missing facts once, as a single list.
   *Check:* every row has an itemised receipt file and a GBP amount with a source.
4. **`claim.json`.** Put it in a new folder `~/cdt-expenses/<Conference><Year>/`, in the shape of `SKILL/examples/claim.example.json`:
   - `claim`: e.g. `MICCAI 2026`. It feeds the filenames and the email subject.
   - `fx_source`.
   - `rows`, each with: `date`, `section` (`travel` for Travel; omit for Subsistence/Hotels/Others, which includes registration, rule 20), `mode`/`from`/`to` for travel, and `type` as the exact Expense Type dropdown value (`CONFERENCE FEES`, `HOTEL / SUBSISTENCE`, ...). Then `description`, `vendor`, `amount`, `currency`, `gbp`, `receipt`, and where needed `translation` and `note` (justification or exclusions).
   - More rows than the form holds means a second claim, not extra rows (rules 18, 26).
   - Add `"sign": true` only when the student says to sign the declaration.
   *Check:* `python3 SKILL/scripts/build_claim.py claim.json --dry-run` shows the right cells, receipt names and email.
5. **Build.** `python3 SKILL/scripts/build_claim.py claim.json` writes `out/`:
   - `<Surname>_E1_<claim>.xlsx`: the student's details and purpose (rule 19); cost centre and approvals cleared; analysis-code formulas intact (rule 21).
   - `<Surname>_<claim>_receipts.zip`: one `<Surname>_YYYYMMDD_Vendor_12.34EUR.jpg` per receipt page, plus translations (rules 24–25, 27).
   - `email.txt`.
   Never re-save the workbook with any other tool.
   *Check:* the printed ZIP listing has one JPG per receipt page.
6. **Excel.** Open the workbook in Microsoft Excel. Confirm there is no repair prompt, the total matches `email.txt`, the Analysis Code cells show numbers, and the details are right. LibreOffice shows `#N/A` there, so only Excel counts.
   *Check:* report what Excel showed.
7. **Draft** to Britta, using the subject and body from `email.txt`; add a taxi justification or other context in the student's words if needed.
   ```sh
   surface mail send --account uni --to b.ross@imperial.ac.uk --subject "<subject>" --body "<body>" \
     --attach out/<workbook>.xlsx --attach out/<receipts>.zip --draft
   ```
   *Check:* read the draft back; the attachments and body match `out/`. Surface 0.4.6 doesn't verify read-back itself.
8. **Send** only when the student explicitly approves, using the same command without `--draft`. After any change, rebuild and redraft; never patch one file alone.
   *Check:* find it in Sent with `surface mail search --account uni --mailbox "Sent Items" --subject "<subject>"`, download the attachments, and confirm their SHA-256 matches `out/`.
9. **After the trip.** Remind the student to email the CDT a short conference report (rule 28) and to cite the UKRI grant (rule 29).
