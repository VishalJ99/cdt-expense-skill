---
name: cdt-expense
description: AI4Health CDT conference approval requests and expense claims; applies the CDT rules, builds the E1 workbook, receipts ZIP and email, and sends after approval.
---

# AI4Health CDT expenses

Rules are in `references/ai4health-guidance.md`; rule numbers below refer to it. A later instruction from Britta in the student's mail overrides them; say so in the reply.

Every approval request and claim goes to **Britta Ross, `b.ross@imperial.ac.uk`**. A claim is one E1 workbook, one ZIP of its receipts and one email. `SKILL` is this skill's directory.

**Never send mail without the student's explicit OK.**

## First run

Personal files live in `~/.config/cdt-expense/`.

1. **`me.json`.** If missing, ask the student in one message for their title, first name, surname, CID, date of birth, home address (up to 4 lines), postcode, bank name, branch, sort code and account number. Write it in the shape of `SKILL/examples/me.example.json` and `chmod 600` it. Never copy or send it anywhere.
2. **`E1.xlsx`.** Download Imperial's [E1 for non-employees](https://www.imperial.ac.uk/media/imperial-college/administration-and-support-services/finance/public/E1-Non-employee-expense-form.xls) into that folder. Ask the student to open it in Microsoft Excel and *Save As → Excel Workbook* as `E1.xlsx` there.
3. **Mail.** If `surface auth status uni` is not authenticated:
   ```sh
   npm install -g surface-cli
   surface skill install claude-code
   surface account add uni --provider outlook --email <student's Imperial email>
   surface auth login uni    # Chrome opens for Imperial sign-in
   ```

## Approval (before booking)

Copy `SKILL/assets/AI4Health_conference_cost_estimate_template.xlsx` and fill it in with the student (rules 1–6, 9–13). Show the student the request email and spreadsheet; on their OK, send it to Britta as in claim step 6. Nothing is booked before approval.

## Claim

1. **Collect.** `surface mail search --account uni --text "<conference>" --limit 50`, then `surface mail read <message_ref>` / `surface mail thread get <thread_ref>`. Find the approval, earlier claims for this trip and any later instruction from Britta. A "submitted for payment" reply means those items are done; an approved estimate is not a claim (rule 5). Each expense is claimed once, and the trip total stays within £3,000 (rule 8).
2. **Receipts.** `surface attachment download <message_ref> <attachment_id>`. Read every page, since one PDF can hold several receipts. For each expense record date, vendor, currency, amount and items:
   - itemised receipts only (rule 23); no alcohol or tips (rules 14–15); eligible dates and nights only (rule 9); justify any taxi (rule 10); translate foreign receipts (rule 24);
   - GBP amount: the card debit if the student has it, else the ECB reference rate for the transaction date (`EXR.D.GBP.EUR.SP00.A` for EUR). Never invent a rate (rule 22).

   Ask for missing facts once, as one list.
3. **`claim.json`** in a new folder `~/cdt-expenses/<Conference><Year>/`, shaped like `SKILL/examples/claim.example.json`:
   - `claim`: e.g. `MICCAI 2026`; it names the files and the email subject. `fx_source` if any amount is foreign.
   - `rows`: `date`, `section: "travel"` with `mode`/`from`/`to` for travel (omit `section` for everything else, including registration, rule 20), `type` as the exact Expense Type dropdown value (`CONFERENCE FEES`, `HOTEL / SUBSISTENCE`, ...), `description`, `vendor`, `amount`, `currency`, `gbp`, `receipt`, and where needed `translation` and `note` (justification or exclusions).
   - More rows than the form holds means a second claim (rules 18, 26).
   - `"sign": true` only if the student says to sign the declaration.
4. **Build.** `python3 SKILL/scripts/build_claim.py claim.json` writes `out/`: `<Surname>_E1_<claim>.xlsx`, `<Surname>_<claim>_receipts.zip` (one JPG per receipt page, plus translations) and `email.txt`. Don't re-save the workbook with any tool but Excel. After any change, rebuild everything.
5. **Excel.** The student opens the workbook in Microsoft Excel: no repair prompt, total matches `email.txt`, Analysis Code cells show numbers, details right. (LibreOffice shows `#N/A` there.)
6. **Send.** Show the student `email.txt` (add any context in their words) and the ZIP listing. Only on their explicit OK:
   ```sh
   surface mail send --account uni --to b.ross@imperial.ac.uk --subject "<subject>" --body "<body>" \
     --attach out/<workbook>.xlsx --attach out/<receipts>.zip
   ```
   Then confirm it is in Sent Items: `surface mail search --account uni --mailbox "Sent Items" --subject "<subject>"`.
7. **After the trip.** Remind the student to email the CDT a short conference report (rule 28) and to cite the UKRI grant (rule 29).
