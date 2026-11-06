---
name: cdt-expense
description: Use when claiming CDT or doctoral conference/travel expenses on the Imperial E1 Excel form; finds the mail thread, builds the workbook and receipts ZIP, drafts and sends.
---

# CDT expense claim

End to end: find the thread, collect receipts, build the E1 workbook and receipts ZIP, draft, verify, send. One claim = one workbook + one receipts ZIP + one email.

## Setup (once)

Record these in your agent memory or a local note:

- `SURNAME`, `PROGRAMME` (e.g. `AI4Health`), `ADMIN_EMAIL` (your CDT administrator), `ACCOUNT` (Surface account name for your university mail).
- **Mail CLI**: [Surface](https://github.com/VishalJ99/surface-cli) (`npm install -g surface-cli`). Any CLI that can search, read, download attachments and send with attachments works; adapt the commands below.
- **Template**: the original E1 `.xlsx` from your administrator, or from Imperial's Finance forms page. Prefer the latest copy your administrator sent you; never a re-saved or regenerated one.
- `python3`, `pdftoppm` (poppler). Microsoft Excel for the final check.

`scripts/` and `layouts/` below are relative to this skill's directory.

## Steps

1. **Read the rules.** Find your CDT's written claims guidance and the latest thread with your administrator:
   `surface mail search --account ACCOUNT --text "expenses" --limit 100`, then `surface mail read <message_ref>` / `surface mail thread get <thread_ref>`.
   The latest explicit instruction from the administrator overrides earlier precedent. Imperial CDT defaults, unless told otherwise:
   - E1 in its original Excel format, never PDF; one claim per email, with a ZIP of exactly that claim's receipts.
   - Itemised receipts only (date, items, vendor); card statements and payment confirmations are not receipts.
   - Receipts as JPG, named surname-first. Translate foreign-language receipts.
   - Purpose of claim: `SURNAME- PROGRAMME CDT-Conference expenses`, or the wording your guidance gives.
   - No alcohol, no tips; justify any taxi. Check your programme's per-trip cap.
   - Filenames: letters, digits and underscores only, no dashes.
   *Check:* you can quote the instruction behind each of these for your CDT.

2. **Reconcile before filling.** List what was already claimed in earlier threads; a "submitted for payment" reply means that expense is done. A travel-approval cost estimate is not a claim. Read every page of each receipt PDF, since one file can hold several receipts.
   *Check:* every expense appears once across all claims.

3. **Collect receipts and amounts.** Download attachments with `surface attachment download <message_ref> <attachment_id>`. For each expense record date, vendor, currency, amount and what was bought. Claim the subtotal plus unavoidable fees, never the tip. For foreign currency use the actual GBP card debit if you have it; otherwise use a public reference rate (e.g. ECB `EXR.D.GBP.EUR.SP00.A`) for the transaction date. Record the source and never invent a rate. If a fact is missing, gather everything else first, then ask once for the whole list.
   *Check:* every row has a receipt file, and every foreign amount has a GBP figure with a source.

4. **Write `claim.json`** (copy `examples/claim.example.json`). Rows with `"section": "travel"` go in Travel (rows 13–18); all others go in Subsistence/Hotels/Others (rows 41–47). Set `type` to the form's Expense Type dropdown value exactly (e.g. `HOTEL / SUBSISTENCE`, `CONFERENCE FEES`); conference registration also goes in this section. If a section overflows, split into a second claim instead of adding rows. Add `"declaration": {"name", "date"}` only if the claimant has authorised signing.
   *Check:* `python3 scripts/build_claim.py --template E1.xlsx --claim claim.json --dry-run` shows the right cells, receipt names and total.

5. **Build.** `python3 scripts/build_claim.py --template E1.xlsx --claim claim.json --out-dir out`
   This edits the template's XML in place. It keeps styles, validation and namespaces, clears the old rows, cost-centre and approval cells, and restores the analysis-code VLOOKUP formulas (the administrator adds the cost centre; hard-coded codes have caused rejections). It drops every cached total and forces recalculation. Receipts become `SURNAME_YYYYMMDD_Vendor_12.34CCY.jpg`: PDFs render page by page at 150 dpi, translations go in `_translation.txt`, and macOS junk is excluded. Don't re-save the workbook with openpyxl or similar, which strips features Excel needs.
   *Check:* the printed ZIP listing has one JPG per receipt page, and `summary.txt` shows the total you expect.

6. **Verify in Excel.** Open the workbook in Microsoft Excel. Confirm there is no repair prompt, the total matches `summary.txt`, the Analysis Code cells show numbers, and no previous claim's values remain. LibreOffice and Quick Look previews don't count: LibreOffice shows `#N/A` for the analysis codes even when Excel resolves them.
   *Check:* report what Excel showed, labelled as an Excel check.

7. **Draft.** Write the email in your own voice: purpose, total, one line per foreign amount with its rate and source, and any justification (e.g. taxi). Mention only attachments that are actually included. Save as a draft for review:
   `surface mail reply <message_ref> --body "$(cat body.txt)" --attach out/<workbook>.xlsx --attach out/<receipts>.zip --draft`
   (or `surface mail send --account ACCOUNT --to ADMIN_EMAIL --subject "..." ... --draft` for a new thread).
   *Check:* read the draft back. The attachment names and body should match what you planned.

8. **Send only on explicit approval** from the claimant. Approving a self-test is not approval to send to the administrator. Run the same command without `--draft`. After any change, rebuild the workbook and ZIP together and redraft; never patch one file alone.
   *Check:* find the message in Sent (`surface mail search --account ACCOUNT --mailbox "<sent folder>" --subject "..."`), download its attachments, and confirm their SHA-256 matches `out/`.

## Other forms

`layouts/imperial-e1.json` maps the Imperial E1 v1.2 cells. For a different version, copy it and change the cell references; the script refuses to write a cell into a row the template lacks.
