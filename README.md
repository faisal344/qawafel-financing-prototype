# Qawafel financing readiness prototype

Task 3 of the Qawafel Graduate Software Engineer assessment.

An SME wholesaler uploads 12 months of invoices and bank statements. This prototype turns them into what a credit analyst and a bank actually need:

1. **A readiness verdict and an indicative limit.** The verdict is *Ready to send*, *Send after fixes* or *Not ready*. It comes from an 8-line scorecard (cash-verified sales, stability, growth, collections, concentration, liquidity, debt service, data quality). Each line has a traffic light and the reason behind it.
2. **Evidence.** It shows monthly sales against the customer money that actually reached the bank, the daily bank balance, receivables aging and a per-buyer table.
3. **A borrowing base you can move.** There are sliders for advance rate, overdue cut-off, single-buyer cap, late-payer exclusion and a cash-flow cap, so you can see how the bank's policy changes the limit.
4. **Data checks.** It catches duplicates, VAT that doesn't add up, the same buyer under two names, bank credits that aren't sales, and invoices marked paid that never reached the bank. Each check says what to do about it.
5. **The bank package.** A versioned JSON (`qawafel.financing-package/v1`) that matches the spec in the written answers. This is what would be sent to the bank partner.

The point of the design: **an invoice file is a claim and the bank statement is the evidence.** The most useful thing the tool does is check one against the other. In the sample data, one buyer has SAR 568k of invoices marked "paid" and none of it ever reaches the bank. The tool blocks the package on that buyer until the SME explains it.

## Open it

- **Live:** see the link in the submission. It is a single page with nothing to install.
- **Locally:** open `index.html` in any browser. The sample data is inlined; charts and Excel upload load Chart.js and SheetJS from cdnjs, so you need to be online.
- **Your own files:** use *Upload invoices* (.xlsx/.csv) and *Upload bank statement* (.csv) in the page header. Column names are matched loosely (e.g. `Invoice No`, `Invoice Number`, `Customer`, `Buyer` all work). See `data/invoices.xlsx` for the expected shape.

## Sample data (all made up)

| File | What it is |
| --- | --- |
| `data/invoices.xlsx` | 2,465 sales invoices and 96 credit notes for *Rimal Al-Sharq Trading Co. (fictional)*, Oct 2025 – Sep 2026. This is what the SME would upload. |
| `data/bank_statement.pdf` | The same 12 months as monthly statements from a fictional bank, as the SME would upload them. |
| `data/bank_statement.csv` | The same transactions as CSV. This stands in for the output of the PDF parser, which is out of scope for the prototype. |
| `data/sample-data.js` | Both datasets as JSON, inlined into `index.html` by `scripts/build.py`. |

Regenerate with `python3 scripts/generate_data.py && python3 scripts/build.py` (needs `openpyxl` and `reportlab`). The seed is fixed, so you get the same files every time.

### How the data was built, and why it should look real

- **Company:** a Riyadh FMCG/food wholesaler. About SAR 16.5M net sales a year (~SAR 1.35M a month), 23 staff, one warehouse, 34 buyers. All names, CR, VAT numbers and IBANs are invented.
- **Buyers:** a long tail. One hypermarket chain is ~25% of sales, the top 5 are ~58%, and the rest are minimarkets, HoReCa and groceries. Terms are 0 (cash on delivery), 15, 30, 45 or 60 days, and bigger buyers get longer terms.
- **Seasonality:** a Ramadan 1447 stock-up (Jan–Feb 2026 at +12% / +38%), an Eid al-Adha bump (May), a summer dip (Jul–Aug) and back-to-school (Sep). Underlying growth is about 1.2% a month. No deliveries on Fridays.
- **Amounts:** order sizes are log-normal and scale with buyer size. VAT is 15% and amounts are in SAR.
- **Payment behaviour:** each buyer has its own average delay and spread. About 3% of invoices are short-paid (damaged goods deducted). Two HoReCa buyers pay late and leave 2–4% of invoices uncollected. Credit notes come to ~1.7% of sales.
- **Bank:** big buyers pay by SARIE transfer, often one transfer for several invoices, sometimes with a truncated payer name. Cash buyers' money is deposited at a branch twice a week. Debits cover suppliers (about 75% of net sales, plus VAT), payroll (WPS), GOSI, quarterly rent, quarterly VAT to ZATCA, fuel, utilities, bank charges and an **existing term loan of SAR 42.5k a month**.
- **Opening receivables:** invoices from the two months before the window were generated too. Their payments land in Oct–Nov 2025, but the invoices themselves are not in the upload, just like a real 12-month cut.

### Problems planted on purpose (the prototype should catch every one)

| Problem | Where |
| --- | --- |
| 5 duplicated invoice rows | re-exported rows |
| 4 rows where VAT ≠ 15% of net | typed by hand |
| Largest buyer under two names, same VAT no. | "Al-Waha Hypermarkets" / "Al Waha Hypermarket Co." Splitting the name hides that the buyer is really ~25% of sales |
| Invoices marked paid with no bank trace | Nakheel Corner Mini Mart, SAR 568k |
| Non-sales credits | own-account transfer SAR 300k, owner injection SAR 250k, VAT refund |
| Chronic late payer | Fanar Restaurants Group, ~35 days late |
| Unattributable transfers | short "IPS IN …" payer names |

## Assumptions

1. **Product:** invoice financing (an advance against open receivables). I chose it because Qawafel sits on B2B trade, and it's the product where invoice data matters most. A term loan would weight cash flow and debt service more.
2. **Qawafel prepares, the bank decides.** The verdict is about whether the *package is ready to send*. The limit is labelled indicative.
3. **Scorecard thresholds and borrowing-base defaults are my own first guesses.** They are an 80% advance rate, a 60-days-past-due cut-off, a 25% single-buyer cap, excluding buyers more than 30 days late on average, and a cap of 1× monthly verified collections. All of them are levers to set with the bank's credit policy.
4. **The as-of date is the last bank transaction date (30 Sep 2026).** Invoices with a paid date after that count as unpaid.
5. **Matching bank credits to buyers** uses the first 12 letters of the buyer's name (payer names are often cut short). Production would use the buyer's IBAN, the invoice reference and amount combinations.
6. **Dates in uploads are read as `dd/mm/yyyy`** when ambiguous (Saudi convention). Hijri dates are not handled in the prototype.
7. **Bank statement PDFs are not parsed in the prototype.** The CSV stands in for the parser's output. In production I'd use a PDF table parser for digital statements, OCR with confidence scores for scans, and ideally open banking (SAMA framework) instead of PDFs.
8. Everything runs in the browser. Nothing is uploaded anywhere.

## What I'd build next

- A server-side ingestion API (see the written spec) with stored, versioned packages.
- PDF statement parsing with balance-chain validation (each opening balance = the previous closing balance).
- Match by IBAN and invoice reference, and check invoices against ZATCA e-invoicing (Fatoora) data.
- Unit tests around `analyse()`. It is a pure function of (invoices, bank, policy), so it is easy to test.

## Files

```
index.html               the app, with sample data inlined (open this)
app.html                 the app source, before data is inlined
scripts/generate_data.py builds the sample data
scripts/build.py         inlines data into index.html
data/                    sample invoices (xlsx), bank statement (pdf + csv), JSON
```
