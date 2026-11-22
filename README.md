# Loan Tape Review

Upload a Setpoint-style loan tape workbook (`.xlsx`) and get a **shareable report**: grades, guideline checks, missing documents, data observations, a draft client email, and a **source index that cites the exact cell behind every finding** (for example `Report!HM3`, `Guidelines!D11`, `Waivers!A2`).

- 100% client-side: the file is read in your browser and **never uploaded**.
- Free to host on **GitHub Pages**. No server, no build step, no API keys.

## Use it

1. Open your Pages site (see "Publish" below).
2. Drop in the workbook. Optionally set the report title, your name, the client's name and the grade the client expects.
3. Share the result: **Download report (HTML)** (a single self-contained file you can email or attach), **Print / Save as PDF**, **Copy client email**, **Sources (CSV)**, or **Full results (JSON)**.

Try **Load demo data** first to see a report without uploading anything.

## Publish for free on GitHub Pages

1. Create a new repository on GitHub and push these files to the `main` branch.
2. In the repo go to **Settings > Pages** and set **Source** to **GitHub Actions**.
3. Push (or run the **Deploy to GitHub Pages** workflow). Your app appears at `https://<your-username>.github.io/<repo-name>/`.

Run locally instead: `npm start`, then open http://localhost:8080 (any static server works; opening `index.html` directly also works in most browsers).

## What it checks

The app recomputes everything from the raw tabs, using the same rules as the template Summary tab:

| Area | Rule |
|---|---|
| Bucket | From **Loan Type** on the Report tab: blank = PENDING, starts with "C" = Construction, 5+ units = Commercial, else SFR 1-4 |
| Initial Cost | As-Is Value if the purchase is over 730 days old (DAYS360) or the price is 0; otherwise Purchase Price + Cost Incurred |
| Cost Basis | Initial Cost + Repair Budget + Debt Service - Cost Incurred to Date |
| Initial LTC | (Note - Holdback) / Initial Cost |
| Total LTC | Note / Cost Basis |
| LTAIV | (Note - Holdback - Interest Reserve) / As-Is Value |
| LTARV | Note / As-Repaired Value |
| Term | ROUND(DAYS360(First Payment, Maturity) / 30 + 1), with Overrides tab applied |
| Guideline checks | FICO, Loan Amount (Construction has its own max), Term, Initial LTC, Total LTC, LTAIV, LTARV; the Waivers tab is checked first and shows **Waived** |
| Missing docs | Case-insensitive count of "missing" in the Exceptions column |
| Credit grade | 0 issues: A (B with 1-2 waivers, C with 3+); 1 issue: B (C with a waiver); 2: C; 3+: D |
| Valuation grade | Deficiency score (0 = A, 1-2 = B, 3+ = C); BPO / hybrid / desktop / other cap at B |
| Overall grade | The lower of Credit and Valuation |

It also raises **observations**, each with source cells: blank Loan Type, client-vs-tape mismatches, non-zero `... Variance` flags (and whether the underlying client and tape values actually agree), note loan term vs calculated term, flood policy flagged missing while not in a flood zone, hazard policy expiry, FICO on the Report vs the Guarantors tab, an aged credit report flagged not aged (assumed after 120 days; edit `AGED_CREDIT_DAYS` in `js/engine.js`), and any difference from the grade cached in the workbook's own Summary tab.

## Expected workbook layout

- **Report**: headers in row 2, one row per loan from row 3, `Property Code` in column A. Columns are found by header name (e.g. `Note Amount ($)`, `Client Note Amount ($)`).
- **Guidelines**: labels in column A (`Max Term`, `FICO`, `Max Loan Amount`, `Max Loan Amount (Construction)`, `LTAIV`, `LTARV`, `Initial LTC`, `Total LTC`) with bucket columns B-D under a row headed `SFR 1-4 | Commercial | Construction`.
- **Waivers**: `Property Code | Guideline | Waived At | Comment`.
- **Overrides**: `Property Code | Override Field | Override Value | Comment` (field names match the Summary headers, e.g. `Term`).
- **Guarantors** (optional) and **Summary** (optional, for the cross-check).

If a required tab is missing the app says so; missing Report columns are listed as warnings in the report.

## Test

```
node test/run.js                       # built-in tests with fictional demo data
node test/run.js path/to/workbook.xlsx # also analyse a real workbook and print the grades
```

## Notes and limits

- **Do not commit real loan data.** `.gitignore` excludes spreadsheets and generated reports. Pages hosts only the app, never your uploads.
- Dates are read from cached values; date formulas like `=VALUE("2026-04-01")` with no cached value are parsed from the formula text.
- Template quirks are reproduced rather than "fixed", except where noted in the report. Review grades before sending anything to a client.
- Spreadsheet parsing uses SheetJS Community Edition 0.18.5 (vendored in `vendor/`, Apache-2.0). It runs only on files you choose yourself, locally.
