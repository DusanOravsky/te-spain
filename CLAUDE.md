# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

A single static `index.html` the user opens locally (double-click → browser via `file://`). It builds the monthly fuel-price liquidation in three steps the user does once a month:

1. Pick the vendor's `Flota Viva Covestro <date>.xlsx` (from Ayvens).
2. Pick the running `Liquidaciones de combustible …xlsx`.
3. Pick `preciosEESS_es.xls` downloaded from `geoportalgasolineras.es`.

The page does everything in-browser via SheetJS (CDN):

- Averages **Precio gasolina 95 E5** and **Precio gasóleo A** across every station that reports a price.
- Replaces the `flota` sheet of Liquidaciones with the vendor's fleet rows (columns B..P, header on row 10).
- Computes consumption averages directly from the vendor file (replacing what `consumo` pivot used to provide via `GETPIVOTDATA`):
  - Diesel avg → rows where `Tipo Combustible` contains "Diésel"
  - Gasolina avg → rows where `Tipo Combustible` contains "Gasolina"
- Finds the email-snippet header row (`Fecha del cálculo` in column A, scanned bottom-up) on the `Liquidaciones de combustible` sheet and overwrites the row immediately below it with: A=date, H=diesel price, I=diesel consumption, J=`=H`, K=`=I*J/100`, L=gasolina price, M=gasolina consumption, N=`=L`, O=`=M*N/100`.
- Outputs (a) a downloadable `<original>_updated.xlsx` and (b) a copy-paste email draft (greeting → header+row tab-separated → tarifas summary → "Thanks").

## Why this shape

Every other shape was tried and failed against Covestro constraints:

- **Browser fetches the .xls directly** — no CORS on `geoportalgasolineras.es`, browsers block it.
- **Local Python server / .exe** — Python isn't installed on the user's PC and .exes are blocked.
- **GitHub Actions + Pages** — works mechanically but `cov-secure-policy` requires PR + CodeQL on `main`, blocking the bot from committing `data.json`. PR auto-merge would need a separate bot account, which contradicts the policy's intent.
- **Excel + Power Query** — works (and `fuel_prices.xlsx` is the proof), but the user wanted a web app for the broader workflow.

The "user picks the .xls themselves" version is the only one that survives all constraints.

## Why a few things look the way they do

- **Cell-address reads, not `sheet_to_json`.** SheetJS's `sheet_to_json({header: 1})` strips leading empty rows, so a header on row 10 becomes index 9 only sometimes. Reading by `XLSX.utils.encode_cell({r, c})` is unaffected.
- **Pivot bypass.** The `consumo` sheet's pivot table cells (`E8`, `E17`) feed `I` and `M` of the new month row via `GETPIVOTDATA`. SheetJS cannot rebuild Excel pivots — it can only write values. So the page recomputes the same averages from the vendor file and writes them as plain numbers. Pivot remains in the workbook for backward compatibility but is no longer load-bearing.
- **The email "header row" lives mid-sheet.** Layout is: history block top → snippet header (`Fecha del cálculo` etc.) → exactly one current-month data row. The user copies header+row into Outlook each month. We locate the header by scanning column A from the bottom up and overwrite the row below it.
- **No `data.json`, no GitHub Actions, no hosting.** Everything was deleted; previous incarnations are in git history.

## Files

- `index.html` — the entire app. SheetJS loaded from `cdn.jsdelivr.net`.
- `fuel_prices.xlsx` — standalone Power Query workbook (the user already wired this up). Independent of `index.html`; kept because it's useful as a quick price check.
- `docs/` — local-only test fixtures (vendor file, current Liquidaciones, sample preciosEESS .xls). **Gitignored** — contains internal Covestro data.

## How to test changes

There's no harness in the repo. To validate end-to-end, mirror the browser logic in Node + SheetJS and run against `docs/`:

```bash
# from /tmp:
npm install xlsx@0.18.5
node test_build.mjs   # see /tmp/test_build.mjs in the project author's session for the template
```

The expected output (against the snapshot in `docs/`):
- Diesel price avg ≈ 1.7345 €/L, gasolina ≈ 1.5248 €/L
- Diesel consumption ≈ 4.756, gasolina ≈ 5.788 — these match `consumo!E8` / `E17` exactly
- 45 fleet rows pasted into `flota`
- Snippet header row 144, new month row written at row 145

If those numbers drift, either the vendor file format changed or the price-file format changed — check the column-name constants near the top of `index.html`.

## Out of scope / known gaps

- **No history block update.** The previous month row (`143`) stays untouched. If the workflow eventually needs to "promote" the current row into history before adding the new one, that's a small additive change in the section that finds the snippet header.
- **CDN dependency.** SheetJS is loaded from jsdelivr at runtime. If the user's network blocks CDNs, vendoring `xlsx.full.min.js` next to `index.html` and pointing the `<script src>` at it would fix that.
- **Date format.** Dates from the vendor file are written to `flota` as Excel serial numbers — Excel renders them correctly with its date format applied (which it already has). If the column ever shows up as a number, apply a date format to those cells in Excel.
