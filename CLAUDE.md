# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

A single-page web app — `index.html` — that runs the monthly Spain fuel-price liquidation end-to-end. Branded as a Covestro internal tool. The user opens the file from disk, picks the vendor fleet workbook and the running Liquidaciones, gets back an updated `.xlsx` and a copy-paste-ready Outlook email draft.

Live at:
- **Covestro org**: `covestro/accounting-spain-gas` (private, SSO, `cov-secure-policy`)
- **Personal mirror**: `DusanOravsky/te-spain` — `main` matches origin; branch `docs-personal` adds the internal reference fixtures (vendor file, Liquidaciones snapshot, sample .xls). Never push docs to `origin`.

## Architecture (single file, no build, no backend)

```
   index.html ───┬───► api.minetur.gob.es (CORS-open JSON)   ── live prices
                 ├───► reads vendor .xlsx (drag/drop or pick) ── fleet rows
                 └───► reads Liquidaciones .xlsx               ── existing history

           ↓ in browser, via SheetJS
   ┌─────────────────────────────────────────────────────────┐
   │  • Average Gasolina 95 E5 / Gasóleo A across stations   │
   │  • Average Diesel / Gasolina consumption from fleet     │
   │  • Replace `flota` sheet with vendor rows (cols B..P)   │
   │  • Write new month row at row 145 — ALL VALUES,         │
   │    NO FORMULAS (kills the #REF! bug from copying rows)  │
   │  • Render email preview as HTML table matching the      │
   │    team's Outlook format (dark grey fuel cells,         │
   │    yellow tarifa cells, signature, branded footer)      │
   └─────────────────────────────────────────────────────────┘
                                ↓
   ┌──────────────────────────────────────────────┐
   │  Download <name>_updated.xlsx                │
   │  Copy email draft (rich clipboard for table) │
   └──────────────────────────────────────────────┘
```

## UI shape

- **Header** — Covestro rainbow-circle logo, dark navy bar (`#1a1a2e`), title in EN/ES, language toggle, "Download prices" button.
- **Help section** — `<details>` collapsible with full monthly procedure, what the build does, data sources, troubleshooting. Localised.
- **Live averages** — two cards (Gasolina 95 E5 / Gasóleo A), auto-fetched on load. Shows snapshot timestamp, station counts, min/max.
- **Step 1** — vendor `Flota Viva Covestro <date>.xlsx` (drag-drop or pick).
- **Step 2** — `Liquidaciones de combustible …xlsx` (drag-drop or pick).
- **Step 3** — calc date (defaults to today), email sign-off name + role.
- **Build** — produces preview + downloadable .xlsx + clipboard-ready email.

## Why the values-only write matters

The team's `Liquidaciones` sheet had a unique header-mid-sheet layout (history above row 144, "current month" on row 145, snippet header on 144). Each month, the user copies row 145's values into history above the header, then types a new month into row 145.

Previously, row 145's I/M cells used `=GETPIVOTDATA("Consumo Fabricante", consumo!$E$8)` referencing a pivot on the `consumo` sheet. When the row was promoted to history, those formulas continued to point at `consumo!E8` — but the pivot recomputes whenever fleet data changes, so old historical rows would show `#REF!` or wrong numbers.

The page now writes **plain numbers in all 9 cells** (A=date string, H/L=prices, I/M=consumption averages computed directly from the vendor file replacing the pivot, J=H, K=I·J/100, L, M, N=L, O=M·N/100). When promoted to history, these stay correct forever.

## Why the JSON API instead of the .xls

The original `geoportalgasolineras.es/.../preciosEESS_es.xls` does not send CORS headers. Browsers can't fetch it from any other origin including `null` (file://). The same data is also published as JSON at `https://sedeaplicaciones.minetur.gob.es/ServiciosRESTCarburantes/PreciosCarburantes/EstacionesTerrestres/` with `Access-Control-Allow-Origin: *`. Use the JSON path for everything new — see `memory/reference_geoportal_api.md`.

JSON keys are slightly different from the .xls column names:
- `Precio Gasolina 95 E5` (capitalised)
- `Precio Gasoleo A` (no accent on "Gasoleo")

## Hourly snapshot drift (not a bug)

The Ministry's feed regenerates **hourly** and overwrites the previous snapshot — past hours are not archived. If the user downloads `preciosEESS_es_<date>.xls` at 11:00 to spot-check, and the tool fetched the JSON at 12:00, the two sources describe Spain at different moments. Averages can differ by ~0.001–0.005 €/L for that reason alone. Confirmed: David's 08/05/2026 11:00 .xlsx → Gasóleo A 1.7240; tool's 12:00 snapshot → 1.7224. Same .xlsx for Gasolina 95 = 1.5467 in both.

Don't treat "tool ≠ manual by ~0.001" as a tool bug — first check `state.prices.snapshot` against cell B1 of the user's .xls.

**Pending (2026-05-10):** optional Step 1B prices-file picker — if user drops the `preciosEESS_es_<date>.xlsx`, tool averages cols J + O directly from that file; else falls back to live API. Makes manual = automatic by construction. Awaiting David's go-ahead via Dusan before implementing.

## Internationalisation

Two languages, EN and ES, switched via header toggle (defaults to ES if browser language starts with "es"; persists in `localStorage`). All UI strings, the help section, and email body strings (greeting, intro, sign-off, "Diesel"/"Gasoline" → "Gasoil"/"Gasolina", month names) come from a `I18N` dictionary near the top of the script.

The **email table headers stay in Spanish** in both languages — those are the actual column names in the Liquidaciones workbook and the recipients' tarifa email format.

## Files in the repo

- **`index.html`** — the entire app, ~1100 lines, no external runtime deps except SheetJS from CDN. Logo embedded as base64.
- **`fuel_prices.xlsx`** (gitignored on `main`) — independent Power Query workbook. Auto-refreshes the .xls on open. Useful as the M-code reference; could be migrated to use the JSON API.
- **`SETUP.md`** — Excel + Power Query + VBA recipe documenting an alternative implementation that lives entirely inside the workbook. Parked as a fallback; not active.
- **`docs/`** (gitignored on `main`, present on `docs-personal` branch only) — internal Covestro reference fixtures. **Do not commit to `main`** under any circumstance — the gitignore protects it.

## When changing this

- **Don't switch the price source back to the .xls.** Re-introduces CORS pain and BIFF parsing.
- **Keep the new month row in row 145 as plain values, never formulas.** The whole point of the build pipeline.
- **Don't add a build step.** This is one HTML file the user double-clicks. No bundlers, no frameworks, no npm.
- **Read by cell address, not `sheet_to_json({header: 1})`.** SheetJS strips leading empty rows in array-of-arrays mode, which broke our row indexing on the vendor file.
- **Test against `docs/`** if available locally — `/tmp/test_build.mjs` mirrors the browser logic in Node and validates against the real fixtures (vendor file produces 45 fleet rows; cons.diesel ≈ 4.756, cons.gasoline ≈ 5.788 matching the old pivot exactly).

## Repo / git

- `git push origin main` → Covestro repo. SSO-cached credentials on this WSL machine handle auth.
- `git push personal main` → personal mirror at `DusanOravsky/te-spain`. Uses `gh` CLI auth (run `gh auth login` once).
- **Never embed tokens in remote URLs or chat.** See `memory/feedback_no_chat_tokens.md`. If a token has been pasted, revoke it at https://github.com/settings/tokens immediately.

## Out of scope

- No automatic email sending (user pastes draft into Outlook themselves).
- No history of past months as a separate dashboard — the Liquidaciones workbook itself is the history of record.
- No Power Automate / Power BI / scheduled flows. Tier-1 in-browser app shipped; further automation parked unless the user asks.
