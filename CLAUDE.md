# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

A small web app that shows the Spain national average prices for **Gasolina 95 E5** and **Gasóleo A**, fetched live from the Spanish Ministry of Energy's REST API. The user opens `index.html` from disk and clicks **Download prices**.

Originally part of a larger month-end fuel-liquidation workflow (vendor fleet file → Liquidaciones workbook → tarifa email). The price-display side is now a finished one-file static page; the broader workflow is parked (see "What's parked" below).

## Architecture (current)

```
   index.html  ──fetch──►  https://sedeaplicaciones.minetur.gob.es/
   (file:// or                ServiciosRESTCarburantes/
   anywhere)                  PreciosCarburantes/EstacionesTerrestres/
                              (CORS-open JSON, ~12 MB, ~11,400 stations)
        │
        ▼
   averages two columns
   ("Precio Gasolina 95 E5", "Precio Gasoleo A")
   skipping empty cells, displays cards
```

That's the whole app. Single file. No build step, no SheetJS, no proxy, no backend, no install.

## Why this works (and didn't, for ages)

The original data source we reached for was `https://geoportalgasolineras.es/.../preciosEESS_es.xls`. That URL **does not send CORS headers**, so a browser can't fetch it from any other origin (including `null` / `file://`). That CORS wall blocked every "static HTML" plan we tried.

The same data is **also** published as a JSON REST API at `sedeaplicaciones.minetur.gob.es`, on a different domain, **with `Access-Control-Allow-Origin: *`**. Once we found that endpoint, the simple "static HTML with a button" architecture became possible. See `memory/reference_geoportal_api.md` for the full endpoint reference.

## Key technical facts

- **API endpoint**: `https://sedeaplicaciones.minetur.gob.es/ServiciosRESTCarburantes/PreciosCarburantes/EstacionesTerrestres/` (GET, JSON, CORS-open).
- **JSON keys differ from .xls column names** — JSON uses `Precio Gasolina 95 E5` (capitalized) and `Precio Gasoleo A` (no accent on "Gasoleo"). The .xls had lowercase keys with accents.
- **Spanish decimal comma** — every price is a string like `"1,449"`. Normalize `,` → `.` before parsing as Number.
- **Empty string** means "this station doesn't sell that fuel" — skip it from the average, don't treat as zero.
- **Snapshot timestamp** is on the top-level `Fecha` field of the JSON, format `dd/MM/yyyy HH:mm:ss`.

## Files in the repo

- **`index.html`** — the entire app. ~190 lines, no external deps at runtime.
- **`fuel_prices.xlsx`** — Power Query workbook (the user wired this himself; auto-refreshes from the .xls on open). Independent of `index.html`. Could be migrated to use the JSON API instead of the .xls — small M-code change.
- **`SETUP.md`** — Excel + Power Query + VBA recipe for a more ambitious automation (read vendor file, write to Liquidaciones, generate email). Not built; documented as a fallback if the broader month-end workflow needs to be picked up.
- **`docs/`** (gitignored) — internal Covestro reference fixtures. Don't commit.

## What's parked (not gone, not active)

A more ambitious flow combines (a) live prices, (b) a vendor "Flota Viva" Excel file from Ayvens, (c) the running `Liquidaciones de combustible` workbook, into one updated workbook plus an email draft. This would replace the user's current ~15 minute monthly process. Two paths are designed and ready if needed:

- **Power Automate flow** (in Covestro's M365 tenant) — fully automated; user just reviews the Outlook draft.
- **Excel + Power Query + VBA** workbook — see `SETUP.md`; user clicks Refresh All + a "Close month" button.

Neither was built because the immediate ask was "show me the two prices on a web page", which the current `index.html` does cleanly.

## When changing this

- **Don't switch the price source back to the .xls.** It re-introduces the CORS wall and the BIFF-parsing pain. If extra fields are needed, they're almost certainly in the JSON.
- **Don't add a build step.** This is one HTML file, deliberately. No bundlers, no frameworks. The whole point is "the user double-clicks it."
- If the API ever returns 5xx or changes shape, the failure modes the user would see: `Failed: HTTP …` in the status line. The page does not retry on its own.

## Repo / git

`covestro/accounting-spain-gas` (private, SSO, `cov-secure-policy` rules on `main`). Push works through SSO-cached credentials — never embed tokens in URLs or chat. See `memory/feedback_no_chat_tokens.md`.
