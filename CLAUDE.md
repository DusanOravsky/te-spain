# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

Automation of Covestro Spain's monthly fuel-price liquidation: the user gathers Spanish national fuel-station prices and the latest fleet roster from a vendor, computes consumption-weighted tarifas, and emails the result to colleagues. The project explores the right shape for that automation given the user's corporate environment.

## Current direction (2026-05-07)

**Power Automate flow in Covestro's M365 tenant.** The flow:

- Fetches `preciosEESS_es.xls` server-side (no CORS, no TLS-chain pain on Microsoft's side)
- Reads the newest `Flota Viva Covestro *.xlsx` from a SharePoint/OneDrive folder
- Computes price + consumption averages
- Writes them into row 145 of `Liquidaciones de combustible …xlsx` via the Excel Online connector
- Creates an Outlook draft with the email body inline

User's monthly job: drop vendor file in folder → review draft → click Send.

The flow is **not yet built** — gating questions on SharePoint location, trigger model, sender mailbox, recipients, and HTTP-action availability in Dusan's M365 plan are still open.

## Why this shape

Every other architecture hit a wall in the Covestro environment. See `memory/project_covestro_constraints.md` for full detail; summary:

- **Browser-only static HTML** can't fetch the geoportal — no CORS headers.
- **GitHub Actions + Pages** can't commit the result back — `cov-secure-policy` requires PR + CodeQL.
- **Local Python / .exe** — Python isn't installed, .exes are blocked.
- **Excel + Power Query alone** works (see `SETUP.md`) but still requires the user to refresh manually and run a VBA close-month macro.
- **External hosting** (Cloudflare Worker, Azure Function) was rejected by the user.

Power Automate is the first option that's both fully automated and inside Covestro's already-approved infrastructure.

## What's currently in this repo

- **`index.html`** — static HTML with three file pickers (vendor / Liquidaciones / preciosEESS .xls). Browser-side merge via SheetJS; produces an updated Liquidaciones xlsx + an email draft. Verified end-to-end against `docs/`. **Kept as a fallback** for one-off use; not the production path.
- **`fuel_prices.xlsx`** — small Power Query workbook the user wired himself. Auto-refreshes the geoportal on open and writes Gasolina 95 E5 / Gasóleo A averages to a `Prices` sheet. Useful M-code reference.
- **`SETUP.md`** — detailed instructions for the **Excel + PQ + VBA fallback** path (Tier 1 in our planning conversation). 7 steps; ~15 min one-time setup. Use this if Power Automate falls through.
- **`docs/`** (gitignored) — reference fixtures: vendor file, current Liquidaciones, sample preciosEESS .xls. Internal Covestro data — don't commit.

## Key technical facts (the things that have bitten us)

- The geoportal serves real **BIFF** (`.xls`, CDFV2), not `.xlsx`. Tools that handle it: SheetJS, `xlrd` 2.x, Excel, Power Query. Tools that don't: `openpyxl`.
- The geoportal **does not send CORS headers**. Browsers (any origin including `null` / `file://`) block fetches. Server-side fetchers are fine.
- Header is on row 4 (index 3); data from row 5. Snapshot timestamp in cell `B1`. Spanish decimal comma (`1,449`) — normalize before parsing.
- The geoportal **TLS chain is incomplete** — strict verification fails on clean runners. `curl -k` or `ssl._create_unverified_context()` are acceptable mitigations because the data is public and unauthenticated.
- The Liquidaciones email-snippet structure is unusual: the table is **in the middle of the sheet** (header on row 144, current month on row 145, history above). The user copies header + one row into Outlook each month. Anything that automates this needs to write to row 145 and leave the surrounding structure alone. Old `consumo` pivot at `consumo!E8/E17` feeds I and M via `GETPIVOTDATA` — replaced in our automation by direct averages computed from the vendor file.

## Org / git

Repo: `covestro/accounting-spain-gas` (private, SSO, `cov-secure-policy` rules). `git push` works through SSO-authorized cached credentials — never embed tokens in URLs or pipe them through chat. See `memory/feedback_no_chat_tokens.md`.

## Out of scope

- Power BI / Power Apps frontends — over-engineered for this volume of work.
- Standalone executables — corporate environment blocks them.
- Public hosting (Vercel, Cloudflare, Render) — rejected by the user.
- Auto-approving the bot's PR to bypass `cov-secure-policy` — declined as a policy-evasion pattern.
