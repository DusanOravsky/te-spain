# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project does

Tiny web app that surfaces the average **Precio gasolina 95 E5** and **Precio gasóleo A** from the Spanish national fuel-price spreadsheet (`preciosEESS_es.xls`) published by `geoportalgasolineras.es`.

The user runs this once a month and uses the averages as input for a downstream workflow (other steps not in scope).

## Architecture

The browser cannot fetch the upstream .xls directly — `geoportalgasolineras.es` does not send CORS headers, and the user's Windows machine has no Python/exe permissions. Instead, the work happens in **GitHub Actions** and the page (delivered as a single HTML file the user double-clicks) reads the result via the GitHub API.

```
   user's PC                                GitHub
   ┌──────────────────┐                    ┌─────────────────────────────────┐
   │  index.html      │   dispatch (POST)  │  refresh.yml → refresh.py       │
   │  (file:// open)  │ ─────────────────► │   downloads .xls                │
   │                  │                    │   computes averages             │
   │  button click    │                    │   commits data.json to main     │
   │                  │   GET contents     │                                 │
   │                  │ ◄───────────────── │  /repos/.../contents/data.json  │
   └──────────────────┘                    └─────────────────────────────────┘
            ▲
            │ user's fine-grained PAT in localStorage
            │ (Actions: write, Contents: read)
```

- **`refresh.py`** — server-side script. Downloads `preciosEESS_es.xls`, parses the two configured columns (skipping empty cells), writes `data.json`. Reads everything tunable from `config.json`.
- **`config.json`** — source URL, header row index, snapshot cell, and column-name → key map. **Edit this if the upstream format changes — never hardcode values back into `refresh.py`.** The script fails loudly with a "Column(s) not found" error pointing here.
- **`.github/workflows/refresh.yml`** — runs `refresh.py` on `workflow_dispatch` (button click) and on a monthly cron (1st @ 06:00 UTC). If `data.json` changed, force-pushes it to a `refresh-data` branch, opens or updates a PR against `main`, and enables auto-merge (squash). The page polls `main` until the merge lands. **Why PR mode:** the Covestro org enforces `cov-secure-policy` on `main` (PRs required + CodeQL gate); a direct push from the bot is rejected. Single-flight via `concurrency: refresh`.
- **`index.html`** — single static file. Opened from the user's filesystem (no server). On load with credentials configured: fetches `data.json` via the GitHub Contents API and renders. On "Download prices" click: dispatches the workflow, polls `/actions/runs` until completion, then re-fetches `data.json`. The PAT and repo are stored in `localStorage` (set via the ⚙ button). The HTML file itself contains **no** secrets and is safe to email or share.

**Why the GitHub API instead of `raw.githubusercontent.com`?** `raw.githubusercontent.com` doesn't send `Access-Control-Allow-Origin`, so a `file://`-loaded page can't fetch from it. The Contents API (`api.github.com`) does, and the same PAT works for both dispatch and read.

## Source-file quirks worth knowing

The upstream is real BIFF (CDFV2 Microsoft Excel), not .xlsx — requires `xlrd` 2.x. Do not "upgrade" to a version that drops .xls support, and do not switch to `openpyxl` (it can't read this format).

Spanish locale: comma is the decimal separator (`1,449`). `to_float()` normalizes `,` → `.`. Empty cells (stations not selling that fuel) are skipped from the average — they are not zero.

## Why this shape (constraints that ruled out simpler options)

- The user's Windows machine: **no Python, no .exe, no manual file downloads.** Rules out a local Python server, PyInstaller bundles, and the user-uploads-the-xls flow.
- The geoportal sends no CORS headers → static HTML can't fetch the .xls directly. Rules out a pure-static SPA.
- A self-hosted CORS proxy would also work, but GitHub Pages + Actions removes the need to operate any infrastructure.

## Setup (one-time, per deployment)

1. Push this repo to GitHub.
2. **Settings → Actions → General → Workflow permissions: Read and write.**
3. Create a fine-grained PAT scoped to this repo with **Actions: Read and write** + **Contents: Read-only**. Org PATs may need admin approval.
4. Email the user `index.html`. They open it (double-click → opens in browser via `file://`), click ⚙, paste `owner/repo` + the PAT.

## Commands

```bash
# run the refresh script locally (writes data.json)
pip install -r requirements.txt
python refresh.py

# preview the page locally (data.json must exist)
python -m http.server 8000
```

No tests, no linter, no build step.

## Out of scope / known gaps

- The user mentioned "a few other steps" after the averages — not implemented; structure deliberately doesn't anticipate them.
- No history / month-over-month series. Each refresh overwrites `data.json`. Adding history would mean appending to a JSON file or keeping commits as the timeline.
- The frontend hardcodes the two fuel keys (`gasolina_95_e5`, `gasoleo_a`); making it data-driven from `config.json` is a small additive change, not done by design.
- Token lives in the user's browser `localStorage`. If they use multiple browsers, they re-enter it. If the token leaks, worst case is someone else can also click the button — they cannot push code or read other repo secrets (fine-grained scope).
