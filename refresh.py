"""Download preciosEESS_es.xls, compute averages, write data.json.

Designed to run inside GitHub Actions. The workflow commits the resulting
data.json back to the repo; GitHub Pages serves it next to index.html.

Configuration lives in config.json so the source URL or column names can be
changed without touching code.
"""

from __future__ import annotations

import json
import sys
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

import xlrd

ROOT = Path(__file__).parent
CONFIG_PATH = ROOT / "config.json"
XLS_PATH = ROOT / "preciosEESS_es.xls"
OUT_PATH = ROOT / "data.json"


def to_float(v):
    if isinstance(v, (int, float)):
        return float(v) if v != "" else None
    s = str(v).strip().replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def download(url: str, dest: Path) -> None:
    req = urllib.request.Request(url, headers={"User-Agent": "te_spain/0.1"})
    with urllib.request.urlopen(req, timeout=120) as r, dest.open("wb") as f:
        f.write(r.read())


def main() -> int:
    cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    url = cfg["source_url"]
    header_row = int(cfg.get("header_row", 3))
    columns = cfg["columns"]
    snap_r, snap_c = cfg.get("snapshot_cell", [0, 1])

    print(f"Downloading {url}", flush=True)
    download(url, XLS_PATH)

    book = xlrd.open_workbook(str(XLS_PATH))
    sh = book.sheet_by_index(0)

    header = [sh.cell_value(header_row, c) for c in range(sh.ncols)]
    missing = [name for name in columns.values() if name not in header]
    if missing:
        raise RuntimeError(
            f"Column(s) not found in sheet header: {missing}. "
            f"Update config.json or check that the source format hasn't changed."
        )
    col_idx = {key: header.index(name) for key, name in columns.items()}

    snapshot_date = str(sh.cell_value(snap_r, snap_c)).strip()

    results = {}
    for key, idx in col_idx.items():
        values = []
        for r in range(header_row + 1, sh.nrows):
            v = to_float(sh.cell_value(r, idx))
            if v is not None:
                values.append(v)
        avg = sum(values) / len(values) if values else None
        results[key] = {
            "label": columns[key],
            "stations": len(values),
            "average_eur_per_l": round(avg, 4) if avg is not None else None,
            "min": min(values) if values else None,
            "max": max(values) if values else None,
        }

    payload = {
        "source_url": url,
        "snapshot": snapshot_date,
        "fetched_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "total_rows": sh.nrows - (header_row + 1),
        "fuels": results,
    }

    OUT_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2))
    print(f"Wrote {OUT_PATH}", flush=True)
    for key, info in results.items():
        print(f"  {info['label']}: {info['average_eur_per_l']} €/L  ({info['stations']} stations)", flush=True)

    XLS_PATH.unlink(missing_ok=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
