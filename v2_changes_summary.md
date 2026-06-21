# Trade Reconciliation Utility v2 — What Changed and Why

## Root cause of the inaccurate v1 results

v1 grouped trades by business attributes (instrument + direction + dates +
price) and effectively ignored Trade ID. Your files DO carry a reliable
shared ID on every trade (Broker "Unique Trade ID" / "Orig ID" = Murex
"G.ID", once a system prefix like "BFU:" is stripped) — but v1 never used
it as the primary key, so two trades with the same ID but a slightly
different price/date format fell into different groups and were reported
as "all mismatched" even though they were the same trade.

v2 rebuilds the engine around the process described in your document,
step by step.

## How v2 reconciles (mirrors your manual process exactly)

1. **Trade count check** — total rows in each file (Executive Summary).
2. **ID extraction & matching** — every link ID is normalised by stripping
   known system prefixes (`BFU`, `ICE`, `GID`, `BBG`, `TPI` — editable in
   the UI) and leading zeros, so `BFU:75287073` and `75287073` are
   recognised as the same trade.
3. **Per-ID field comparison** — for every ID present in both files:
   - Broker legs for that ID are summed (lots) — handles the "3 legs of 10
     lots = 1 Murex line of 30 lots" case exactly as described.
   - Direction is checked (Bought/Sold ↔ Buy/Sell normalised to BUY/SELL).
   - Price is checked (with a configurable tolerance).
   - Quantity total is checked (broker sum vs Murex nominal).
   - **PASS** only if all three line up. **FAIL** otherwise, with the
     specific field, broker value, and Murex value spelled out — e.g.
     *"Price mismatch (Broker: 60.0 | Murex: 58.0)"*.
4. **Missing trades** — IDs only in the broker file → Missing in Murex;
   IDs only in Murex → Missing in Broker.
5. **Fallback (rare case)** — if a row genuinely has no usable ID on either
   side, the old attribute-based matcher kicks in as a safety net so
   nothing is silently dropped — but this is no longer the primary
   mechanism, just a backstop.

## Verified against your document's exact example

Using the numbers from your Word doc (ID 60: three legs of 10 lots @ price
45 vs Murex 30 lots @ 45; ID 61: 25 lots @ 50 vs Murex 25 @ 50; a price
break and a missing trade added for testing):

| Link ID | Outcome |
|---|---|
| 60 | **PASS — Aggregated** (3 broker legs × 10 lots = 30, matches Murex 30 lots @ 45) |
| 61 | **PASS — Single leg** (25 lots @ 50 on both sides) |
| 72 | **FAIL — Price mismatch (Broker: 60.0 \| Murex: 58.0)** |
| 73 | **Missing in Murex** (no corresponding Murex ID found) |

This is exactly the classification your manual process describes.

Also re-tested against your real `SFTP-File.xlsx` / `Murexextracttest.xlsx`
pair: both linked trades (`922776111` → 3 legs summing to 32 lots,
`777036614` → 5 legs summing to 24 lots) now correctly show as **PASS —
Aggregated**, instead of "all mismatched" as before. And the full 1,700-row
Tullett Prebon file reconciles in under 6 seconds.

## Report structure (renamed/refocused around ID-level results)

| Sheet | Purpose |
|---|---|
| Executive Summary | Row counts, unique ID counts, PASS/FAIL/Missing counts, % clean |
| ID-Level Reconciliation | One row per Link ID — PASS/FAIL, leg counts, totals, and exact issue text |
| Matched (PASS) | Full leg-level detail for every passed ID (single and aggregated) |
| Mismatched (FAIL) | Full leg-level detail with reason/broker value/Murex value per ID |
| Missing in Murex | Broker IDs with no Murex counterpart |
| Missing in Broker | Murex IDs with no Broker counterpart |
| Aggregated Matches Detail | Just the multi-leg rollups — broker lots per leg, sum, Murex sum, PASS/FAIL |
| Exception Analysis (No-ID) | Anything that had no usable ID — fallback matches and true unmatched exceptions |

## What you need to do to use it correctly

In the **Column Mapping** panel, make sure **`link_id`** is mapped to:
- Broker side: whichever column holds the value that's also embedded in
  Murex's G.ID — in your files this is `Orig ID` or `Unique Trade ID`
  depending on the broker template.
- Murex side: `G.ID`.

If your Murex G.ID uses a prefix not in the default list (`BFU`, `ICE`,
`GID`, `BBG`, `TPI`), add it in the **"ID prefixes to strip"** box in the
Matching Configuration panel — comma-separated, no need to restart.

## File

`trade_recon_app.py` — single file, same run command as before:
```
pip install -r requirements.txt
streamlit run trade_recon_app.py
```
