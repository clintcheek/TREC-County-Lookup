# Texas Broker County Resolver v7

V7 keeps the v6 checkpointing, resume, threading, cache, workbook, Census geocoder, audit, and history infrastructure, while replacing the single-address acceptance engine with county evidence voting.

## What changed

- Evidence is grouped and voted by Texas county.
- Government and likely official-office sources receive the strongest weight.
- Independent domains increase confidence; repeated pages from one domain cannot overwhelm the result.
- Property-listing evidence is down-weighted and cannot independently produce `Verified`.
- New result tiers: `Verified`, `Very Likely`, `Likely`, `Needs Review`, and `Unresolved`.
- Existing stronger results are not downgraded during rechecks.

## Recommended first test

Run the GitHub Action manually with:

- mode: `upgrade_version`
- max_rows: `25`
- checkpoint_every: `5`

Review `output/brokers_enriched.xlsx`, `state/candidates.csv`, and `state/resolution_history.csv` before increasing the batch size.

## Required secret

Create the GitHub Actions repository secret `SERPER_API_KEY`.

## Modes

- `new`
- `recheck_unresolved`
- `recheck_review`
- `upgrade_confidence`
- `upgrade_version`
- `recheck_stale`
- `flagged`
- `recheck_all`
