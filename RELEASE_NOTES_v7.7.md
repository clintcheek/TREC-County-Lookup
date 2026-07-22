# v7.7 — Fifteen Dynamic Workers

## Production focus

v7.7 keeps the proven v7.6 resolver and increases the default in-process worker pool from five to fifteen. Python's shared executor queue dynamically assigns the next available brokerage to each free worker, so slow records do not leave fixed worker partitions idle.

## Changes

- New default mode: `parallel_15`.
- Default worker count: 15, configurable from GitHub Actions.
- Legacy `parallel_5` remains accepted for backward compatibility.
- Worker count is safety-capped at 32.
- Live progress includes completion percentage, processing rate, remaining records, and ETA.
- Workflow validates `resolver.py` with `py_compile` before starting.
- Existing checkpoint CSV, search cache, history, candidate audit, and workbook formats are retained.

## Recommended rollout

1. Run `100` records with mode `parallel_15`, workers `15`.
2. Inspect the output and logs.
3. Run `1000` records.
4. Proceed to the full eligible population after validation.
