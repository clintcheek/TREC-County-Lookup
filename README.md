# Texas Broker County Resolver v6

Broker-office-first county enrichment with persistent checkpoint resume, version-aware reprocessing, confidence upgrades, stale-record rechecks, and append-only resolution history.

## Recommended first run

Use `upgrade_version` with `max_rows: 25`. This rechecks records produced by older resolver versions without discarding stronger existing results.

## Modes

- `new`: records absent from `state/results.csv`
- `recheck_unresolved`: unresolved and error records
- `recheck_review`: unresolved and needs-review records
- `upgrade_confidence`: anything below the configured confidence threshold
- `upgrade_version`: records produced by an older resolver version
- `recheck_stale`: records older than `recheck_after_days`
- `flagged`: rows marked `Needs Recheck = Yes`
- `recheck_all`: every record

Every attempt is appended to `state/resolution_history.csv`. Existing stronger results are retained when a recheck produces weaker evidence.
