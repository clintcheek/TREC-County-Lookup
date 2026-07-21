# V7 Release Notes

## Replace
- `resolver.py`
- `config.json`
- `.github/workflows/enrich.yml`
- `README.md`

## Keep
- `input/brokers.xlsx` (replace with your production workbook when publishing)
- `requirements.txt`
- existing `state/` and `output/` files when upgrading an active repository

## Safe rollout
1. Back up the v6 repository.
2. Publish the four replacement files.
3. Run `upgrade_version` for 25 rows.
4. Audit results before scaling to 100 or more rows.

V7 does not delete prior results. The existing stronger-result retention behavior remains active.
