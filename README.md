# Texas Broker County Resolver v7.4

V7.4 resolves the brokerage entity and office address before using listing activity.

## Search order

1. Exact brokerage name + address/contact terms.
2. Brokerage name + brokerage or individual license number.
3. Individual broker + brokerage name.
4. Individual + brokerage listings.
5. Individual + company + both license numbers on listings.

A confirmed exact-company office address stops the search and becomes the primary county. Property listing addresses are excluded from office-address selection. Listing concentration is used only as the final operating-county fallback.

## Optional Gemini fallback

Set `GEMINI_API_KEY` and change `enable_gemini_fallback` to `true` in `config.json`. Gemini is called only after deterministic company and person/company searches fail to produce a verified office. It uses Google Search grounding and is instructed not to treat property listings as office addresses.

## Recommended validation run

- Mode: `upgrade_version`
- Max rows: `25`
- Checkpoint every: `5`

Keep the existing `input/`, `state/`, and `output/` data when upgrading.
