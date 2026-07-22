# Texas Broker County Resolver v7.7

V7.5 keeps the company-first address and county resolver from v7.7 and adds practical contact enrichment.

## Primary output

The generated workbook now starts with a **Clean Results** worksheet containing:

- License number
- Brokerage name
- Related broker name
- Office address, city, state and ZIP
- County
- Office phone, when readily available on a discovered page or search result
- Brokerage website, when a likely official company website is available
- Up to two verification websites
- Resolution status and confidence

The original input worksheet is preserved and receives the full audit and resolver fields.

## Search order

1. Exact brokerage name and office/contact terms.
2. Brokerage name plus brokerage or individual license number.
3. Individual broker plus brokerage name.
4. Individual plus brokerage listings.
5. Individual, company and license numbers on listings.

A confirmed company office address stops the search. Listing addresses are used only for operating-county inference when an office cannot be confirmed.

## Website and phone handling

The resolver reuses pages already discovered during address resolution. It records a likely official brokerage website and a readily visible office phone without requiring a separate search stage. Directory, government, listing and social-media domains are not labeled as official brokerage websites.

## Upgrade from v7.7

Replace the program files with this package, retain your `input/` and `state/` folders, and run with `--mode upgrade_version` to revisit existing records and populate the new website and phone fields.

## v7.6 fifteen-worker mode

For the normal combined repair run, launch GitHub Actions with **Mode = `parallel_15`**. The resolver classifies new, related-broker-without-county, out-of-state, unresolved, and review records into non-duplicating priority queues and processes up to fifteen records simultaneously.
