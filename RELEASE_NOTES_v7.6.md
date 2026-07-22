# v7.6 — Five Parallel Workers

## Main change

The GitHub workflow now starts the resolver with five concurrent workers inside one protected workflow run. This is faster than five sequential passes while avoiding multiple GitHub jobs writing the same workbook, cache, or checkpoint simultaneously.

## New `parallel_5` mode

Every eligible row is assigned to exactly one priority queue:

1. `new` — no prior result
2. `broker_no_county` — related broker is present but county is missing
3. `out_of_state` — saved state is not Texas
4. `unresolved` — status is Unresolved or Error
5. `review` — Very Likely, Likely, Needs Review, or explicitly flagged

A row can qualify for several conditions, but priority assignment ensures it is submitted only once.

## Safety

- One GitHub workflow run owns the shared state files.
- Five in-process workers handle different records concurrently.
- The pending list contains each license only once.
- Existing atomic checkpoint writes remain in place.
- GitHub concurrency prevents two complete v7.6 runs from writing the repository at the same time.

## Workflow default

Choose `parallel_5` when manually launching the workflow. The workflow passes `--workers 5` automatically.
