# Texas Broker County Enrichment

This private GitHub workflow enriches the supplied Texas broker workbook with office address, county, evidence URL, confidence, and status. Progress is checkpointed in `state/results.csv`, so every run resumes where the previous run stopped.

## One-time setup

1. Create a **private** GitHub repository.
2. Upload every file and folder from this package to the repository root. Keep the folder structure intact, including `.github/workflows/enrich.yml`.
3. Create a Serper API key at Serper.dev. The search API is used to locate public broker office-address evidence.
4. In the GitHub repository, open **Settings → Secrets and variables → Actions → New repository secret**.
5. Name the secret exactly `SERPER_API_KEY` and paste the key as its value.
6. Open **Settings → Actions → General**. Under **Workflow permissions**, select **Read and write permissions**, then save.

## Start the job

1. Open the repository's **Actions** tab.
2. Select **Enrich Texas Broker Counties**.
3. Click **Run workflow**.
4. Leave `500` for the first run, then click the green **Run workflow** button.

The workflow runs again automatically every six hours and resumes from its checkpoint. You can also manually run it whenever desired.

## Get the latest results

The latest workbook is always stored at:

`output/brokers_enriched.xlsx`

It is also attached to each completed Actions run under **Artifacts → texas-brokers-enriched**.

## Result columns

- Office Address
- City
- State
- ZIP
- County
- Resolution Status
- Confidence
- Evidence Type
- Evidence URL
- Resolution Notes
- Last Updated UTC

`Resolved` means the evidence score met the configured threshold. `Needs Review` means a geocoded address was found but the source confidence was lower. `Unresolved` means no reliable geocodable public address was found during that record's search.

## Controls

Edit `config.json` to change concurrency, request delay, confidence threshold, or default batch size. Larger batches consume more search API credits.
