# CritCareIQ monthly update (v2 — from the CC Data repository)

The site's data now comes from `CC Data/raw/` (your own Vituity Analytics exports), not from
Wendy's spreadsheets or screenshots. One script regenerates every data block in `index.html`.

## Each month

1. **Pull the month from Vituity Analytics** (Tableau): EM Billing view, Provider Type = MD/DO,
   set the month → Download ▸ Crosstab ▸ CSV. Save it as
   `CC Data/raw/EM_Billing_MD_YYYY-MM_pulled_YYYY-MM-DD.csv`.
   Re-pull the previous month at the same time (it won't have been fully coded on the first pull)
   and keep both files — the newest pull date wins.
2. **Rebuild the repository tables** (optional but keeps everything in sync):
   ```
   cd "Critical Care Documentaion Project 2026"
   python3 "CC Data/ingest.py"
   ```
3. **Regenerate the site** — preview first, then write:
   ```
   python3 "Claude Crit Care/critcareiq_update.py" --billing "CC Data/raw" --index <path to repo>/index.html --refresh-rates
   python3 "Claude Crit Care/critcareiq_update.py" --billing "CC Data/raw" --index <path to repo>/index.html --refresh-rates --write
   ```
   It prints the group table for every month and backs up the old file as `index.html.bak`.
   Add `--pph "<weekly pts per hour sheet>.xlsx"` when the ESI sheet is current (needs `pip3 install openpyxl`);
   without it the ESI block already in the file is kept.
4. **Hand-edit the few things the script doesn't own** (all near the top of `index.html`):
   `PARTIAL_KEY` (newest month), `LATEST_FINAL` (month before), `STATUS_KEYS` (last 3 final months),
   `CODED` (% coded for the provisional month — from Wendy, since Analytics doesn't show it), `DATA_ASOF` (pull date).
   Also the header badge ("NOV 2025 – SEP 2026") and the footer line, and `CC_NIGHTS` if a nocturnist joins or leaves.
5. **Push to GitHub**; Vercel redeploys.

## What the script regenerates

- Between `/* @@CC_DATA START */` and `/* @@CC_DATA END */`: `CC_MONTH_KEYS`, `PRE_SERIES_KEYS`,
  `CC_COUNTS` (cases / CC / Level 5 per physician-month), and the ESI block (only with `--pph`).
  `REPL_RAW` is read from the existing file and preserved.
- With `--refresh-rates`: `MONTHS`, `MONTH_KEYS`, `GROUP_RATES`, and the whole `PHYSICIANS` table,
  rebuilt from billing for every month present. `lowVol` is set automatically when a physician's busiest
  month is under 50 cases. (`--add-months` is the old behaviour: append only what's missing.)

## What changed on the site (10/22/26)

- All 11 months (Nov '25–Sep '26) come from one Analytics pull; the Nov–Mar numbers that used to come
  from screenshots are now the real report (Jan/Feb at 4.5%, Nov/Dec/Mar at 5.2–5.3%).
- CCAR tab: windows now include **Series to date** (Apr–Sep) and **Pre-series** (Nov–Mar) alongside the
  rolling 3-month windows; a fourth baseline, **vs own pre-series**, measures each physician against their
  own Nov–Mar conversion (answers "did I change?" rather than "how do I compare?").
- New flags: *Up vs own pre-series*, *Review: down vs own pre-series*, *No pre-series months*.
  "Review: down vs spring" is retired.
- Overview chart marks the series start; sparklines in the CCAR detail view do too.
- Stat Guide: new "vs own pre-series" card; the CCAR card now states that 4.5% is the Jan–Feb low and
  that the Nov–Mar average is ~4.9%.

## Baseline note

`REPL_RAW = 0.045` is the Jan–Feb '26 low. The five-month pre-series average is 4.9%. If you decide to
move to 4.9%, change the one constant in `index.html` (and `REPLACEMENT_RATE` in `CC Data/ingest.py`) and
update the wording on the CCAR card in the Stat Guide.

## If something looks wrong

- Run without `--write` and read the printout first.
- `index.html.bak` is the previous version.
- The script never touches anything outside the markers except the four rate tables under `--refresh-rates` / `--add-months`.
