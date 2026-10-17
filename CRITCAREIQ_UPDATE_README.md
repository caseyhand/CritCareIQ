# CritCareIQ monthly update

One command regenerates the CCAR data in `index.html` from the source spreadsheets.

## One-time setup
```
pip3 install openpyxl
```
Keep `critcareiq_update.py` next to the files below (or anywhere — paths are arguments).

## Each month
1. Save Wendy's new Vituity E&M Billing report into the `Critical Care Data` folder (any subfolder; the script finds every `.xlsx` with "billing" in the name and reads the month from the sheet itself, so file naming doesn't matter — newest copy of a month wins).
2. Save the latest weekly `pts per hour physicians unblinded 2026.xlsx`.
3. Preview:
   ```
   python3 critcareiq_update.py \
     --billing "Critical Care Data" \
     --pph "pts per hour physicians unblinded 2026.xlsx" \
     --index index.html --add-months
   ```
   It prints each billing month found, which ESI months are complete vs partial, the newest month's group numbers, and any physician in billing who isn't in the PHYSICIANS table.
4. Add `--write` to save. The original is kept as `index.html.bak`.
5. Hand-edit the five things the script doesn't own (they're one line each, all near the top):
   `PARTIAL_KEY` (newest month), `LATEST_FINAL` (month before), `STATUS_KEYS` (last 3 final months), `CODED` (% coded, from Wendy's email), `DATA_ASOF`.
   Also the header badge ("NOV 2025 – SEP 2026") and footer date, and `CC_NIGHTS` if a nocturnist joins or leaves.
6. Push to GitHub; Vercel redeploys.

## What it regenerates
Everything between `/* @@CC_DATA START */` and `/* @@CC_DATA END */`:
- `CC_MONTH_KEYS`, `CC_COUNTS` — cases / CC charts / Level 5 charts per physician-month (billing)
- `ESI_COMPLETE`, `ESI_PARTIAL`, `ESI_COUNTS` — pts / ESI 1–5 / hours per physician-month (weekly sheet)

`REPL_RAW` (the 4.5% pre-series replacement level) is read from the existing file and preserved.

With `--add-months`, any billing month not yet in `MONTHS` / `MONTH_KEYS` / `GROUP_RATES` / `PHYSICIANS` is appended, so the older tabs update too.

## ESI completeness
A month counts as complete when every weekly sheet whose majority of days fall in that month has data, and the sheet series runs past the month's end. Partial months show in ESI mode's yellow note and are excluded from ESI-pool rankings. Override with `--esi-complete may26 jun26 …` if the auto-detection disagrees with you.

## If something looks wrong
- Run without `--write` and read the printout first.
- `index.html.bak` is the previous version.
- The script never touches anything outside the markers except the four rate tables under `--add-months`.
