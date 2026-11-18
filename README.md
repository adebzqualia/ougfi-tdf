# Positive Balance reporting automation

This project turns a quarterly folder of country Excel files and exchange rates into the detailed Positive Balance Excel report. It preserves a separate history so the next quarter can compare against the previous one.

The output contains the global, CORE, NON CORE and individual country pages, with native Excel charts. Months run from oldest to newest. Visible input tabs let you adjust existing values directly in Excel; formulas update the presentation tables, chart data and numerical comments. Internal calculation sheets remain hidden; file paths and technical source columns are not displayed.

Country views use local currency, including **GBP for UK and ZAR for South Africa**; consolidated tables use EUR. Efficiency charts use a 0–100% axis. Review out-of-range input results instead of interpreting them as a valid 100% rate.

## Quick start

1. Extract the ZIP into a local working folder.
2. Open a terminal in that folder and install the dependencies:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

3. Put the quarterly inputs in one folder, for example:

```text
inputs/
  Q4_2026/
    FRANCE.xlsx
    GERMANY.xlsx
    AUSTRIA.xlsx
    BELGIUM.xlsx
    ...
    NORDICS.xlsx
    BRAZIL.xlsb
    ROMANIA.xlsx
    Taux.xlsx
```

4. Run the full sequence:

```powershell
.\.venv\Scripts\python.exe run_reporting.py --input inputs/Q4_2026
```

The result is written to `reports/Synthese_detaillee_2026Q4.xlsx`. The input Excel files are read without modification.

Countries use a configured Base or KPI route. Standard KPI files contain `KPI_1`, `KPI_2` and `KPI_3`. Nordic summary sheets have country prefixes such as `DEN_KPI_2`, `NOR_KPI_2` and `SWE_KPI_2`. Brazil and Romania use their configured special formats. File aliases, currencies and sheet mappings are in the configuration files; naming rules are not embedded in the command. The country files remain unchanged.

## Four separate steps

Run these commands instead of the wrapper when you want to inspect the intermediate results:

```powershell
python 01_extract.py --input inputs/Q4_2026
python 02_validate.py --period 2026Q4
python 03_update_history.py --period 2026Q4
python 04_generate_report.py --period 2026Q4
```

| Step | What it does | Main result |
|---|---|---|
| 1. Extract | Reads configured Base rows or KPI summaries and closing rates into one consistent structure. Saves an immutable copy of the extraction configuration and its hash. | `work/2026Q4/extracted.json`, `config_snapshots/<hash>.json` |
| 2. Validate | Checks dates, currencies, country coverage, required data and reconciliations; applies explicit quarter-specific overrides. | Validated data and a validation report in `work/2026Q4/` |
| 3. Update history | Adds the validated quarter or replaces it with a numbered revision; identical reruns leave history unchanged. Keeps other quarters and backs up the whole store. | `database/history.json` and `database/backups/` |
| 4. Generate report | Reads the selected quarter and its predecessor, calculates comparisons and writes the detailed Excel. | `reports/Synthese_detaillee_2026Q4.xlsx` |

Read [the step-by-step explanation](docs/HOW_IT_WORKS.md) for examples of the data at each stage.

Code follows the same separation: `pb_automation/extract.py` and `base.py` read files; `validate.py` checks them; `history.py` saves revisions; `reporting/` creates the Excel pages. `run_reporting.py` is only the wrapper. `common.py` contains shared period, hash and file-writing helpers.

## Base sheets versus KPI sheets

`config/base.json` controls routing per country. Austria, Belgium, France, Germany, Denmark, Norway, Sweden, Portugal, Spain and UK are declared Base-capable. **Verified profiles currently cover the supplied Nordic and Austrian layouts.** The other countries need their own verified column/status/population mapping; a sheet named `Base` is insufficient evidence that another country's rules apply.

| Mode | Behavior |
|---|---|
| `base_if_present` | Use a mapped Base when present. A present but unmapped Base stops with an actionable error. When Base is absent, use KPI sheets with an explicit warning. |
| `base` | Require a mapped Base; no KPI fallback. |
| `kpi` | Explicitly use the existing KPI reader. |

Configured Base aggregation uses fields, dates, statuses, amount-band boundaries and population rules. It does **not** execute arbitrary Excel formula strings. Unknown headers/statuses or invalid mapped data stop extraction rather than silently switching to cached KPI results. Translate the workbook's formulas into these explicit rules, then reconcile an approved quarter before enabling a profile.

For Austria, the supplied layout contains invalid decision-date strings, decisions preceding origins, and repeated contract keys with distinct records. Its configured `snapshot_status` and warning policies preserve the submitted snapshot's row-count convention and recorded status; they do not establish when an invalid historical date occurred. Exact duplicate records fail. Nordic missing/invalid decision-date handling is stricter. Future decisions remain pending at the reporting cutoff.

The supplied Nordic data also contains decisions dated before the positive-balance origin. Its explicit `decision_before_origin: "warn"` policy retains the snapshot category and records the anomaly; it does not alter either date or assert that the chronology is valid. Review the warnings before presenting.

Read [BASE_EXTRACTION.md](docs/BASE_EXTRACTION.md) before changing Base rules. KPI-only workbooks still need recalculation and saving in Excel so their cached results are available.

For a new country, `python inspect_workbook.py --input inputs/Q3_2026/FRANCE.xlsx --output build/france_formula_inventory.json` exports its configured Base header and KPI formula definitions for review, without exporting contract records. The documentation explains how to adjust its header row and sheet name.

## What is the average monthly new flow?

The mean is **the six rolling monthly incoming amounts added together, divided by 6**. Q3 covers April–September; Q1 covers October–March across the year boundary. A missing month does not become zero or disappear from the denominator. The `Flow averages` sheet makes the window and exact inputs visible.

Example: the approved France Q3 flows total 9,444,535 EUR. `9,444,535 / 6 = 1,574,089.17 EUR`, rounded to `1.6m EUR` in the country comment. Incoming flow includes false positives when they are part of the country's submitted incoming population. Processed flow and remaining stock are different measures.

Brazil's calendar IN is one editable source feeding its average and chart data, so edits cannot leave two different monthly averages. The supplementary `UK ZA Amount bands` sheet shows available amounts in GBP and ZAR. Where source thresholds are labelled EUR or otherwise ambiguous, the source labels and caveats remain; local amounts do not prove that the thresholds themselves are in local currency.

Average balance per contract is `sum(incoming amounts) / sum(incoming contracts)` over those same months, not the unweighted mean of monthly averages. Treatment efficiency uses the corresponding processed/incoming amount or count. The quarter comparison of six-month efficiency uses the **oldest cohort (M-5)**, with combined numerators/denominators for Germany + Austria.

## Starting history

`database/history.json` is the starting historical store supplied with this project. It includes the available reporting history through **Q3 2026**, including the latest France update and Romania's confirmed zero Q3 closing stock. It is reporting data, not a set of empty examples. Earlier history can be incomplete where only published summary figures were available.

Keep this file when running the next quarter. Q4 uses Q3 as its comparison; Q1 2027 uses Q4. Running only step 4 can regenerate a report from history without the original country workbooks.

History retains each quarter's own exchange rates. For a report's quarter-on-quarter stock comparison, **both periods are converted using the current reporting quarter's closing rates**. The original stored history is not revalued or overwritten by report generation.

Every changed save records a period revision number, UTC timestamp, previous/new snapshot hashes and relative backup path in the store's `revisions` list. The prior full store is saved under `database/backups/`. Identical normalized reruns create neither a duplicate quarter nor a new revision. Legacy opening history is treated as its first revision when corrected.

Extraction also stores `configuration_sha256` and saves the exact extraction configuration under `work/<period>/config_snapshots/<hash>.json`. Keep these files with the source workbooks to reproduce the extraction rules. Changing configuration does not silently recalculate historical snapshots.

An older-quarter change requires `--replace-historical`; later snapshots remain untouched. Regenerate later comparison reports when their predecessor changes. To restore a backup, stop running jobs, preserve a copy of the current database, then copy the chosen backup over `database/history.json`. This restores the **whole store**, not one country. Regenerate the report afterward.

## Configuration

| File | Change it when… |
|---|---|
| `config/countries.json` | A file alias, country currency, sheet prefix or country-specific stock rule changes. |
| `config/layouts.json` | A source worksheet's layout, labels or exchange-rate columns change. |
| `config/base.json` | Base/KPI routing, verified Base profiles, header aliases, statuses, population rules and band thresholds change. |
| `config/presentation.json` | You want to change country order, CORE/NON CORE grouping, colors or presentation options. |
| `config/overrides/<period>.json` | You have an explicit correction or confirmation that is not in a source workbook. |

Use the supplied override files as the schema example. Overrides belong to one reporting period. Romania's zero-stock confirmation for Q3 2026 must **not** be carried into Q4 automatically.

Rates mean **local-currency units for 1 EUR**: divide a local amount by the rate.
The reader checks the closing-rate (`TC`) table and its month. If a rate file has
dated worksheet names such as `2026.09`, it selects the reporting quarter's
closing-month tab and reads currency codes in column B and closing rates in F.
No file rename or date override is needed for this layout. If a rate file has
no date, keep its period in the filename (for example `taux_q42026.xlsx`) or set
`fx.date_override` in `countries.json` to the verified quarter-end date. Do not
reuse the Q3 table for Q4. Country `units` can be set to `1000` for a confirmed
input in thousands; this never scales contract counts.

For a separately received Nordic CST account balance, create a period-specific
override such as `config/overrides/2026Q4.json`:

```json
{"period":"2026Q4","countries":{"DK":{"cst":{"amount":12345,"currency":"DKK"},"reason":"Closing CST balance confirmed separately"}}}
```

That number is illustrative. Enter the confirmed balance in local currency.
Only `stock_zero` and `cst` overrides are supported; other corrections belong in
the submitted workbook or in the documented extraction mapping.

CORE: Spain, Germany, Austria, Belgium, Italy, Portugal, France, UK and South Africa. Germany and Austria are combined in consolidated tables and retain separate country pages. NON CORE: Romania, Brazil, Denmark, Norway and Sweden.

## Operating rules

- **Missing country:** a complete run stops for review. Use the documented `--allow-partial` option only when you intentionally want a partial report. Missing countries are not silently copied from the previous quarter.
- **Blank versus zero:** blank monetary inputs remain unavailable. A confirmed zero is stored as zero. Country pages can display zero for unavailable detail, with a note; missing data is not silently included as zero in totals or efficiency rates.
- **Excel formulas:** the KPI reader uses saved cached results; recalculate and save those workbooks in Excel. The Base adapter calculates its mapped aggregations in Python, without depending on KPI summary caches. Neither route executes arbitrary Excel formulas.
- **Rounding and inconsistencies:** reported totals are preserved. Validation findings are written during extraction; the editable workbook also includes formula-linked reconciliation checks. A passed technical check does not resolve an inconsistency in country-submitted figures.
- **Stock definition:** total identified stock adds the configured category back to pending stock. For most standard countries that is unreachable-customer balances; for Belgium it is public-authority refunds. Do not add every refund category indiscriminately.
- **Brazil:** IN and OUT are calendar movements. They do not establish treatment efficiency for the same monthly cohort. The chart uses orange IN and green OUT, with clearly labelled scales.
- **Nordics:** separate CST account balances are not inferred from CSR transfer movements.

## Editing values directly in the report

Open **Edit guide** in the generated workbook. Filter the input tab by period and country, then change yellow cells:

| Tab | What to edit | What updates |
|---|---|---|
| `Stock` | KPI_2 pending stock and processed balances still held; TOTAL and vintage rows | Global/CORE/NON CORE stock, stock charts, country numerical comments |
| `Flux` | Monthly incoming/processed amounts, contracts and processing categories | Monthly tables, efficiencies, country charts, numerical comments |
| `Repartition stock` | KPI_3 amount bands, vintages and reported totals | Country stock tables, percentages and contract counts |
| `Taux` | Current-quarter closing exchange rates | Both quarters' EUR comparisons |
| `Other inputs` | Nordic CST and Brazil calendar IN/OUT | CST comparisons and Brazil chart |
| `Comments` | Manually written presentation notes | The corresponding presentation comments |

Green numbers are formulas. Blank inputs are unknown, not zero. Formula cells and keys are protected without a password; filtering is allowed. Do not insert/delete/sort input rows. For new periods, countries or categories, regenerate with Python.

Reported totals sometimes differ from detail. Yellow TOTAL cells remain independent to preserve those supplied figures. When correcting detail, update its total too if necessary and review **Checks**. KPI_2 drives consolidated stock; KPI_3 describes a separately supplied distribution and never silently replaces KPI_2.

Calculation is automatic in Excel. If your Excel session is in Manual mode, press F9. Numerical captions are rounded for slides; manual notes remain your text. The two CORE stock charts share a fixed scale; after a large change, adjust both chart maxima together or regenerate the report.

**Excel edits are saved only in that workbook.** They do not update country files or `database/history.json`. Save an adjusted workbook under a new name. To retain a correction in future Python runs, correct the original country workbook and rerun the pipeline for that quarter. For a correction older than the latest quarter, use `--replace-historical`. Do not regenerate over your only edited copy.

Editable inputs are enabled by default. Set `"editable_inputs": false` in `config/presentation.json` only if you need the original generated-snapshot variant.

## Running options and tests

Useful options:

```powershell
# Extract and validate without updating history or creating a report:
python run_reporting.py --input inputs/Q4_2026 --dry-run
# Deliberately generate an incomplete quarter (coverage is reported):
python run_reporting.py --input inputs/Q4_2026 --allow-partial
# Rebuild the included Q3 report from the opening history:
python 04_generate_report.py --period 2026Q3
```

Correcting an older quarter after newer quarters have been saved requires
`--replace-historical` at step 3 or on the full command. A partial rerun replaces
that quarter's snapshot; it does not merge with countries from an older run.
Use `--database`, `--work-dir` and `--output-dir` to select separate paths for a trial run.

For a safe Q3 extraction test, copy the historical store to a test location and pass that copy using `--database`; do not delete or comment out saved quarters. `--dry-run` changes only the work files. A new Q4 folder needs Q4 closing rates; Q1_2027 automatically crosses the year boundary and compares with Q4_2026.

```powershell
python -m unittest discover -s tests -v
```

All processing is local. Back up the project data before replacing source inputs or changing configuration. The historical backup written by the program protects the previous saved store; it is not a substitute for your normal file backup.
