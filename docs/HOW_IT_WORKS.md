# How the automation works

The flow is:

```text
Quarter folder + configuration
          |
          v
1. Extract -> 2. Validate -> 3. Update history -> 4. Generate Excel
                                  ^                     |
                                  |                     v
                         Existing history          Detailed report
```

## 1. Extract the submitted figures

The folder name identifies the reporting period: `Q4_2026` becomes `2026Q4`. Country configuration tells the reader which files and worksheets belong to each country. The reader stores numbers under consistent names, even when their worksheet locations differ.

For example, France's `KPI_2` may report:

```text
Expected amount to be processed:      2,145,041 EUR
Customer not reachable:                 738,601 EUR
```

The extracted record keeps these two components separately. It does not replace the original workbook or convert everything into EUR yet. Monthly cohorts, amount bands, vintage distributions and reported totals are also preserved. `KPI_1` is used to check country/period context; the presentation mainly uses `KPI_2` and `KPI_3`.

The reader first follows `config/base.json`. A verified Base profile aggregates raw rows directly in Python; otherwise a configured KPI route reads summary cells. Current verified Base layouts cover Austria and the supplied Nordic workbook. A present but unconfigured Base stops extraction, while a missing Base can use the explicitly configured KPI fallback with a warning. Brazil's profile reads its movement and aging report. Romania can supply stock-only data.

The extracted snapshot contains a `configuration_sha256`. The exact extraction configuration is retained under `work/<period>/config_snapshots/<hash>.json`, so later mapping changes can be distinguished from the rules used for this run. See [BASE_EXTRACTION.md](BASE_EXTRACTION.md) for mapping and reconciliation guidance.

## 2. Check and approve the data for processing

Validation checks that a Q4 run contains Q4 data, that required currencies have rates, and that countries and records are not duplicated. It also compares reported totals with their components and checks missing fields.

A quarter-specific override is applied only when you explicitly record a confirmed fact. The supplied Q3 Romania override says the closing stock is zero. It does **not** say that Romania had zero inflows during the quarter, or 100% treatment efficiency.

Example of a rounding finding:

```text
France amount bands add to:        2,883,641 EUR
France reported stock total:       2,883,642 EUR
Difference:                              1 EUR
```

The program retains the reported total and records the difference. It does not change an individual band to make the sum look exact. Errors that prevent a reliable calculation stop the run; warnings remain available for review outside the presentation workbook.

## 3. Add the quarter to history

The database is a JSON file organised by reporting period and country. JSON keeps the structure explicit and is easy to inspect, back up and read with Python. It is separate from the presentation Excel.

Before Q4:

```text
2026Q2 -> countries, stock, cohorts, distributions, Q2 rates
2026Q3 -> countries, stock, cohorts, distributions, Q3 rates
```

After Q4:

```text
2026Q2 -> unchanged
2026Q3 -> unchanged
2026Q4 -> new validated snapshot
```

Running Q4 again with the same validated snapshot leaves history unchanged. Changed Q4 data replaces Q4's snapshot and increments its revision number; it does not create two Q4s. Before the change, a full-store backup is saved. The `revisions` list records quarter, revision, UTC save time, previous/new snapshot hashes and a relative backup path. A failed validation must not update history.

Correcting Q3 after Q4 exists requires `--replace-historical`; Q4's saved snapshot remains unchanged. Regenerate the Q4 presentation to include the corrected Q3 comparison. Q1 2027 follows the same process and compares against Q4 2026 automatically.

The same cohort month can appear in successive quarters. For example, September's treatment status in Q3 is different from September's status in Q4. Both observations are kept under their respective reporting dates; they are not duplicate records.

## 4. Create the detailed Excel

The report generator reads the chosen quarter and its immediate predecessor. It derives the country pages and the global/CORE/NON CORE tables from the same figures.

Example stock calculation:

```text
France pending stock                    2,145,041
+ unreachable balances                    738,601
= total identified stock                2,883,642 EUR
```

The stock-to-process chart uses the first amount. The total-stock table uses the sum. Amount-band and vintage tables retain their submitted scope; a source distribution need not equal the pending-stock measure.

Example constant-FX comparison using illustrative figures:

```text
Previous-quarter stock:                  100 GBP
Current-quarter stock:                   110 GBP
Current closing rate:                  0.80 GBP per EUR

Previous stock at current FX:        100 / 0.80 = 125.00 EUR
Current stock at current FX:         110 / 0.80 = 137.50 EUR
Change:                         137.50 / 125 - 1 = +10.0%
```

Both sides use the same rate. The saved previous-quarter FX remains unchanged. Contract counts are never currency-converted.

For Q4 the monthly charts show July–December. In Q1 2027 they show October 2026–March 2027. The oldest cohort's efficiency is treated amount divided by incoming amount, or treated count divided by incoming count. Missing or zero denominators do not become an invented efficiency rate.

The average monthly **new** flow is the sum of the six incoming monthly amounts divided by 6. It does not use treated flow or stock. France Q3 provides a simple example: 9,444,535 EUR / 6 = 1,574,089.17 EUR, rounded to 1.6m EUR in the slide comment. A missing month keeps the mean unavailable. The `Flow averages` sheet exposes the period window and exact inputs. Average amount per contract uses total incoming amount divided by total incoming count, not a mean of monthly ratios.

Brazil's OUT series may include balances originating before the displayed month, so its calendar IN/OUT ratio is not a cohort-efficiency measure. A country with only confirmed closing stock participates in stock totals but not in monthly-flow totals without monthly data.

## What to review before presenting

Open the validation output, check country coverage and confirm any warnings against the country submissions. Then inspect the report's comparison quarter, exchange rates and material stock movements. Comments describing operational decisions or explanations still need human review; the program calculates figures, not business reasons.

For an adjustment to the current presentation, use the yellow cells in the generated workbook's input tabs. Example: filter `Stock` to France / 2026Q3, then change the TOTAL row's pending stock. The total-stock summary and rounded comments recalculate. Change the corresponding vintage row as well if the correction affects its stacked chart. Review `Checks` for differences between the total and detail. The `Edit guide` tab explains each input table.

These Excel adjustments stay in that workbook. For a permanent correction that future quarters must use, update the country input workbook (or a supported explicit override), then rerun the four steps for the affected quarter. Regeneration does not import changes from the presentation workbook back into history.
