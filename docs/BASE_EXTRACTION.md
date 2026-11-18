# Adding or changing a Base-sheet profile

The Base route replaces Excel aggregation with explicit Python aggregation. It does not infer business meaning from the word `Base`, and it does not execute arbitrary formula strings. Current verified profiles cover the supplied Nordic and Austrian layouts. Other Base-capable countries require verified mappings before a present Base sheet can be used.

## Routing

`config/base.json` contains per-country routing and reusable profiles:

| Mode | Meaning |
|---|---|
| `base_if_present` | Mapped Base present: aggregate it. Base absent: explicit KPI fallback warning. Base present but unmapped: stop for configuration. |
| `base` | Require Base and a verified profile. |
| `kpi` | Intentionally read the KPI summaries. |

An invalid mapped Base never silently falls back to cached KPI figures. Header/status/data errors need review. Do not assign the Nordic profile to another country just because its columns look similar.

## Translate the workbook formulas into business rules

For an approved quarter, inspect the formulas producing each KPI and record:

1. **Fields:** account/contract identifier, amount, positive-balance origin date, status and treatment dates. Match required columns using the profile's header aliases.
2. **Population:** which rows enter gross incoming flow, pending stock and KPI_3. Closed, refunded, unreachable and false-positive records can have different treatment.
3. **Timing:** reporting cutoff and six monthly cohort windows. Origin month and processing month are not interchangeable. A future decision remains pending at the cutoff.
4. **Statuses:** exact values mapped to client refund, unreachable, false positive, public authority or other processing. Unknown values must be reviewed.
5. **Units:** currency and whether amounts are units or thousands. Scaling amounts must never scale contract counts.
6. **Bands/vintages:** thresholds, inclusivity at every boundary, source currency, reporting population and older-year grouping.

Profiles use disjoint amount bands with `(lower, upper]` boundaries for both amounts and counts. The supplied Nordic spreadsheet used inconsistent boundary formulas; Python does not copy an overlap or gap into its model. Nordic band thresholds use EUR-equivalent limits converted with the workbook Parameters rate for the source population. Consolidated report EUR conversion still uses the current `Taux` file; these serve different purposes.

For Austria, `Refund_C` maps to client refund, `Refund_A` to retained/unreachable balances, and `No refund` to false positives. Its configured missing/invalid-decision-date policy is `snapshot_status`, with warnings. The inspected file has 165 invalid date strings, 170 decisions preceding balance origins and 584 repeated contract/origin keys with distinct records. Its row-count convention is preserved; exact duplicate records fail. Snapshot status does not reconstruct when an undated treatment occurred. Nordic profiles use the stricter missing-date policy. Review policy changes as business decisions, not just parser adjustments.

`kpi3_exclusions` applies only to the KPI_3 retained population. It does not remove incoming amounts from KPI_1 or KPI_2.

To collect a new country's schema and formulas without exporting its contract rows:

```powershell
python inspect_workbook.py --input inputs/Q3_2026/FRANCE.xlsx --output build/france_formula_inventory.json
```

The default Base header is row 4; set `--base-header-row` and `--base-sheet` when the workbook differs. The utility exports that header row, worksheet names and KPI formula definitions (first 200 rows / 80 columns). It does not execute formulas, export Base records or update history. Use the inventory to review a profile; it does not automatically approve one.

The Nordic profile explicitly sets `decision_before_origin: "warn"` because the supplied data contains such anomalies. It retains the snapshot status and counts the anomaly in `base_audit`, with a validation warning. Dates are not repaired or replaced. Other profiles keep the default error policy unless reviewed configuration explicitly changes it.

These explicit profile rules are the configurable replacement for the Excel formulas. A more complicated rule that the existing profile cannot express needs a reviewed adapter; never add unrestricted `eval` of user-supplied formulas.

## Reconcile before enabling

Run `python run_reporting.py --input inputs/QX_20XX --dry-run` with a real quarter folder. Compare the derived results with approved KPI figures for each cohort, processing category, stock vintage and amount band. Confirm both amounts and counts, units, cutoff and currency.

Do not overwrite Base results with KPI caches just to make the checks agree. A stale KPI cache, a different population, rounding and a wrong mapping can all cause differences. Keep the reconciliation findings visible and resolve the cause.

Add synthetic tests with known results: month/year boundaries, exact band edges, zero and missing values, each supported status, unsupported statuses, future decisions and the intended duplicate-ID rule. Keep personal contract data out of fixtures and logs.

After approval, enable the profile for the country and rerun extraction, validation and history. The extraction stores a configuration hash and an immutable configuration copy in the work folder. Corrections create a numbered history revision and backup. Updating a profile alone does not change historical snapshots.
