# Opening history and limits

The supplied JSON store is a starting point for the next reporting cycle:

- Q1 2026: available published summary history. It is not a full set of country KPI submissions.
- Q2 2026: consolidated country reporting with each quarter's original exchange rates.
- Q3 2026: all 14 countries, including the revised France data, Germany and Austria, and Romania's confirmed zero closing stock.

Q3 is the supplied example report and the comparison baseline for Q4. The store
is organized by reporting quarter; overlapping monthly cohorts in successive
quarters are kept separately because their treatment results evolve.

Country-submitted rounding differences and mismatches between KPI2 and KPI3
remain in the history. The supplementary UK/South Africa amount-band table
displays amounts in GBP and ZAR respectively. Source thresholds labelled EUR
or otherwise ambiguous retain their source labels and caveats; a local-currency
amount does not establish the currency of its band thresholds. Missing standalone
CST balances are not inferred from transfers; add confirmed balances through
period overrides.

The starting history was assembled from the earlier reporting material. New
extractions additionally retain the available KPI1 observations and file hashes.
It does not contain original source workbooks, contract-level records, or
invented data for fields absent from earlier reporting. Keep your source files
alongside this project in your normal document storage.

Recreating earlier reports with fair constant-FX comparisons requires the
previous quarter's currency conversion evidence. Q3 versus Q2 was verified;
older partial summary history may need completion before being regenerated.

The opening store can predate revision metadata. Its current observations are
treated as initial revisions; subsequent changes create numbered revision
records and full-store backups. New Base extractions have their own verified
profile and configuration hash. They do not retrospectively recompute every
historical snapshot just because the adapter or configuration has changed.
