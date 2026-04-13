# Quoting App V21

Stable export-fixed build based on your uploaded app.py.

Includes:
- row-based sheet settings
- per-sheet bundle export
- price writing back into Excel
- DS loading
- SQM tier markups
- history page


SQM markup rules are loaded from `data/markup_rules.json` and hidden from the UI.


Double-sided loading is now loaded from `data/ds_rules.json` and hidden from the UI.


Adds a Material Cost Summary to the on-screen totals and the exported Quote Summary sheet.


V26 final: SQM markup factor is applied using `total_sqm` in both UI calculations and Excel export recalculation.
