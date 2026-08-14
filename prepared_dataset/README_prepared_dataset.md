# Prepared ML Dataset

Source folder:
- C:\Users\User\Documents\G3\Data set

Outputs:
- indicator_reference.csv: indicator list parsed from reference file column A
- ml_entities.csv: one row per center per year
- ml_dataset_long.csv: one row per center per year per indicator
- ml_dataset_wide.csv: one row per center per year with 0/1 indicator columns
- unknown_improvement_codes.csv: improvement codes found in annual files but not in the reference list
- dataset_summary.csv: imported center counts by year

Label meaning:
- needs_improvement = 1 means the indicator code appears in the source improvement-code field
- needs_improvement = 0 means the indicator code does not appear in that field

Assumptions:
- indicator reference comes from column A of the REP006 file
- rows without user_code or center_name are skipped
