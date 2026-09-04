# Report tools

Desktop Python tools for turning website analytics CSVs into branded spreadsheets, Word/PDF reports, cover pages, and work logs.

## Scripts

- `scripts/clarity_csv_to_spreadsheet.py` — Microsoft Clarity CSV to Excel
- `scripts/spreadsheet_to_report.py` — spreadsheet to branded report
- `scripts/cover_page.py` — cover page generator
- `scripts/worklog.py` — work log generator

Profiles live in `profiles/`. Generated files go to `reports/processed/`; place source CSVs and similar inputs in `reports/raw_files/`. Those folders are gitignored so local reports are not committed.

## License

MIT. See [LICENSE](LICENSE).
