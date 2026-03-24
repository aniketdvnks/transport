# ERPNext v15 Trip Route Import Outputs

Source file: `/Users/aniketshinde/Downloads/ROUTE EXCEL.xlsx`
Generated on: `2026-03-24`

Matched to the transport route doctypes in this checkout of the transport app:
- `Trip Route`
- `Trip Steps Table`
- `Trip Location`
- `Trip Location Type`

Import order:
1. `trip_location_type_seed_v15.xlsx` or `.csv`
2. `trip_location_import_v15.xlsx` or `.csv`
3. `trip_route_import_v15.xlsx` or `.csv`

Counts:
- Source rows read: 2612
- Routes generated: 2605
- Trip Route import rows: 5210
- Trip Location rows: 1476
- Exact duplicates skipped: 5
- Incomplete rows skipped: 2
- Same-origin/destination routes flagged: 17
- Review rows: 25

Important assumptions:
- Each source row was converted into a two-step route.
- Step 1 is `Loading Point` using `RouteFrom`.
- Step 2 is `Offloading Point` using `RouteTo`.
- `Total Distance (km)` was set to `0` for every route because the source file does not provide distance data.
- Child-step distances and fuel consumption quantities were left blank for the same reason.

Use `route_source_map_v15` to trace each generated route back to the original spreadsheet row.
