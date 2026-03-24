# ERPNext v15 Party Import Outputs

Source file: `/Users/aniketshinde/Downloads/PARTY NAME.xlsx`
Generated on: `2026-03-24`

Import order:
1. `customer_import_v15.xlsx` or `.csv`
2. `supplier_import_v15.xlsx` or `.csv`
3. `driver_import_v15.xlsx` or `.csv`
4. `address_import_v15.xlsx` or `.csv`

Counts:
- Customers: 482
- Suppliers: 969
- Drivers: 836
- Addresses: 1063
- Unmapped review rows: 303

Mapping assumptions:
- `CLIENT` rows were mapped to `Customer`.
- `Supplier`, `Tyre Supplier`, `PUMP`, `Tax Vendor`, `OWNER`, `Broker`, `Agent`, `Financier`, and `Bank` rows were mapped to `Supplier`.
- `Driver Ledger` and `Debit Driver` rows were mapped to `Driver`.
- Explicit `name` values were assigned so the address file can link to the imported masters reliably.
- Address rows were generated only when the source row had at least one address line or a GST number.
- Remaining roles were kept in `unmapped_review_v15` for manual review because they look like ledger heads, vehicles, or mixed-purpose records.
