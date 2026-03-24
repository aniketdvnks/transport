# Hazchem Transport Import Set

Report date used for operational seed: `2026-03-24`

Import order:
1. `customers_hazchem_import_v15`
2. `drivers_hazchem_import_v15`
3. `trailers_hazchem_import_v15`
4. `vehicles_hazchem_import_v15`
5. `service_items_hazchem_import_v15`
6. `transport_locations_hazchem_import_v15`
7. `trip_locations_hazchem_import_v15`
8. `transport_cargo_types_hazchem_import_v15`
9. `trip_routes_hazchem_import_v15`
10. `transportation_orders_hazchem_import_v15`
11. `vehicle_trips_hazchem_import_v15`
12. `sales_invoices_hazchem_import_v15`

Seeded record counts:
- Customers: 3
- Drivers: 8
- Trailers: 8
- Vehicles: 8
- Service Items: 4
- Transport Locations: 10
- Trip Locations: 10
- Cargo Types: 4
- Trip Routes: 8
- Transportation Orders: 8
- Vehicle Trips: 8
- Sales Invoices: 8

Notes:
- Tanker is represented by the `Vehicle` master.
- The daily tanker board is available in the `Daily Trip Schedule` report.
- Transportation orders include child-table exports for cargo and assignment rows so the production import preserves dispatch structure.
- Vehicle trip exports include `main_route_steps`, which the tracking report uses for arrival and departure milestones.
