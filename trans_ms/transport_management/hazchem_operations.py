from __future__ import unicode_literals

import csv
from pathlib import Path
from urllib.parse import quote

import frappe
from frappe.core.doctype.data_import.exporter import Exporter
from frappe.utils import add_years, getdate, today
from openpyxl import Workbook

WORKSPACE_NAME = "Transport"
WORKSPACE_FLOW_BLOCK_NAME = "Transport Hazchem Flow Guide"
WORKSPACE_TIMELINE_BLOCK_NAME = "Transport Process Timeline"
WORKSPACE_CONTENT = (
    '[{"type":"header","data":{"text":"<span class=\\"h4\\">Transport</span>","col":12}},'
    '{"type":"custom_block","data":{"custom_block_name":"Transport Hazchem Flow Guide","col":12}},'
    '{"type":"custom_block","data":{"custom_block_name":"Transport Process Timeline","col":12}},'
    '{"type":"header","data":{"text":"<span class=\\"h4\\"><b>Report &amp; Masters</b></span>","col":12}},'
    '{"type":"card","data":{"card_name":"Document","col":3}},'
    '{"type":"card","data":{"card_name":"Trip","col":3}},'
    '{"type":"card","data":{"card_name":"Setup","col":3}},'
    '{"type":"card","data":{"card_name":"Setting","col":3}},'
    '{"type":"card","data":{"card_name":"Reports","col":4}}]'
)
DEFAULT_OUTPUT_DIR = (
    Path(__file__).resolve().parents[2]
    / "output"
    / "spreadsheet"
    / "hazchem_transport_operational_imports"
)

CUSTOMERS = [
    {
        "customer_name": "Apex Chlorochem Pvt Ltd",
        "customer_group": "Commercial",
        "territory": "India",
        "customer_type": "Company",
        "default_currency": "INR",
    },
    {
        "customer_name": "Nova Specialty Chemicals Ltd",
        "customer_group": "Commercial",
        "territory": "India",
        "customer_type": "Company",
        "default_currency": "INR",
    },
    {
        "customer_name": "Shakti Industrial Gases Pvt Ltd",
        "customer_group": "Commercial",
        "territory": "India",
        "customer_type": "Company",
        "default_currency": "INR",
    },
]

CARGO_TYPES = [
    "Liquid Caustic Soda",
    "Hydrochloric Acid",
    "Sulphuric Acid",
    "Sodium Hypochlorite",
]

SERVICE_ITEMS = {
    "Liquid Caustic Soda": {
        "item_code": "HAZ-LCS-TNK",
        "item_name": "Hazchem Tanker Service - Liquid Caustic Soda",
        "description": "Dedicated tanker haulage service for liquid caustic soda movements.",
    },
    "Hydrochloric Acid": {
        "item_code": "HAZ-HCL-TNK",
        "item_name": "Hazchem Tanker Service - Hydrochloric Acid",
        "description": "Dedicated tanker haulage service for hydrochloric acid movements.",
    },
    "Sulphuric Acid": {
        "item_code": "HAZ-H2SO4-TNK",
        "item_name": "Hazchem Tanker Service - Sulphuric Acid",
        "description": "Dedicated tanker haulage service for sulphuric acid corridors.",
    },
    "Sodium Hypochlorite": {
        "item_code": "HAZ-NAOCL-TNK",
        "item_name": "Hazchem Tanker Service - Sodium Hypochlorite",
        "description": "Dedicated tanker haulage service for sodium hypochlorite movements.",
    },
}

SERVICE_RATE_PER_KM = {
    "Liquid Caustic Soda": 135,
    "Hydrochloric Acid": 145,
    "Sulphuric Acid": 155,
    "Sodium Hypochlorite": 140,
}

TRANSPORT_LOCATIONS = [
    "Dahej",
    "Taloja",
    "Jhagadia",
    "Pithampur",
    "Vadodara",
    "Panipat",
    "Hazira",
    "Sanand",
    "Bharuch",
    "Nagothane",
]

TRIP_LOCATIONS = [
    "Apex Chlorochem Dahej Plant",
    "Nova Taloja Tank Farm",
    "Jhagadia Industrial Estate",
    "Pithampur Chemical Hub",
    "Vadodara Chlor Alkali Depot",
    "Panipat Process Plant",
    "Hazira Bulk Terminal",
    "Sanand Specialty Park",
    "Nagothane Chemical Complex",
    "Bharuch Industrial Tank Farm",
]

TRAILERS = [
    {"number_plate": "GJ06TT9101", "chassis_number": "SATCHEM9101", "make": "Satyam Trailers", "year": 2021, "axles": 3},
    {"number_plate": "GJ06TT9107", "chassis_number": "SATCHEM9107", "make": "Satyam Trailers", "year": 2021, "axles": 3},
    {"number_plate": "GJ06TT9112", "chassis_number": "SATCHEM9112", "make": "Satyam Trailers", "year": 2022, "axles": 3},
    {"number_plate": "GJ27TT2205", "chassis_number": "SATCHEM2205", "make": "Satyam Trailers", "year": 2020, "axles": 3},
    {"number_plate": "MH04TT7821", "chassis_number": "SATCHEM7821", "make": "Satyam Trailers", "year": 2019, "axles": 3},
    {"number_plate": "RJ14TT3342", "chassis_number": "SATCHEM3342", "make": "Satyam Trailers", "year": 2022, "axles": 3},
    {"number_plate": "GJ06TT4124", "chassis_number": "SATCHEM4124", "make": "Satyam Trailers", "year": 2023, "axles": 3},
    {"number_plate": "GJ06TT4131", "chassis_number": "SATCHEM4131", "make": "Satyam Trailers", "year": 2023, "axles": 3},
]

DRIVERS = [
    {"full_name": "Harshad Patel", "cell_number": "9876543101", "license_number": "GJ-HC-483921"},
    {"full_name": "Imran Shaikh", "cell_number": "9876543107", "license_number": "GJ-HC-483927"},
    {"full_name": "Rakesh Solanki", "cell_number": "9876543112", "license_number": "MP-HC-483932"},
    {"full_name": "Devendra Rathod", "cell_number": "9876543205", "license_number": "HR-HC-483945"},
    {"full_name": "Nilesh Parmar", "cell_number": "9876543821", "license_number": "MH-HC-483982"},
    {"full_name": "Ajay Yadav", "cell_number": "9876543342", "license_number": "RJ-HC-483952"},
    {"full_name": "Sandeep Chouhan", "cell_number": "9876544124", "license_number": "GJ-HC-484024"},
    {"full_name": "Mahesh Prajapati", "cell_number": "9876544131", "license_number": "GJ-HC-484031"},
]

VEHICLES = [
    {"license_plate": "GJ06TX4101", "make": "BharatBenz", "model": "24 KL Chemical Tanker", "last_odometer": 186420, "location": "Dahej Fleet Yard", "trailer": "GJ06TT9101", "driver": "Harshad Patel", "fuel_consumption": 0.36},
    {"license_plate": "GJ06TX4107", "make": "BharatBenz", "model": "18 KL Chemical Tanker", "last_odometer": 149880, "location": "Dahej Fleet Yard", "trailer": "GJ06TT9107", "driver": "Imran Shaikh", "fuel_consumption": 0.34},
    {"license_plate": "GJ06TX4112", "make": "Ashok Leyland", "model": "24 KL Acid Tanker", "last_odometer": 172640, "location": "Bharuch Yard", "trailer": "GJ06TT9112", "driver": "Rakesh Solanki", "fuel_consumption": 0.36},
    {"license_plate": "GJ27TX2205", "make": "Tata", "model": "32 KL Bulk Chemical Tanker", "last_odometer": 240510, "location": "Hazira Terminal", "trailer": "GJ27TT2205", "driver": "Devendra Rathod", "fuel_consumption": 0.35},
    {"license_plate": "MH04LM7821", "make": "Mahindra", "model": "16 KL Caustic Tanker", "last_odometer": 131760, "location": "Jhagadia Yard", "trailer": "MH04TT7821", "driver": "Nilesh Parmar", "fuel_consumption": 0.37},
    {"license_plate": "RJ14TX3342", "make": "BharatBenz", "model": "24 KL Stainless Tanker", "last_odometer": 164920, "location": "Nagothane Yard", "trailer": "RJ14TT3342", "driver": "Ajay Yadav", "fuel_consumption": 0.36},
    {"license_plate": "GJ06TX4124", "make": "Tata", "model": "32 KL Chemical Tanker", "last_odometer": 201345, "location": "Dahej Fleet Yard", "trailer": "GJ06TT4124", "driver": "Sandeep Chouhan", "fuel_consumption": 0.35},
    {"license_plate": "GJ06TX4131", "make": "Ashok Leyland", "model": "20 KL Hypo Tanker", "last_odometer": 154605, "location": "Bharuch Yard", "trailer": "GJ06TT4131", "driver": "Mahesh Prajapati", "fuel_consumption": 0.37},
]

OPERATIONS = [
    {
        "dispatch_ref": "HAZCHEM-OPS-001",
        "customer": "Apex Chlorochem Pvt Ltd",
        "cargo_type": "Liquid Caustic Soda",
        "route_name": "Dahej Plant to Taloja Tank Farm",
        "loading_city": "Dahej",
        "destination_city": "Taloja",
        "loading_point": "Apex Chlorochem Dahej Plant",
        "offloading_point": "Nova Taloja Tank Farm",
        "distance": 580,
        "planned_fuel": 210,
        "fuel_issued": 220,
        "vehicle": "GJ06TX4101",
        "trailer": "GJ06TT9101",
        "driver": "Harshad Patel",
        "stage": "LO",
        "container_number": "BULK-HC-001",
        "net_weight": 24.0,
        "tare_weight": 8.4,
        "extra_details": "UN 1824, Class 8 caustic soda lye, sealed dome loading, PPE mandatory.",
    },
    {
        "dispatch_ref": "HAZCHEM-OPS-002",
        "customer": "Apex Chlorochem Pvt Ltd",
        "cargo_type": "Hydrochloric Acid",
        "route_name": "Dahej Plant to Jhagadia Estate",
        "loading_city": "Dahej",
        "destination_city": "Jhagadia",
        "loading_point": "Apex Chlorochem Dahej Plant",
        "offloading_point": "Jhagadia Industrial Estate",
        "distance": 90,
        "planned_fuel": 34,
        "fuel_issued": 28,
        "vehicle": "GJ06TX4107",
        "trailer": "GJ06TT9107",
        "driver": "Imran Shaikh",
        "stage": "EM",
        "container_number": "BULK-HC-002",
        "net_weight": 18.5,
        "tare_weight": 7.8,
        "extra_details": "UN 1789, Class 8 hydrochloric acid for captive process dosing, bottom discharge only.",
    },
    {
        "dispatch_ref": "HAZCHEM-OPS-003",
        "customer": "Nova Specialty Chemicals Ltd",
        "cargo_type": "Sulphuric Acid",
        "route_name": "Bharuch Tank Farm to Pithampur Hub",
        "loading_city": "Bharuch",
        "destination_city": "Pithampur",
        "loading_point": "Bharuch Industrial Tank Farm",
        "offloading_point": "Pithampur Chemical Hub",
        "distance": 525,
        "planned_fuel": 190,
        "fuel_issued": 180,
        "vehicle": "GJ06TX4112",
        "trailer": "GJ06TT9112",
        "driver": "Rakesh Solanki",
        "stage": "WORKING",
        "container_number": "BULK-HC-003",
        "net_weight": 23.8,
        "tare_weight": 8.6,
        "extra_details": "UN 1830, Class 8 sulphuric acid, route restricted to approved hazchem corridor.",
    },
    {
        "dispatch_ref": "HAZCHEM-OPS-004",
        "customer": "Shakti Industrial Gases Pvt Ltd",
        "cargo_type": "Sodium Hypochlorite",
        "route_name": "Hazira Terminal to Panipat Plant",
        "loading_city": "Hazira",
        "destination_city": "Panipat",
        "loading_point": "Hazira Bulk Terminal",
        "offloading_point": "Panipat Process Plant",
        "distance": 1210,
        "planned_fuel": 420,
        "fuel_issued": 395,
        "vehicle": "GJ27TX2205",
        "trailer": "GJ27TT2205",
        "driver": "Devendra Rathod",
        "stage": "EM",
        "container_number": "BULK-HC-004",
        "net_weight": 28.0,
        "tare_weight": 9.2,
        "extra_details": "UN 1791 sodium hypochlorite solution, no mixed loading, route monitored every 4 hours.",
    },
    {
        "dispatch_ref": "HAZCHEM-OPS-005",
        "customer": "Nova Specialty Chemicals Ltd",
        "cargo_type": "Liquid Caustic Soda",
        "route_name": "Jhagadia Estate to Vadodara Depot",
        "loading_city": "Jhagadia",
        "destination_city": "Vadodara",
        "loading_point": "Jhagadia Industrial Estate",
        "offloading_point": "Vadodara Chlor Alkali Depot",
        "distance": 72,
        "planned_fuel": 28,
        "fuel_issued": 32,
        "vehicle": "MH04LM7821",
        "trailer": "MH04TT7821",
        "driver": "Nilesh Parmar",
        "stage": "LO",
        "container_number": "BULK-HC-005",
        "net_weight": 16.5,
        "tare_weight": 7.0,
        "extra_details": "Internal plant-to-depot replenishment, controlled unloading window between 14:00 and 18:00.",
    },
    {
        "dispatch_ref": "HAZCHEM-OPS-006",
        "customer": "Apex Chlorochem Pvt Ltd",
        "cargo_type": "Hydrochloric Acid",
        "route_name": "Nagothane Complex to Sanand Park",
        "loading_city": "Nagothane",
        "destination_city": "Sanand",
        "loading_point": "Nagothane Chemical Complex",
        "offloading_point": "Sanand Specialty Park",
        "distance": 640,
        "planned_fuel": 230,
        "fuel_issued": 240,
        "vehicle": "RJ14TX3342",
        "trailer": "RJ14TT3342",
        "driver": "Ajay Yadav",
        "stage": "UNL",
        "container_number": "BULK-HC-006",
        "net_weight": 24.2,
        "tare_weight": 8.3,
        "extra_details": "UN 1789 hydrochloric acid, unloading under site escort, wash certificate required on return.",
    },
    {
        "dispatch_ref": "HAZCHEM-OPS-007",
        "customer": "Shakti Industrial Gases Pvt Ltd",
        "cargo_type": "Sulphuric Acid",
        "route_name": "Dahej Plant to Panipat Plant",
        "loading_city": "Dahej",
        "destination_city": "Panipat",
        "loading_point": "Apex Chlorochem Dahej Plant",
        "offloading_point": "Panipat Process Plant",
        "distance": 1135,
        "planned_fuel": 395,
        "fuel_issued": 365,
        "vehicle": "GJ06TX4124",
        "trailer": "GJ06TT4124",
        "driver": "Sandeep Chouhan",
        "stage": "WORKING",
        "container_number": "BULK-HC-007",
        "net_weight": 30.0,
        "tare_weight": 9.8,
        "extra_details": "Long-haul acid movement, relay checkpoint at Jaipur, ADR kit and spill kit mandatory.",
    },
    {
        "dispatch_ref": "HAZCHEM-OPS-008",
        "customer": "Nova Specialty Chemicals Ltd",
        "cargo_type": "Sodium Hypochlorite",
        "route_name": "Bharuch Tank Farm to Taloja Tank Farm",
        "loading_city": "Bharuch",
        "destination_city": "Taloja",
        "loading_point": "Bharuch Industrial Tank Farm",
        "offloading_point": "Nova Taloja Tank Farm",
        "distance": 430,
        "planned_fuel": 160,
        "fuel_issued": 150,
        "vehicle": "GJ06TX4131",
        "trailer": "GJ06TT4131",
        "driver": "Mahesh Prajapati",
        "stage": "EM",
        "container_number": "BULK-HC-008",
        "net_weight": 20.0,
        "tare_weight": 7.5,
        "extra_details": "UN 1791 sodium hypochlorite, vented discharge, daylight movement only.",
    },
]

OLD_DAILY_SCHEDULE_DEMO = {
    "plates": ["9171", "5472", "7299", "5390", "271", "1334", "6391", "6199", "1470", "6099"],
    "routes": [
        "DAHEJ / HOPL",
        "BHIWADI / MODASA",
        "HOPL / SRF",
        "CHATRAL / BODAL",
        "VAPI / BODAL S",
        "ANKLESHWAR / BODAL",
        "KHAMBHAT / BODAL",
        "BIRLA / MADHU",
    ],
    "locations": ["DAHEJ", "HOPL", "BHIWADI", "MODASA", "SRF", "CHATRAL", "BODAL", "VAPI", "BODAL S", "ANKLESHWAR", "KHAMBHAT", "BIRLA", "MADHU"],
}

EXPORT_CONFIG = {
    "Customer": {
        "fields": {
            "Customer": ["name", "customer_name", "customer_type", "customer_group", "territory", "default_currency"]
        },
        "filename": "customers_hazchem_import_v15",
    },
    "Driver": {
        "fields": {
            "Driver": ["name", "naming_series", "full_name", "status", "cell_number", "license_number", "expiry_date"]
        },
        "filename": "drivers_hazchem_import_v15",
    },
    "Trailer": {
        "fields": {
            "Trailer": ["name", "number_plate", "chassis_number", "make", "year", "axles", "trailer_type", "tyre_specification", "suspension_type"]
        },
        "filename": "trailers_hazchem_import_v15",
    },
    "Vehicle": {
        "fields": {
            "Vehicle": [
                "name",
                "license_plate",
                "make",
                "model",
                "last_odometer",
                "acquisition_date",
                "location",
                "trans_ms_default_trailer",
                "trans_ms_driver",
                "trans_ms_fuel_consumption",
            ]
        },
        "filename": "vehicles_hazchem_import_v15",
    },
    "Item": {
        "fields": {
            "Item": [
                "name",
                "item_code",
                "item_name",
                "item_group",
                "stock_uom",
                "is_stock_item",
                "is_sales_item",
                "description",
                "disabled",
            ]
        },
        "filename": "service_items_hazchem_import_v15",
    },
    "Transport Location": {
        "fields": {"Transport Location": ["name", "location", "country"]},
        "filename": "transport_locations_hazchem_import_v15",
    },
    "Trip Location": {
        "fields": {"Trip Location": ["name", "description", "location", "is_local_border", "is_international_border"]},
        "filename": "trip_locations_hazchem_import_v15",
    },
    "Transport Cargo Type": {
        "fields": {"Transport Cargo Type": ["name", "cargo_name"]},
        "filename": "transport_cargo_types_hazchem_import_v15",
    },
    "Trip Route": {
        "fields": {
            "Trip Route": ["name", "route_name", "total_distance", "total_fuel_consumption_qty"],
            "trip_steps": ["name", "location", "distance", "location_type", "fuel_consumption_qty"],
        },
        "filename": "trip_routes_hazchem_import_v15",
    },
    "Transportation Order": {
        "fields": {
            "Transportation Order": [
                "name",
                "date",
                "customer",
                "loading_date",
                "company",
                "transport_type",
                "cargo_type",
                "goods_description",
                "special_instructions_to_transporter",
                "assignment_status",
                "version",
            ],
            "cargo": [
                "name",
                "container_size",
                "container_number",
                "net_weight",
                "tare_weight",
                "cargo_location_country",
                "cargo_location_city",
                "cargo_destination_country",
                "cargo_destination_city",
                "cargo_type",
                "extra_details",
            ],
            "assign_transport": [
                "name",
                "cargo",
                "amount",
                "expected_loading_date",
                "container_number",
                "units",
                "transporter_type",
                "assigned_vehicle",
                "assigned_trailer",
                "assigned_driver",
                "driver_name",
                "route",
                "status",
                "created_trip",
                "vehicle_plate_number",
                "trailer_plate_number",
                "item",
                "rate",
                "invoice",
            ],
        },
        "filename": "transportation_orders_hazchem_import_v15",
    },
    "Vehicle Trip": {
        "fields": {
            "Vehicle Trip": [
                "name",
                "date",
                "transporter_type",
                "vehicle",
                "vehicle_plate_number",
                "trailer",
                "trailer_plate_number",
                "driver",
                "driver_name",
                "customer",
                "start_date",
                "main_cargo_category",
                "main_loading_point",
                "main_offloading_point",
                "main_route",
                "total_distance",
                "total_fuel_consumption_qty",
                "fuel_stock_out",
                "reference_doctype",
                "reference_docname",
                "transportation_order",
                "company",
                "main_status",
                "status",
                "invoice_number",
            ],
            "main_route_steps": [
                "name",
                "location",
                "distance",
                "location_type",
                "arrival_date",
                "departure_date",
                "loading_date",
                "offloading_date",
                "fuel_consumption_qty",
                "comment",
            ]
        },
        "filename": "vehicle_trips_hazchem_import_v15",
    },
    "Sales Invoice": {
        "fields": {
            "Sales Invoice": [
                "name",
                "posting_date",
                "due_date",
                "customer",
                "company",
                "currency",
                "remarks",
                "docstatus",
            ],
            "items": [
                "name",
                "item_code",
                "qty",
                "uom",
                "rate",
                "amount",
                "description",
            ],
        },
        "filename": "sales_invoices_hazchem_import_v15",
    },
}


@frappe.whitelist()
def setup_hazchem_operations(report_date=None, output_dir=None):
    report_date = getdate(report_date or today())
    output_path = Path(output_dir) if output_dir else DEFAULT_OUTPUT_DIR

    frappe.reload_doc("transport_management", "report", "daily_trip_schedule")

    cleanup_legacy_daily_schedule_demo()
    cleanup_existing_hazchem_transactions()

    ensure_vehicle_type("Chemical Tanker")
    customer_names = [ensure_customer(customer) for customer in CUSTOMERS]
    cargo_type_names = [ensure_transport_cargo_type(cargo_name) for cargo_name in CARGO_TYPES]
    service_item_names = [ensure_service_item(cargo_name) for cargo_name in CARGO_TYPES]
    transport_location_names = [ensure_transport_location(location_name) for location_name in TRANSPORT_LOCATIONS]
    trip_location_names = [ensure_trip_location(location_name) for location_name in TRIP_LOCATIONS]
    trailer_names = [ensure_trailer(trailer) for trailer in TRAILERS]
    driver_map = {driver["full_name"]: ensure_driver(driver, report_date) for driver in DRIVERS}
    vehicle_names = [ensure_vehicle(vehicle, driver_map, report_date) for vehicle in VEHICLES]

    order_names = []
    trip_names = []
    route_names = []
    invoice_names = []

    for operation in OPERATIONS:
        route_names.append(
            ensure_trip_route(
                operation["route_name"],
                operation["loading_point"],
                operation["offloading_point"],
                operation["distance"],
                operation["planned_fuel"],
            )
        )
        order_name, trip_name, invoice_name = create_order_assignment_and_trip(operation, driver_map, report_date)
        order_names.append(order_name)
        trip_names.append(trip_name)
        if invoice_name:
            invoice_names.append(invoice_name)

    records = {
        "customers": customer_names,
        "cargo_types": cargo_type_names,
        "service_items": service_item_names,
        "transport_locations": transport_location_names,
        "trip_locations": trip_location_names,
        "routes": sorted(set(route_names)),
        "drivers": list(driver_map.values()),
        "trailers": trailer_names,
        "vehicles": vehicle_names,
        "orders": order_names,
        "trips": trip_names,
        "sales_invoices": invoice_names,
    }

    ensure_workspace_custom_blocks(records, report_date)
    frappe.reload_doc("transport_management", "workspace", "transport")
    ensure_workspace_flow_block(records, report_date)
    exported_files = export_hazchem_import_sheets(records, output_path, report_date)

    frappe.clear_cache()
    frappe.db.commit()

    return {
        "report_date": str(report_date),
        "counts": {key: len(value) for key, value in records.items()},
        "files": exported_files,
    }


def cleanup_legacy_daily_schedule_demo():
    old_trip_names = frappe.get_all(
        "Vehicle Trip",
        filters={"vehicle_plate_number": ("in", OLD_DAILY_SCHEDULE_DEMO["plates"])},
        pluck="name",
    )
    for trip_name in old_trip_names:
        delete_doc_if_exists("Vehicle Trip", trip_name)

    for route_name in OLD_DAILY_SCHEDULE_DEMO["routes"]:
        if frappe.db.exists("Trip Route", route_name) and not frappe.db.exists("Vehicle Trip", {"main_route": route_name}):
            delete_doc_if_exists("Trip Route", route_name)

    for location_name in OLD_DAILY_SCHEDULE_DEMO["locations"]:
        if frappe.db.exists("Trip Location", location_name) and not frappe.db.exists("Route Steps Table", {"location": location_name}):
            delete_doc_if_exists("Trip Location", location_name)


def cleanup_existing_hazchem_transactions():
    hazchem_invoice_names = frappe.get_all(
        "Sales Invoice",
        filters=[["Sales Invoice", "remarks", "like", "%Hazchem Dispatch Ref:%"]],
        pluck="name",
    )
    for invoice_name in hazchem_invoice_names:
        delete_doc_if_exists("Sales Invoice", invoice_name)

    hazchem_trip_names = frappe.get_all(
        "Vehicle Trip",
        filters={"vehicle_plate_number": ("in", [vehicle["license_plate"] for vehicle in VEHICLES])},
        pluck="name",
    )
    for trip_name in hazchem_trip_names:
        delete_doc_if_exists("Vehicle Trip", trip_name)

    hazchem_order_names = frappe.get_all(
        "Transportation Order",
        filters=[["Transportation Order", "special_instructions_to_transporter", "like", "%Dispatch Ref: HAZCHEM-OPS-%"]],
        pluck="name",
    )
    for order_name in hazchem_order_names:
        delete_doc_if_exists("Transportation Order", order_name)


def delete_doc_if_exists(doctype, name):
    if not frappe.db.exists(doctype, name):
        return
    frappe.delete_doc(doctype, name, ignore_permissions=True, force=1)


def save_with_flags(doc):
    doc.flags.ignore_mandatory = True
    if doc.is_new():
        doc.insert(ignore_permissions=True, ignore_mandatory=True)
    else:
        doc.save(ignore_permissions=True)
    return doc


def ensure_vehicle_type(vehicle_type_name):
    if frappe.db.exists("Vehicle Type", vehicle_type_name):
        return vehicle_type_name

    doc = frappe.get_doc({"doctype": "Vehicle Type", "vehicle_type": vehicle_type_name})
    doc.insert(ignore_permissions=True, ignore_mandatory=True)
    return doc.name


def ensure_customer(customer):
    name = frappe.db.get_value("Customer", {"customer_name": customer["customer_name"]}, "name")
    if name:
        doc = frappe.get_doc("Customer", name)
    else:
        doc = frappe.new_doc("Customer")

    doc.update(customer)
    save_with_flags(doc)
    return doc.name


def ensure_transport_cargo_type(cargo_name):
    if frappe.db.exists("Transport Cargo Type", cargo_name):
        doc = frappe.get_doc("Transport Cargo Type", cargo_name)
    else:
        doc = frappe.new_doc("Transport Cargo Type")

    doc.cargo_name = cargo_name
    doc.permits = []
    save_with_flags(doc)
    return doc.name


def ensure_service_item(cargo_name):
    config = SERVICE_ITEMS[cargo_name]
    ensure_gst_hsn_code("999512", "Hazchem transport service")
    if frappe.db.exists("Item", config["item_code"]):
        doc = frappe.get_doc("Item", config["item_code"])
    else:
        doc = frappe.new_doc("Item")

    doc.item_code = config["item_code"]
    doc.item_name = config["item_name"]
    doc.item_group = "Services"
    doc.stock_uom = "Nos"
    doc.is_stock_item = 0
    doc.is_sales_item = 1
    doc.gst_hsn_code = "999512"
    doc.description = config["description"]
    doc.disabled = 0
    save_with_flags(doc)
    return doc.name


def ensure_gst_hsn_code(hsn_code, description):
    if not frappe.db.exists("DocType", "GST HSN Code"):
        return

    if frappe.db.exists("GST HSN Code", hsn_code):
        doc = frappe.get_doc("GST HSN Code", hsn_code)
    else:
        doc = frappe.new_doc("GST HSN Code")

    doc.hsn_code = hsn_code
    doc.description = description
    save_with_flags(doc)
    return doc.name


def ensure_transport_location(location_name):
    if frappe.db.exists("Transport Location", location_name):
        doc = frappe.get_doc("Transport Location", location_name)
    else:
        doc = frappe.new_doc("Transport Location")

    doc.location = location_name
    doc.country = "India"
    save_with_flags(doc)
    return doc.name


def ensure_trip_location(location_name):
    if frappe.db.exists("Trip Location", location_name):
        doc = frappe.get_doc("Trip Location", location_name)
    else:
        doc = frappe.new_doc("Trip Location")

    doc.description = location_name
    doc.location = ""
    doc.is_local_border = 0
    doc.is_international_border = 0
    save_with_flags(doc)
    return doc.name


def ensure_trailer(trailer):
    if frappe.db.exists("Trailer", trailer["number_plate"]):
        doc = frappe.get_doc("Trailer", trailer["number_plate"])
    else:
        doc = frappe.new_doc("Trailer")

    doc.number_plate = trailer["number_plate"]
    doc.chassis_number = trailer["chassis_number"]
    doc.make = trailer["make"]
    doc.year = trailer["year"]
    doc.axles = trailer["axles"]
    doc.trailer_type = "Fuel Tanker"
    doc.tyre_specification = "Double Tyres"
    doc.suspension_type = "Air Suspension"
    save_with_flags(doc)
    return doc.name


def ensure_driver(driver, report_date):
    name = frappe.db.get_value("Driver", {"full_name": driver["full_name"]}, "name")
    if name:
        doc = frappe.get_doc("Driver", name)
    else:
        doc = frappe.new_doc("Driver")

    doc.naming_series = "HR-DRI-.YYYY.-"
    doc.full_name = driver["full_name"]
    doc.status = "Active"
    doc.cell_number = driver["cell_number"]
    doc.license_number = driver["license_number"]
    doc.issuing_date = add_years(report_date, -4)
    doc.expiry_date = add_years(report_date, 2)
    save_with_flags(doc)
    return doc.name


def ensure_vehicle(vehicle, driver_map, report_date):
    if frappe.db.exists("Vehicle", vehicle["license_plate"]):
        doc = frappe.get_doc("Vehicle", vehicle["license_plate"])
    else:
        doc = frappe.new_doc("Vehicle")

    doc.license_plate = vehicle["license_plate"]
    doc.make = vehicle["make"]
    doc.model = vehicle["model"]
    doc.last_odometer = vehicle["last_odometer"]
    doc.acquisition_date = add_years(report_date, -3)
    doc.location = vehicle["location"]
    doc.trans_ms_default_trailer = vehicle["trailer"]
    doc.trans_ms_driver = driver_map[vehicle["driver"]]
    doc.trans_ms_fuel_consumption = vehicle["fuel_consumption"]
    doc.status = "Available"
    doc.trans_ms_current_trip = ""
    save_with_flags(doc)
    return doc.name


def ensure_trip_route(route_name, loading_point, offloading_point, distance, planned_fuel):
    if frappe.db.exists("Trip Route", route_name):
        doc = frappe.get_doc("Trip Route", route_name)
    else:
        doc = frappe.new_doc("Trip Route")

    ensure_trip_location(loading_point)
    ensure_trip_location(offloading_point)

    doc.route_name = route_name
    doc.total_distance = distance
    doc.total_fuel_consumption_qty = planned_fuel
    doc.trip_steps = []
    doc.append(
        "trip_steps",
        {
            "location": loading_point,
            "distance": 0,
            "location_type": "Loading Point",
            "fuel_consumption_qty": 0,
        },
    )
    doc.append(
        "trip_steps",
        {
            "location": offloading_point,
            "distance": distance,
            "location_type": "Offloading Point",
            "fuel_consumption_qty": planned_fuel,
        },
    )
    save_with_flags(doc)
    return doc.name


def get_service_item_for_operation(operation):
    return SERVICE_ITEMS[operation["cargo_type"]]["item_code"]


def get_service_rate_for_operation(operation):
    base_rate = SERVICE_RATE_PER_KM[operation["cargo_type"]]
    gross_rate = (operation["distance"] * base_rate) + (operation["net_weight"] * 250)
    return int(round(gross_rate / 100.0) * 100)


def build_trip_steps(operation, report_date):
    loading_step = {
        "location": operation["loading_point"],
        "distance": 0,
        "location_type": "Loading Point",
        "fuel_consumption_qty": 0,
        "arrival_date": report_date,
        "loading_date": report_date,
        "comment": "Vehicle reported at loading point for hazchem dispatch.",
    }
    offloading_step = {
        "location": operation["offloading_point"],
        "distance": operation["distance"],
        "location_type": "Offloading Point",
        "fuel_consumption_qty": operation["planned_fuel"],
        "comment": "Consignee delivery milestone for the assigned hazardous load.",
    }

    stage = (operation["stage"] or "").upper()
    if stage in {"LO", "EM", "WORKING", "UNL"}:
        loading_step["departure_date"] = report_date
    if stage == "LO":
        loading_step["comment"] = "Loaded and ready for gate-out from source terminal."
    elif stage == "EM":
        loading_step["comment"] = "Dispatched from source terminal and currently en route."
    elif stage == "WORKING":
        loading_step["comment"] = "Long-haul movement underway on approved hazchem route."
    elif stage == "UNL":
        loading_step["comment"] = "Dispatch completed from source and unloading started at destination."
        offloading_step["arrival_date"] = report_date
        offloading_step["offloading_date"] = report_date

    return [loading_step, offloading_step]


def create_hazchem_sales_invoice(order, assignment, trip, operation, report_date):
    item_details = frappe.db.get_value(
        "Item",
        assignment.item,
        ["item_name", "stock_uom"],
        as_dict=True,
    )
    description = "<b>VEHICLE NUMBER: {vehicle}<br>ROUTE: {route}<br>TRIP: {trip}</b>".format(
        vehicle=assignment.assigned_vehicle,
        route=assignment.route,
        trip=trip.name,
    )
    invoice = frappe.get_doc(
        {
            "doctype": "Sales Invoice",
            "naming_series": "ACC-SINV-.YYYY.-",
            "invoice_no": "HZINV-{0}".format(operation["dispatch_ref"].split("-")[-1]),
            "customer": order.customer,
            "posting_date": report_date,
            "due_date": report_date,
            "company": order.company,
            "currency": assignment.currency or "INR",
            "remarks": "Hazchem Dispatch Ref: {0}".format(operation["dispatch_ref"]),
            "items": [
                {
                    "item_code": assignment.item,
                    "item_name": item_details.item_name if item_details else assignment.item,
                    "gst_hsn_code": "999512",
                    "qty": 1,
                    "uom": (item_details.stock_uom if item_details else None) or "Nos",
                    "rate": assignment.rate,
                    "description": description,
                }
            ],
        }
    )
    invoice.flags.ignore_links = True
    invoice.flags.ignore_mandatory = True
    invoice.set_taxes()
    invoice.set_missing_values()
    invoice.calculate_taxes_and_totals()
    invoice.insert(ignore_permissions=True, ignore_mandatory=True, ignore_links=True)
    return invoice


def create_order_assignment_and_trip(operation, driver_map, report_date):
    order = frappe.new_doc("Transportation Order")
    service_item = get_service_item_for_operation(operation)
    service_rate = get_service_rate_for_operation(operation)
    order.date = report_date
    order.customer = operation["customer"]
    order.loading_date = report_date
    order.company = "Transport"
    order.transport_type = "Internal"
    order.cargo_type = "Loose Cargo"
    order.goods_description = operation["cargo_type"]
    order.amount = operation["net_weight"]
    order.unit = "Nos"
    order.assignment_status = "Waiting Assignment"
    order.version = 2
    order.special_instructions_to_transporter = (
        "Dispatch Ref: {dispatch_ref}. {instruction}".format(
            dispatch_ref=operation["dispatch_ref"],
            instruction=operation["extra_details"],
        )
    )
    order.append(
        "cargo",
        {
            "container_size": "Loose",
            "container_number": operation["container_number"],
            "net_weight": operation["net_weight"],
            "tare_weight": operation["tare_weight"],
            "cargo_location_country": "India",
            "cargo_location_city": operation["loading_city"],
            "cargo_destination_country": "India",
            "cargo_destination_city": operation["destination_city"],
            "cargo_type": operation["cargo_type"],
            "extra_details": operation["extra_details"],
        },
    )
    order.insert(ignore_permissions=True, ignore_mandatory=True)

    cargo_row = order.cargo[0]
    driver_name = frappe.db.get_value("Driver", driver_map[operation["driver"]], "full_name")
    order.append(
        "assign_transport",
        {
            "cargo": cargo_row.name,
            "amount": operation["net_weight"],
            "expected_loading_date": report_date,
            "container_number": operation["container_number"],
            "units": "Nos",
            "transporter_type": "In House",
            "assigned_vehicle": operation["vehicle"],
            "assigned_trailer": operation["trailer"],
            "assigned_driver": driver_map[operation["driver"]],
            "driver_name": driver_name,
            "route": operation["route_name"],
            "status": "Processed",
            "vehicle_plate_number": operation["vehicle"],
            "trailer_plate_number": operation["trailer"],
            "item": service_item,
            "rate": service_rate,
            "customer": operation["customer"],
            "currency": "INR",
        },
    )
    save_with_flags(order)
    order.reload()

    assignment = order.assign_transport[0]
    trip = frappe.get_doc(
        {
            "doctype": "Vehicle Trip",
            "date": report_date,
            "transporter_type": "In House",
            "vehicle": operation["vehicle"],
            "vehicle_plate_number": operation["vehicle"],
            "trailer": operation["trailer"],
            "trailer_plate_number": operation["trailer"],
            "driver": driver_map[operation["driver"]],
            "driver_name": driver_name,
            "customer": operation["customer"],
            "start_date": report_date,
            "main_cargo_category": operation["cargo_type"],
            "main_loading_point": operation["loading_point"],
            "main_offloading_point": operation["offloading_point"],
            "main_route": operation["route_name"],
            "total_distance": operation["distance"],
            "total_fuel_consumption_qty": operation["planned_fuel"],
            "fuel_stock_out": operation["fuel_issued"],
            "reference_doctype": "Transport Assignment",
            "reference_docname": assignment.name,
            "transportation_order": order.name,
            "company": "Transport",
            "main_status": operation["stage"],
            "status": "Open",
            "main_requested_funds": [],
            "main_fuel_request": [],
            "trip_permits": [],
            "main_route_steps": build_trip_steps(operation, report_date),
        }
    )
    trip.insert(ignore_permissions=True, ignore_mandatory=True)

    assignment.created_trip = trip.name
    assignment.status = "Processed"
    assignment.db_update()

    invoice_doc = create_hazchem_sales_invoice(order, assignment, trip, operation, report_date)
    if invoice_doc:
        assignment.invoice = invoice_doc.name
        assignment.db_update()
        trip.invoice_number = invoice_doc.name
        trip.save(ignore_permissions=True)

    vehicle = frappe.get_doc("Vehicle", operation["vehicle"])
    vehicle.status = "In Trip"
    vehicle.trans_ms_current_trip = trip.name
    save_with_flags(vehicle)

    return order.name, trip.name, invoice_doc.name if invoice_doc else ""


def ensure_workspace_flow_block(records, report_date):
    ensure_workspace_custom_blocks(records, report_date)
    workspace = frappe.get_doc("Workspace", WORKSPACE_NAME)
    needs_save = False

    if workspace.content != WORKSPACE_CONTENT:
        workspace.content = WORKSPACE_CONTENT
        needs_save = True

    for row in workspace.links:
        if row.label == "Cargo Type" and row.link_to == "Cargo Type":
            row.link_to = "Transport Cargo Type"
            needs_save = True

    existing_blocks = {row.custom_block_name for row in workspace.custom_blocks}
    for block_name in [WORKSPACE_FLOW_BLOCK_NAME, WORKSPACE_TIMELINE_BLOCK_NAME]:
        if block_name not in existing_blocks:
            workspace.append(
                "custom_blocks",
                {"custom_block_name": block_name, "label": block_name},
            )
            needs_save = True

    if not any(row.type == "Card Break" and row.label == "Reports" for row in workspace.links):
        workspace.append(
            "links",
            {
                "type": "Card Break",
                "label": "Reports",
                "link_type": "DocType",
                "hidden": 0,
                "onboard": 0,
                "link_count": 2,
            },
        )
        needs_save = True

    if not any(row.type == "Link" and row.link_type == "Report" and row.link_to == "Daily Trip Schedule" for row in workspace.links):
        workspace.append(
            "links",
            {
                "type": "Link",
                "label": "Daily Trip Schedule",
                "link_type": "Report",
                "link_to": "Daily Trip Schedule",
                "report_ref_doctype": "Vehicle Trip",
                "is_query_report": 1,
                "hidden": 0,
                "onboard": 0,
            },
        )
        needs_save = True

    if not any(row.type == "Link" and row.link_type == "Report" and row.link_to == "Vehicle Tracking Report" for row in workspace.links):
        workspace.append(
            "links",
            {
                "type": "Link",
                "label": "Vehicle Tracking Report",
                "link_type": "Report",
                "link_to": "Vehicle Tracking Report",
                "report_ref_doctype": "Vehicle Trip",
                "is_query_report": 1,
                "hidden": 0,
                "onboard": 0,
            },
        )
        needs_save = True

    for row in workspace.links:
        if row.type == "Card Break" and row.label == "Reports":
            row.link_count = 2
            needs_save = True

    if needs_save:
        workspace.save(ignore_permissions=True)

    return WORKSPACE_FLOW_BLOCK_NAME


def ensure_workspace_custom_blocks(records, report_date):
    get_or_create_custom_html_block(
        WORKSPACE_FLOW_BLOCK_NAME,
        build_workspace_html(records, report_date),
        build_workspace_style(),
    )
    get_or_create_custom_html_block(
        WORKSPACE_TIMELINE_BLOCK_NAME,
        build_workspace_timeline_html(),
        build_workspace_timeline_style(),
    )


def get_or_create_custom_html_block(block_name, html, style):
    if frappe.db.exists("Custom HTML Block", block_name):
        doc = frappe.get_doc("Custom HTML Block", block_name)
    else:
        doc = frappe.new_doc("Custom HTML Block")
        doc.name = block_name

    doc.html = html
    doc.style = style
    doc.script = ""
    doc.private = 0
    save_with_flags(doc)
    return doc


def build_workspace_html(records, report_date):
    first_order = records["orders"][0] if records["orders"] else ""
    first_trip = records["trips"][0] if records["trips"] else ""
    route_count = len(records["routes"])
    tanker_count = len(records["vehicles"])
    trip_count = len(records["trips"])

    return """
    <section class="tm-flow-shell">
        <div class="tm-flow-hero">
            <div>
                <div class="tm-flow-kicker">Hazardous Chemical Fleet Flow</div>
                <h2>Transport management operating path for tanker fleet dispatch</h2>
                <p>
                    This workspace is configured around a liquid-hazchem fleet. Tankers are maintained as
                    <strong>Vehicle</strong> masters, paired with tanker trailers, assigned to licensed drivers,
                    and then moved through transportation orders, route allocation, and live trip execution.
                </p>
            </div>
            <div class="tm-flow-metrics">
                <div><span>{tanker_count}</span><small>tankers prepared</small></div>
                <div><span>{route_count}</span><small>hazchem routes</small></div>
                <div><span>{trip_count}</span><small>live trips on {report_date}</small></div>
            </div>
        </div>
        <div class="tm-flow-grid">
            <article class="tm-step">
                <h3>1. Register fleet and compliance masters</h3>
                <p>Create tankers, tanker trailers, drivers, cargo classes, trip locations, and approved trip routes before dispatching any load.</p>
                <div class="tm-links">
                    <a href="/app/vehicle/view/list">Tankers / Vehicles</a>
                    <a href="/app/trailer/view/list">Tank Trailers</a>
                    <a href="/app/driver/view/list">Drivers</a>
                    <a href="/app/transport-cargo-type/view/list">Hazchem Cargo Types</a>
                    <a href="/app/trip-route/view/list">Trip Routes</a>
                </div>
            </article>
            <article class="tm-step">
                <h3>2. Capture customer movement demand</h3>
                <p>Each customer dispatch begins as a <strong>Transportation Order</strong> with loading date, source, destination, cargo type, and hazchem operating notes.</p>
                <div class="tm-links">
                    <a href="/app/transportation-order/view/list">Transportation Orders</a>
                    <a href="/app/transportation-order/{first_order}">Open live order</a>
                    <a href="/app/customer/view/list">Customers</a>
                    <a href="/app/transport-location/view/list">Transport Locations</a>
                </div>
            </article>
            <article class="tm-step">
                <h3>3. Allocate tanker, trailer, and driver</h3>
                <p>Assignments tie the hazchem load to the actual tanker combination, nominated driver, loading slot, and approved route. This becomes the operational handoff to the fleet desk.</p>
                <div class="tm-links">
                    <a href="/app/transportation-order/{first_order}">Assignments inside order</a>
                    <a href="/app/vehicle/view/list">Check tanker availability</a>
                    <a href="/app/driver/view/list">Validate driver licence</a>
                </div>
            </article>
            <article class="tm-step">
                <h3>4. Execute trip and monitor movement</h3>
                <p>When the tanker departs, create and maintain the <strong>Vehicle Trip</strong>. Track route, stage, diesel issue, loading status, and final offloading against the order reference.</p>
                <div class="tm-links">
                    <a href="/app/vehicle-trip/view/list">Vehicle Trips</a>
                    <a href="/app/vehicle-trip/{first_trip}">Open live trip</a>
                    <a href="/app/fuel-request/view/list">Fuel Requests</a>
                    <a href="/app/requested-payments/view/list">Requested Payments</a>
                </div>
            </article>
            <article class="tm-step">
                <h3>5. Review daily tanker board and route progress</h3>
                <p>The daily control board gives dispatch visibility by tanker, programme, route distance, diesel issue, and operational stage for the fleet team.</p>
                <div class="tm-links">
                    <a href="/app/query-report/{daily_schedule}">Daily Trip Schedule</a>
                    <a href="/app/query-report/{vehicle_tracking}">Vehicle Tracking Report</a>
                    <a href="/app/workspace/Transport">Refresh Transport Workspace</a>
                </div>
            </article>
            <article class="tm-step">
                <h3>6. Recommended hazchem controls</h3>
                <p>Keep each order updated with hazard notes, route constraints, emergency kit checks, unloading window, and customer escalation contacts so operations, finance, and compliance stay aligned.</p>
                <ul>
                    <li>Use cargo types to separate caustic, acid, and oxidizer movements.</li>
                    <li>Keep tankers and trailers paired for loading history and wash certification.</li>
                    <li>Review the daily schedule report before issuing fuel or releasing the next dispatch wave.</li>
                </ul>
            </article>
        </div>
    </section>
    """.format(
        tanker_count=tanker_count,
        route_count=route_count,
        trip_count=trip_count,
        report_date=str(report_date),
        first_order=quote(first_order),
        first_trip=quote(first_trip),
        daily_schedule=quote("Daily Trip Schedule"),
        vehicle_tracking=quote("Vehicle Tracking Report"),
    )


def build_workspace_timeline_html():
    steps = [
        ("Customer", "/app/customer/view/list"),
        ("Transportation Order", "/app/transportation-order/view/list"),
        ("Check Tanker Availability", "/app/vehicle/view/list"),
        ("Assign Vehicle", "/app/transportation-order/view/list"),
        ("Start Vehicle Trip", "/app/vehicle-trip/view/list"),
        ("Fuel Request", "/app/fuel-request/view/list"),
        ("Vehicle Tracking Report", "/app/query-report/{0}".format(quote("Vehicle Tracking Report"))),
        ("Sales Invoice", "/app/sales-invoice/view/list"),
        ("Close Trip", "/app/vehicle-trip/view/list"),
    ]

    return """
    <section class="tm-mini-shell">
        <div class="tm-mini-track">
            {steps}
        </div>
    </section>
    """.format(
        steps="".join(
            [
                """
                <a class="tm-mini-step" href="{link}"><span>{label}</span></a>
                """.format(link=link, label=label)
                for label, link in steps
            ]
        )
    )


def build_workspace_style():
    return """
    .tm-flow-shell {
        font-family: "Segoe UI", Arial, sans-serif;
        color: #122230;
        display: grid;
        gap: 18px;
    }

    .tm-flow-hero {
        display: grid;
        grid-template-columns: 2.3fr 1fr;
        gap: 16px;
        padding: 18px 20px;
        border-radius: 18px;
        background: linear-gradient(135deg, #eef6ff 0%, #f8fbff 60%, #fff5e6 100%);
        border: 1px solid #d9e6f2;
    }

    .tm-flow-kicker {
        text-transform: uppercase;
        letter-spacing: 0.08em;
        font-size: 11px;
        font-weight: 700;
        color: #9a3412;
        margin-bottom: 8px;
    }

    .tm-flow-hero h2 {
        margin: 0 0 10px;
        font-size: 24px;
        line-height: 1.2;
    }

    .tm-flow-hero p {
        margin: 0;
        line-height: 1.6;
        color: #334155;
    }

    .tm-flow-metrics {
        display: grid;
        gap: 10px;
        align-content: start;
    }

    .tm-flow-metrics div {
        background: rgba(255, 255, 255, 0.86);
        border: 1px solid #d6e3ef;
        border-radius: 14px;
        padding: 14px;
    }

    .tm-flow-metrics span {
        display: block;
        font-size: 24px;
        font-weight: 700;
        color: #0f4c81;
    }

    .tm-flow-metrics small {
        display: block;
        margin-top: 4px;
        color: #475569;
    }

    .tm-flow-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
        gap: 14px;
    }

    .tm-step {
        padding: 16px;
        border-radius: 16px;
        border: 1px solid #d9e2ec;
        background: #ffffff;
        box-shadow: 0 10px 24px rgba(15, 76, 129, 0.06);
    }

    .tm-step h3 {
        margin: 0 0 10px;
        font-size: 17px;
        line-height: 1.35;
    }

    .tm-step p,
    .tm-step li {
        margin: 0;
        color: #475569;
        line-height: 1.6;
    }

    .tm-step ul {
        margin: 10px 0 0 18px;
        padding: 0;
    }

    .tm-links {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-top: 12px;
    }

    .tm-links a {
        text-decoration: none;
        border-radius: 999px;
        padding: 7px 11px;
        background: #edf4ff;
        color: #0f4c81;
        border: 1px solid #c9dcf5;
        font-size: 12px;
        font-weight: 600;
    }

    @media (max-width: 920px) {
        .tm-flow-hero {
            grid-template-columns: 1fr;
        }
    }
    """


def build_workspace_timeline_style():
    return """
    .tm-mini-shell {
        overflow-x: auto;
        padding-bottom: 6px;
    }

    .tm-mini-track {
        display: inline-flex;
        align-items: center;
        gap: 14px;
        min-width: 100%;
        padding: 4px 2px;
    }

    .tm-mini-step {
        position: relative;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-width: 168px;
        padding: 12px 16px;
        border-radius: 999px;
        text-decoration: none;
        background: linear-gradient(135deg, #fff7ed 0%, #eff6ff 100%);
        border: 1px solid #d7e3f1;
        color: #1e293b;
        font-size: 12px;
        font-weight: 700;
        box-shadow: 0 6px 18px rgba(15, 76, 129, 0.07);
    }

    .tm-mini-step:not(:last-child)::after {
        content: "";
        position: absolute;
        right: -14px;
        top: 50%;
        width: 14px;
        height: 2px;
        background: #94a3b8;
        transform: translateY(-50%);
    }
    """


def export_hazchem_import_sheets(records, output_dir, report_date):
    output_dir.mkdir(parents=True, exist_ok=True)
    exported_files = []

    export_filters = {
        "Customer": {"name": ("in", records["customers"])},
        "Driver": {"name": ("in", records["drivers"])},
        "Trailer": {"name": ("in", records["trailers"])},
        "Vehicle": {"name": ("in", records["vehicles"])},
        "Item": {"name": ("in", records["service_items"])},
        "Transport Location": {"name": ("in", records["transport_locations"])},
        "Trip Location": {"name": ("in", records["trip_locations"])},
        "Transport Cargo Type": {"name": ("in", records["cargo_types"])},
        "Trip Route": {"name": ("in", records["routes"])},
        "Transportation Order": {"name": ("in", records["orders"])},
        "Vehicle Trip": {"name": ("in", records["trips"])},
        "Sales Invoice": {"name": ("in", records["sales_invoices"])},
    }

    for doctype, config in EXPORT_CONFIG.items():
        exporter = Exporter(
            doctype=doctype,
            export_fields=config["fields"],
            export_data=True,
            export_filters=export_filters[doctype],
            export_page_length=500,
            file_type="CSV",
        )
        rows = exporter.get_csv_array_for_export()
        csv_path = output_dir / "{0}.csv".format(config["filename"])
        xlsx_path = output_dir / "{0}.xlsx".format(config["filename"])
        write_csv(rows, csv_path)
        write_xlsx(rows, xlsx_path)
        exported_files.extend([str(csv_path), str(xlsx_path)])

    readme_path = output_dir / "README.md"
    readme_path.write_text(
        build_import_readme(records, report_date),
        encoding="utf-8",
    )
    exported_files.append(str(readme_path))
    return exported_files


def write_csv(rows, path):
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerows(rows)


def write_xlsx(rows, path):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Import"
    for row in rows:
        worksheet.append(row)
    workbook.save(path)


def build_import_readme(records, report_date):
    return """# Hazchem Transport Import Set

Report date used for operational seed: `{report_date}`

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
- Customers: {customers}
- Drivers: {drivers}
- Trailers: {trailers}
- Vehicles: {vehicles}
- Service Items: {service_items}
- Transport Locations: {transport_locations}
- Trip Locations: {trip_locations}
- Cargo Types: {cargo_types}
- Trip Routes: {routes}
- Transportation Orders: {orders}
- Vehicle Trips: {trips}
- Sales Invoices: {sales_invoices}

Notes:
- Tanker is represented by the `Vehicle` master.
- The daily tanker board is available in the `Daily Trip Schedule` report.
- Transportation orders include child-table exports for cargo and assignment rows so the production import preserves dispatch structure.
- Vehicle trip exports include `main_route_steps`, which the tracking report uses for arrival and departure milestones.
""".format(
        report_date=report_date,
        customers=len(records["customers"]),
        drivers=len(records["drivers"]),
        trailers=len(records["trailers"]),
        vehicles=len(records["vehicles"]),
        service_items=len(records["service_items"]),
        transport_locations=len(records["transport_locations"]),
        trip_locations=len(records["trip_locations"]),
        cargo_types=len(records["cargo_types"]),
        routes=len(records["routes"]),
        orders=len(records["orders"]),
        trips=len(records["trips"]),
        sales_invoices=len(records["sales_invoices"]),
    )
