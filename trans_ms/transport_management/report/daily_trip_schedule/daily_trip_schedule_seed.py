from __future__ import unicode_literals

import frappe
from frappe.utils import getdate, today

DEMO_TRIPS = [
	{"plate": "9171", "route_name": "DAHEJ / HOPL", "loading_point": "DAHEJ", "offloading_point": "HOPL", "distance": 180, "planned_fuel": 75, "fuel_issued": 60, "stage": "EM"},
	{"plate": "5472", "route_name": "BHIWADI / MODASA", "loading_point": "BHIWADI", "offloading_point": "MODASA", "distance": 470, "planned_fuel": 165, "fuel_issued": 150, "stage": "LO"},
	{"plate": "7299", "route_name": "DAHEJ / HOPL", "loading_point": "DAHEJ", "offloading_point": "HOPL", "distance": 180, "planned_fuel": 75, "fuel_issued": 70, "stage": "EM"},
	{"plate": "5390", "route_name": "DAHEJ / HOPL", "loading_point": "DAHEJ", "offloading_point": "HOPL", "distance": 180, "planned_fuel": 75, "fuel_issued": 80, "stage": "WORKING"},
	{"plate": "271", "route_name": "HOPL / SRF", "loading_point": "HOPL", "offloading_point": "SRF", "distance": 225, "planned_fuel": 90, "fuel_issued": 65, "stage": "EM"},
	{"plate": "1334", "route_name": "CHATRAL / BODAL", "loading_point": "CHATRAL", "offloading_point": "BODAL", "distance": 145, "planned_fuel": 42, "fuel_issued": 50, "stage": "LO"},
	{"plate": "6391", "route_name": "VAPI / BODAL S", "loading_point": "VAPI", "offloading_point": "BODAL S", "distance": 120, "planned_fuel": 32, "fuel_issued": 30, "stage": "EM"},
	{"plate": "6199", "route_name": "ANKLESHWAR / BODAL", "loading_point": "ANKLESHWAR", "offloading_point": "BODAL", "distance": 130, "planned_fuel": 34, "fuel_issued": 36, "stage": "LO"},
	{"plate": "1470", "route_name": "KHAMBHAT / BODAL", "loading_point": "KHAMBHAT", "offloading_point": "BODAL", "distance": 140, "planned_fuel": 38, "fuel_issued": 32, "stage": "EM"},
	{"plate": "6099", "route_name": "BIRLA / MADHU", "loading_point": "BIRLA", "offloading_point": "MADHU", "distance": 520, "planned_fuel": 190, "fuel_issued": 175, "stage": "UNL"},
]


def reload_daily_trip_schedule():
	frappe.reload_doc("transport_management", "report", "daily_trip_schedule")
	return "Daily Trip Schedule reloaded"


@frappe.whitelist()
def seed_demo_daily_trip_schedule(report_date=None):
	report_date = getdate(report_date or today())
	created = []

	for row in DEMO_TRIPS:
		ensure_trip_route(
			row["route_name"],
			row["loading_point"],
			row["offloading_point"],
			row["distance"],
			row["planned_fuel"],
		)

		existing_trip = frappe.db.get_value(
			"Vehicle Trip",
			{"date": report_date, "vehicle_plate_number": row["plate"], "main_route": row["route_name"]},
		)
		if existing_trip:
			continue

		doc = frappe.get_doc(
			{
				"doctype": "Vehicle Trip",
				"date": report_date,
				"start_date": report_date,
				"transporter_type": "Sub-Contractor",
				"vehicle_plate_number": row["plate"],
				"driver_name": "Demo Driver {0}".format(row["plate"]),
				"main_route": row["route_name"],
				"main_loading_point": row["loading_point"],
				"main_offloading_point": row["offloading_point"],
				"total_distance": row["distance"],
				"total_fuel_consumption_qty": row["planned_fuel"],
				"fuel_stock_out": row["fuel_issued"],
				"main_status": row["stage"],
				"status": "Open",
				"main_requested_funds": [],
				"main_fuel_request": [],
				"trip_permits": [],
				"main_route_steps": [],
			}
		)
		doc.insert(ignore_permissions=True, ignore_mandatory=True)
		created.append(doc.name)

	frappe.db.commit()
	return {
		"report_date": str(report_date),
		"created_count": len(created),
		"created_trips": created,
	}


def ensure_trip_route(route_name, loading_point, offloading_point, distance, planned_fuel):
	if frappe.db.exists("Trip Route", route_name):
		return route_name

	ensure_trip_location(loading_point)
	ensure_trip_location(offloading_point)

	doc = frappe.get_doc(
		{
			"doctype": "Trip Route",
			"route_name": route_name,
			"total_distance": distance,
			"total_fuel_consumption_qty": planned_fuel,
			"trip_steps": [
				{
					"location": loading_point,
					"distance": 0,
					"location_type": "Loading Point",
					"fuel_consumption_qty": 0,
				},
				{
					"location": offloading_point,
					"distance": distance,
					"location_type": "Offloading Point",
					"fuel_consumption_qty": planned_fuel,
				},
			],
		}
	)
	doc.insert(ignore_permissions=True, ignore_mandatory=True)
	return doc.name


def ensure_trip_location(description):
	if frappe.db.exists("Trip Location", description):
		return description

	doc = frappe.get_doc(
		{
			"doctype": "Trip Location",
			"description": description,
		}
	)
	doc.insert(ignore_permissions=True, ignore_mandatory=True)
	return doc.name
