# Copyright (c) 2023, Aakvatech Limited and contributors
# For license information, please see license.txt

from __future__ import unicode_literals

from collections import OrderedDict

import frappe
from frappe import _
from frappe.utils import flt


def execute(filters=None):
	filters = frappe._dict(filters or {})
	trips = get_trip_details(filters)
	if not trips:
		frappe.msgprint(_("<b>No data found, Please check your filters</b>"))
		return [], []

	trip_names = [row.name for row in trips]
	trip_steps = get_trip_steps(trip_names)
	steps_by_trip, step_columns = group_trip_steps(trip_steps)

	columns = get_columns(step_columns)
	data = [build_row(trip, steps_by_trip.get(trip.name, [])) for trip in trips]
	return columns, data


def get_columns(step_columns):
	columns = [
		{"fieldname": "trip", "fieldtype": "Link", "label": _("Trip"), "options": "Vehicle Trip", "width": 110},
		{
			"fieldname": "transportation_order",
			"fieldtype": "Link",
			"label": _("Transportation Order"),
			"options": "Transportation Order",
			"width": 140,
		},
		{
			"fieldname": "sales_invoice",
			"fieldtype": "Link",
			"label": _("Sales Invoice"),
			"options": "Sales Invoice",
			"width": 120,
		},
		{"fieldname": "customer", "fieldtype": "Link", "label": _("Customer"), "options": "Customer", "width": 140},
		{"fieldname": "tanker", "fieldtype": "Link", "label": _("Tanker"), "options": "Vehicle", "width": 110},
		{"fieldname": "trailer", "fieldtype": "Link", "label": _("Trailer"), "options": "Trailer", "width": 110},
		{"fieldname": "driver", "fieldtype": "Link", "label": _("Driver"), "options": "Driver", "width": 120},
		{"fieldname": "driver_name", "fieldtype": "Data", "label": _("Driver Name"), "width": 130},
		{"fieldname": "contact_number", "fieldtype": "Data", "label": _("Contact Number"), "width": 110},
		{
			"fieldname": "cargo_type",
			"fieldtype": "Link",
			"label": _("Cargo Type"),
			"options": "Transport Cargo Type",
			"width": 140,
		},
		{"fieldname": "service_item", "fieldtype": "Link", "label": _("Service Item"), "options": "Item", "width": 140},
		{"fieldname": "rate", "fieldtype": "Currency", "label": _("Rate"), "width": 100},
		{"fieldname": "tonnage", "fieldtype": "Float", "label": _("Tonnage"), "width": 90},
		{"fieldname": "main_route", "fieldtype": "Link", "label": _("Route"), "options": "Trip Route", "width": 140},
		{"fieldname": "position", "fieldtype": "Data", "label": _("Current Position"), "width": 180},
		{"fieldname": "trip_status", "fieldtype": "Data", "label": _("Trip Status"), "width": 100},
		{"fieldname": "loaded_date", "fieldtype": "Date", "label": _("Loaded Date"), "width": 95},
		{"fieldname": "tracking_date", "fieldtype": "Date", "label": _("Tracking Date"), "width": 95},
	]

	for step in step_columns:
		step_key = frappe.scrub(step.location)
		columns.append(
			{
				"fieldname": "arrival_" + step_key,
				"fieldtype": "Date",
				"label": _("Arrived " + step.location),
				"width": 100,
			}
		)
		columns.append(
			{
				"fieldname": "departure_" + step_key,
				"fieldtype": "Date",
				"label": _("Depart " + step.location),
				"width": 100,
			}
		)

	return columns


def get_trip_details(filters):
	conditions = get_conditions(filters)
	return frappe.db.sql(
		"""
		select
			vt.name,
			vt.vehicle,
			vt.vehicle_plate_number,
			vt.trailer,
			vt.trailer_plate_number,
			vt.customer,
			vt.driver,
			vt.driver_name,
			vt.start_date,
			vt.main_route,
			vt.main_status,
			vt.status,
			vt.transportation_order,
			vt.invoice_number,
			vt.main_loading_point,
			vt.main_offloading_point,
			d.cell_number,
			ta.item as service_item,
			ta.rate,
			ta.currency,
			cd.net_weight,
			cd.cargo_type
		from `tabVehicle Trip` vt
		left join `tabDriver` d on vt.driver = d.name
		left join `tabTransport Assignment` ta
			on vt.reference_doctype = 'Transport Assignment'
			and vt.reference_docname = ta.name
		left join `tabCargo Details` cd on ta.cargo = cd.name
		where vt.docstatus < 2 {conditions}
		order by vt.start_date desc, vt.name desc
		""".format(conditions=conditions),
		filters,
		as_dict=True,
	)


def get_trip_steps(trip_names):
	if not trip_names:
		return []

	return frappe.db.sql(
		"""
		select
			ts.parent,
			ts.location,
			ts.location_type,
			ts.arrival_date,
			ts.departure_date,
			ts.loading_date,
			ts.offloading_date
		from `tabRoute Steps Table` ts
		where ts.parenttype = 'Vehicle Trip'
			and ts.parentfield = 'main_route_steps'
			and ts.parent in %(trip_names)s
		order by ts.parent, ts.idx asc
		""",
		{"trip_names": trip_names},
		as_dict=True,
	)


def group_trip_steps(trip_steps):
	steps_by_trip = OrderedDict()
	step_columns = []
	seen = set()

	for step in trip_steps:
		steps_by_trip.setdefault(step.parent, []).append(step)
		key = step.location
		if key in seen:
			continue
		seen.add(key)
		step_columns.append(step)

	return steps_by_trip, step_columns


def build_row(trip, steps):
	position, tracking_date = get_position_and_tracking_date(trip, steps)
	row = {
		"trip": trip.name,
		"transportation_order": trip.transportation_order,
		"sales_invoice": trip.invoice_number,
		"customer": trip.customer,
		"tanker": trip.vehicle,
		"trailer": trip.trailer,
		"driver": trip.driver,
		"driver_name": trip.driver_name,
		"contact_number": trip.cell_number,
		"cargo_type": trip.cargo_type,
		"service_item": trip.service_item,
		"rate": trip.rate,
		"tonnage": flt(trip.net_weight),
		"main_route": trip.main_route,
		"position": position,
		"trip_status": trip.main_status or trip.status or "Open",
		"loaded_date": trip.start_date,
		"tracking_date": tracking_date,
	}

	for step in steps:
		step_key = frappe.scrub(step.location)
		row["arrival_" + step_key] = step.arrival_date
		row["departure_" + step_key] = step.departure_date

	return row


def get_position_and_tracking_date(trip, steps):
	events = []
	for step in steps:
		if step.loading_date:
			events.append((step.loading_date, "{0} (Loading)".format(step.location)))
		if step.departure_date and step.location_type == "Loading Point":
			events.append(
				(
					step.departure_date,
					"En route to {0}".format(trip.main_offloading_point or step.location),
				)
			)
		if step.arrival_date and step.location_type == "Offloading Point":
			events.append((step.arrival_date, "{0} (Arrived)".format(step.location)))
		if step.offloading_date:
			events.append((step.offloading_date, "{0} (Offloading)".format(step.location)))

	if not events:
		return trip.main_status or trip.status or "Open", trip.start_date

	tracking_date, position = events[-1]
	return position, tracking_date


def get_conditions(filters):
	conditions = ""
	if filters.get("from_date"):
		conditions += " and vt.start_date >= %(from_date)s"
	if filters.get("to_date"):
		conditions += " and vt.start_date <= %(to_date)s"
	if filters.get("customer"):
		conditions += " and vt.customer = %(customer)s"
	if filters.get("transportation_order"):
		conditions += " and vt.transportation_order = %(transportation_order)s"
	if filters.get("vehicle"):
		conditions += " and vt.vehicle = %(vehicle)s"
	if filters.get("transporter_type"):
		conditions += " and vt.transporter_type = %(transporter_type)s"
	if filters.get("driver"):
		conditions += " and vt.driver = %(driver)s"
	if filters.get("main_route"):
		conditions += " and vt.main_route = %(main_route)s"

	return conditions
