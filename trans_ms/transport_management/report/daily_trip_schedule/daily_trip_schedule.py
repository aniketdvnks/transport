# Copyright (c) 2026, contributors
# For license information, please see license.txt

from __future__ import unicode_literals

from collections import OrderedDict
from urllib.parse import quote

import frappe
from frappe import _
from frappe.utils import escape_html, flt, formatdate, getdate, today

BLOCK_COUNT = 3
BLOCK_FIELDS = (
	("trip", _("Trip"), "Link", 92, "Vehicle Trip"),
	("tanker", _("Tanker"), "Link", 95, "Vehicle"),
	("programme", _("Programme"), "Data", 170),
	("km", _("KM"), "Float", 70),
	("stage", _("Filled/Stage"), "Data", 90),
	("avg", _("AVG"), "Float", 70),
	("planned_fuel", _("Diesel Planned"), "Float", 95),
	("fuel_issued", _("Diesel Given"), "Float", 90),
	("fuel_gap", _("Diesel Gap"), "Float", 85),
	("balance", _("Balance"), "Float", 85),
)


def execute(filters=None):
	if isinstance(filters, str):
		filters = frappe.parse_json(filters)

	filters = frappe._dict(filters or {})
	report_date = getdate(filters.get("report_date") or today())
	trips = get_trip_rows(filters, report_date)

	columns = get_columns()
	data = build_matrix_rows(trips)
	message = build_message(report_date, trips)
	report_summary = get_report_summary(trips)

	return columns, data, message, None, report_summary


def get_columns():
	columns = []

	for block_number in range(1, BLOCK_COUNT + 1):
		for field in BLOCK_FIELDS:
			fieldname, label, fieldtype, width = field[:4]
			column = {
				"fieldname": "{0}_{1}".format(fieldname, block_number),
				"label": label,
				"fieldtype": fieldtype,
				"width": width,
			}
			if len(field) > 4:
				column["options"] = field[4]
			columns.append(column)

	return columns


def get_trip_rows(filters, report_date):
	conditions = ["trip.docstatus < 2", "trip.date = %(report_date)s"]
	values = {"report_date": report_date}

	if filters.get("main_route"):
		conditions.append("trip.main_route = %(main_route)s")
		values["main_route"] = filters.get("main_route")

	if filters.get("tanker"):
		conditions.append(
			"coalesce(nullif(trip.vehicle_plate_number, ''), nullif(vehicle.license_plate, ''), nullif(trip.vehicle, ''), trip.name) like %(tanker)s"
		)
		values["tanker"] = "%{0}%".format(filters.get("tanker").strip())

	if filters.get("stage"):
		conditions.append(
			"upper(coalesce(nullif(trip.main_status, ''), nullif(trip.status, ''), 'OPEN')) = %(stage)s"
		)
		values["stage"] = filters.get("stage").upper()

	rows = frappe.db.sql(
		"""
			select
				trip.name,
				trip.main_route,
				trip.main_loading_point,
				trip.main_offloading_point,
				trip.transportation_order,
				trip.invoice_number,
				coalesce(nullif(trip.vehicle_plate_number, ''), nullif(vehicle.license_plate, ''), nullif(trip.vehicle, ''), trip.name) as tanker_alias,
				trip.main_status,
				trip.status,
				trip.total_distance,
				trip.total_fuel_consumption_qty,
			trip.fuel_stock_out
		from `tabVehicle Trip` trip
		left join `tabVehicle` vehicle on vehicle.name = trip.vehicle
		where {conditions}
		order by
			coalesce(nullif(trip.main_route, ''), nullif(trip.main_loading_point, ''), trip.name),
			trip.vehicle_plate_number,
			trip.name
		""".format(conditions=" and ".join(conditions)),
		values,
		as_dict=True,
	)

	return [prepare_trip_row(row) for row in rows]


def prepare_trip_row(row):
	tanker = row.tanker_alias or row.name
	route_bucket = row.main_route or build_programme(row.main_loading_point, row.main_offloading_point) or row.name
	programme = build_programme(row.main_loading_point, row.main_offloading_point) or route_bucket
	stage = row.main_status or row.status or "OPEN"
	km = flt(row.total_distance)
	planned_fuel = flt(row.total_fuel_consumption_qty)
	fuel_issued = flt(row.fuel_stock_out)
	fuel_gap = max(planned_fuel - fuel_issued, 0)
	balance = fuel_issued - planned_fuel
	avg = round(km / planned_fuel, 2) if planned_fuel else 0

	return {
		"name": row.name,
		"trip": row.name,
		"tanker": tanker,
		"route_bucket": route_bucket,
		"programme": programme,
		"stage": stage,
		"km": round(km, 2),
		"avg": avg,
		"planned_fuel": round(planned_fuel, 2),
		"fuel_issued": round(fuel_issued, 2),
		"fuel_gap": round(fuel_gap, 2),
		"balance": round(balance, 2),
		"order": row.transportation_order,
		"invoice": row.invoice_number,
	}


def build_programme(loading_point, offloading_point):
	parts = [part for part in [loading_point, offloading_point] if part]
	return " / ".join(parts)


def build_matrix_rows(trips):
	data = []

	for row_index in range(0, len(trips), BLOCK_COUNT):
		chunk = trips[row_index : row_index + BLOCK_COUNT]
		row = {}

		for block_number in range(1, BLOCK_COUNT + 1):
			if block_number > len(chunk):
				continue

			trip = chunk[block_number - 1]
			row["trip_{0}".format(block_number)] = trip["trip"]
			row["tanker_{0}".format(block_number)] = trip["tanker"]
			row["programme_{0}".format(block_number)] = trip["programme"]
			row["km_{0}".format(block_number)] = trip["km"]
			row["stage_{0}".format(block_number)] = trip["stage"]
			row["avg_{0}".format(block_number)] = trip["avg"]
			row["planned_fuel_{0}".format(block_number)] = trip["planned_fuel"]
			row["fuel_issued_{0}".format(block_number)] = trip["fuel_issued"]
			row["fuel_gap_{0}".format(block_number)] = trip["fuel_gap"]
			row["balance_{0}".format(block_number)] = trip["balance"]

		data.append(row)

	return data


def build_message(report_date, trips):
	if not trips:
		return _(
			"<div><strong>Daily Schedule</strong> : Date : {0}<br>No trips found for this date.</div>"
		).format(escape_html(formatdate(report_date)))

	route_map = OrderedDict()

	for trip in trips:
		route_map.setdefault(trip["route_bucket"], []).append(trip["tanker"])

	cards = []
	for route_name, tankers in route_map.items():
		tanker_preview = []
		for trip in [trip for trip in trips if trip["route_bucket"] == route_name][:4]:
			links = [
				'<a href="/app/vehicle-trip/{trip_name}">{tanker}</a>'.format(
					trip_name=quote(trip["trip"]), tanker=escape_html(trip["tanker"])
				)
			]
			if trip.get("order"):
				links.append(
					'<a href="/app/transportation-order/{order_name}">Order</a>'.format(
						order_name=quote(trip["order"])
					)
				)
			if trip.get("invoice"):
				links.append(
					'<a href="/app/sales-invoice/{invoice_name}">Invoice</a>'.format(
						invoice_name=quote(trip["invoice"])
					)
				)
			tanker_preview.append(" · ".join(links))

		tanker_preview = "<br>".join(tanker_preview)
		if len(tankers) > 4:
			tanker_preview += "<br>..."

		cards.append(
			"""
			<div style="min-width: 180px; border: 1px solid #d1d8dd; border-radius: 8px; padding: 10px 12px; background: #f8fafc;">
				<div style="font-weight: 600; margin-bottom: 4px;">{route_name}</div>
				<div style="font-size: 12px; color: #6b7280; margin-bottom: 6px;">{count} tanker(s)</div>
				<div style="font-size: 12px; color: #111827; line-height: 1.6;">{tankers}</div>
			</div>
			""".format(
				route_name=escape_html(route_name),
				count=len(tankers),
				tankers=tanker_preview,
			)
		)

	return """
	<div style="margin-bottom: 14px;">
		<div style="font-size: 15px; font-weight: 600;">DAILY SCHEDULE : DATE : {date}</div>
		<div style="display: flex; gap: 10px; flex-wrap: wrap; margin-top: 10px;">
			{cards}
		</div>
	</div>
	""".format(date=escape_html(formatdate(report_date)), cards="".join(cards))


def get_report_summary(trips):
	total_km = round(sum(flt(trip["km"]) for trip in trips), 2)
	total_planned_fuel = round(sum(flt(trip["planned_fuel"]) for trip in trips), 2)
	total_gap = round(sum(flt(trip["fuel_gap"]) for trip in trips), 2)
	routes = len({trip["route_bucket"] for trip in trips if trip["route_bucket"]})

	return [
		{"label": _("Trips"), "value": len(trips), "indicator": "Blue"},
		{"label": _("Routes"), "value": routes, "indicator": "Green"},
		{"label": _("Planned KM"), "value": total_km, "indicator": "Orange", "datatype": "Float"},
		{
			"label": _("Fuel Gap"),
			"value": total_gap,
			"indicator": "Red" if total_gap else "Green",
			"datatype": "Float",
		},
		{
			"label": _("Planned Diesel"),
			"value": total_planned_fuel,
			"indicator": "Purple",
			"datatype": "Float",
		},
	]
