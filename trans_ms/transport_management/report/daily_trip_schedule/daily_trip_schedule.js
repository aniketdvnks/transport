// Copyright (c) 2026, contributors
// For license information, please see license.txt
/* eslint-disable */

frappe.query_reports["Daily Trip Schedule"] = {
	"filters": [
		{
			"fieldname": "report_date",
			"label": __("Report Date"),
			"fieldtype": "Date",
			"default": frappe.datetime.get_today(),
			"reqd": 1
		},
		{
			"fieldname": "main_route",
			"label": __("Trip Route"),
			"fieldtype": "Link",
			"options": "Trip Route",
		},
		{
			"fieldname": "tanker",
			"label": __("Tanker"),
			"fieldtype": "Link",
			"options": "Vehicle",
		},
		{
			"fieldname": "stage",
			"label": __("Filled/Stage"),
			"fieldtype": "Data",
		}
	],
	formatter(value, row, column, data, default_formatter) {
		value = default_formatter(value, row, column, data);

		if (!data || value === null || value === undefined || value === "") {
			return value;
		}

		if (column.fieldname && column.fieldname.startsWith("stage_")) {
			const rawValue = (data[column.fieldname] || "").toString().toUpperCase();
			const colorMap = {
				"EM": "#b7791f",
				"LO": "#2f855a",
				"UNL": "#2b6cb0",
				"WORKING": "#6b46c1",
			};
			const color = colorMap[rawValue] || "#4a5568";
			return `<span style="font-weight: 600; color: ${color};">${value}</span>`;
		}

		if (column.fieldname && column.fieldname.startsWith("fuel_gap_")) {
			const numericValue = Number(data[column.fieldname] || 0);
			if (numericValue > 0) {
				return `<span style="font-weight: 600; color: #c53030;">${value}</span>`;
			}
		}

		if (column.fieldname && column.fieldname.startsWith("balance_")) {
			const numericValue = Number(data[column.fieldname] || 0);
			const color = numericValue < 0 ? "#c53030" : numericValue > 0 ? "#2f855a" : "#718096";
			return `<span style="font-weight: 600; color: ${color};">${value}</span>`;
		}

		return value;
	}
};
