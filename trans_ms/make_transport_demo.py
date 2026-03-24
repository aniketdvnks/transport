# file: make_transport_demo.py
import frappe
from frappe.utils import nowdate, add_days

def get_or_create(doctype, filters, defaults=None):
    doc = frappe.db.get_value(doctype, filters, "name")
    if doc:
        return frappe.get_doc(doctype, doc)
    data = filters.copy()
    if defaults:
        data.update(defaults)
    d = frappe.get_doc({"doctype": doctype, **data})
    d.insert(ignore_permissions=True)
    return d

def run():
    frappe.flags.mute_emails = True

    # 1. Basic company / accounts context
    company = frappe.db.get_single_value("Global Defaults", "default_company")
    if not company:
        # fall back to first company
        company = frappe.db.get_value("Company", {}, "name")
    if not company:
        frappe.throw("No Company found. Please create a Company first.")

    # Try to pick a default receivable and income account
    receivable = frappe.db.get_value(
        "Account",
        {"company": company, "account_type": "Receivable", "is_group": 0},
        "name",
    )
    income = frappe.db.get_value(
        "Account",
        {"company": company, "root_type": "Income", "is_group": 0},
        "name",
    )

    # 2. Transport Settings (doctype name in app is Transport Management Settings)
    # If the exact doctype is different, adjust name here based on your app
    settings_dt = "Transport Management Settings"
    if frappe.db.exists("DocType", settings_dt):
        ts_name = frappe.db.get_single_value(settings_dt, "name")
        if not ts_name:
            ts = frappe.get_doc({"doctype": settings_dt})
        else:
            ts = frappe.get_doc(settings_dt, ts_name)

        if company and hasattr(ts, "company"):
            ts.company = company
        if receivable and hasattr(ts, "default_receivable_account"):
            ts.default_receivable_account = receivable
        if income and hasattr(ts, "default_income_account"):
            ts.default_income_account = income
        ts.flags.ignore_mandatory = True
        ts.save(ignore_permissions=True)

    # 3. Masters: Vehicle Type, Vehicle, Driver, Route

    # Vehicle Type
    vt = get_or_create(
        "Vehicle Type",
        {"vehicle_type": "TMS-DEMO-TRUCK"},
        {"description": "Demo 20T Truck"},
    )

    # Vehicle
    vehicle = get_or_create(
        "Vehicle",
        {"license_plate": "TMS-DEMO-TRK-01"},
        {
            "vehicle_name": "Demo Truck 01",
            "make": "Ashok Leyland",
            "model": "2518",
            "vehicle_type": vt.name,
            "chassis_no": "CHS-TMS-DEMO-0001",
            "last_odometer":1000,
            "uom":"Nos"
        },
    )

    # Driver (Employee)
    driver = get_or_create(
        "Employee",
        {"employee_name": "Demo Driver"},
        {
            "company": company,
            "status": "Active",
            "employment_type": "Full-time",
            "date_of_birth": "1990-01-01",
            "date_of_joining": nowdate(),
            "first_name": "Demo Driver",
            "gender":"Male"
        },
    )

    # Route
    route = get_or_create(
        "Trip Route",
        {"route_name": "TMS-DEMO-MUM-DEL"},
        {
            "from_location": "Mumbai",
            "to_location": "Delhi",
            "distance_km": 1450,
            "trip_steps":10,
            "total_distance":1450
        },
    )

    # 4. Customer and Contract

    customer = get_or_create(
        "Customer",
        {"customer_name": "TMS Demo Customer"},
        {
            "customer_group": "Commercial" if frappe.db.exists("Customer Group", "Commercial") else "All Customer Groups",
            "territory": "All Territories",
            "company": company,
        },
    )

    # Transport Contract (doctype name might be Transport Contract or similar)
    contract_dt = None
    for dt in ["Transport Contract", "Transportation Contract", "Transport Agreement"]:
        if frappe.db.exists("DocType", dt):
            contract_dt = dt
            break

    contract = None
    if contract_dt:
        contract = get_or_create(
            contract_dt,
            {"title": "TMS-DEMO-CONTRACT-1"},
            {
                "customer": customer.name,
                "company": company,
                "valid_from": nowdate(),
                "valid_upto": add_days(nowdate(), 365),
                "route": route.name if hasattr(frappe.get_meta(contract_dt), "get_field") else None,
            },
        )

    # 5. Transport Order (Sales-like order for a trip)

    order_dt = None
    for dt in ["Transport Order", "Transportation Order"]:
        if frappe.db.exists("DocType", dt):
            order_dt = dt
            break

    order = None
    if order_dt:
        order = frappe.new_doc(order_dt)
        order.customer = customer.name
        if hasattr(order, "company"):
            order.company = company
        if hasattr(order, "contract"):
            order.contract = contract.name if contract else None
        if hasattr(order, "route"):
            order.route = route.name
        if hasattr(order, "from_location"):
            order.from_location = "Mumbai"
        if hasattr(order, "to_location"):
            order.to_location = "Delhi"
        if hasattr(order, "posting_date"):
            order.posting_date = nowdate()
        if hasattr(order, "pickup_date"):
            order.pickup_date = add_days(nowdate(), 1)
        if hasattr(order, "delivery_date"):
            order.delivery_date = add_days(nowdate(), 3)

        # Child table items may differ; try common patterns
        for child_table_field in ["items", "transport_orders", "order_details"]:
            if hasattr(order, child_table_field):
                row = order.append(child_table_field)
                if hasattr(row, "description"):
                    row.description = "Demo Full Truck Load Mumbai-Delhi"
                if hasattr(row, "qty"):
                    row.qty = 1
                if hasattr(row, "uom"):
                    row.uom = "Trip"
                if hasattr(row, "rate"):
                    row.rate = 50000
                if hasattr(row, "amount"):
                    row.amount = 50000
                if hasattr(row, "from_location"):
                    row.from_location = "Mumbai"
                if hasattr(row, "to_location"):
                    row.to_location = "Delhi"
                break

        order.flags.ignore_mandatory = True
        order.insert(ignore_permissions=True)
        if hasattr(order, "submit"):
            try:
                order.submit()
            except Exception:
                frappe.db.rollback()
                order = frappe.get_doc(order_dt, order.name)

    # 6. Trip / Trip Sheet

    trip_dt = None
    for dt in ["Transport Trip", "Trip Sheet", "Vehicle Trip"]:
        if frappe.db.exists("DocType", dt):
            trip_dt = dt
            break

    if trip_dt:
        trip = frappe.new_doc(trip_dt)
        if hasattr(trip, "vehicle"):
            trip.vehicle = vehicle.name
        if hasattr(trip, "driver"):
            trip.driver = driver.name
        if hasattr(trip, "route"):
            trip.route = route.name
        if hasattr(trip, "from_location"):
            trip.from_location = "Mumbai"
        if hasattr(trip, "to_location"):
            trip.to_location = "Delhi"
        if hasattr(trip, "start_date"):
            trip.start_date = add_days(nowdate(), 1)
        if hasattr(trip, "end_date"):
            trip.end_date = add_days(nowdate(), 3)
        if hasattr(trip, "company"):
            trip.company = company

        # Link to order if field exists
        for fld in ["transport_order", "order", "reference_order"]:
            if hasattr(trip, fld) and order:
                setattr(trip, fld, order.name)
                break

        trip.flags.ignore_mandatory = True
        trip.insert(ignore_permissions=True)
        if hasattr(trip, "submit"):
            try:
                trip.submit()
            except Exception:
                frappe.db.rollback()

    frappe.db.commit()
    print("Transport demo data created successfully.")
