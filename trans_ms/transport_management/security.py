from __future__ import unicode_literals

import frappe
from frappe.permissions import add_permission, update_permission_property

ROLE_DEFINITIONS = {
    "Transport Manager": {
        "desk_access": 1,
        "disabled": 0,
        "home_page": "/app/transport",
        "is_custom": 1,
        "two_factor_auth": 0,
    },
    "Transport Operator": {
        "desk_access": 1,
        "disabled": 0,
        "home_page": "/app/transport",
        "is_custom": 1,
        "two_factor_auth": 0,
    },
}

READ_ONLY_RIGHTS = {"select", "read", "report", "export", "print", "email"}
EDIT_RIGHTS = READ_ONLY_RIGHTS | {"create", "write", "share"}
MANAGE_RIGHTS = EDIT_RIGHTS | {"delete"}
SETTINGS_RIGHTS = {"select", "read", "write", "print", "email", "report"}

DOCTYPE_ROLE_RULES = {
    "Transportation Order": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": EDIT_RIGHTS,
    },
    "Vehicle Trip": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": EDIT_RIGHTS,
    },
    "Fuel Request": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": EDIT_RIGHTS,
    },
    "Requested Payments": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": EDIT_RIGHTS,
    },
    "Vehicle Inspection": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": EDIT_RIGHTS,
    },
    "Trip Route": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
    "Trip Location": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
    "Trip Location Type": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
    "Transport Location": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
    "Transport Cargo Type": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
    "Fixed Expense": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
    },
    "Expense": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
    },
    "Trailer": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
    "Vehicle Inspection Template": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
    "Vehicle Type": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
    "Vehicle Routine Checklist": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
    "Vehicle Axle Type": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
    "Vehicle Documents Type": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
    "Driver Documents": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
    "Transport Settings": {
        "Fleet Manager": SETTINGS_RIGHTS,
        "Transport Manager": SETTINGS_RIGHTS,
    },
    "Vehicle Log": {
        "Fleet Manager": MANAGE_RIGHTS,
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
    "Vehicle": {
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
    "Driver": {
        "Transport Manager": MANAGE_RIGHTS,
        "Transport Operator": READ_ONLY_RIGHTS,
    },
}

REPORT_ROLE_RULES = {
    "Daily Trip Schedule": [
        "System Manager",
        "Fleet Manager",
        "Transport Manager",
        "Transport Operator",
    ],
    "Vehicle Tracking Report": [
        "System Manager",
        "Fleet Manager",
        "Transport Manager",
        "Transport Operator",
    ],
    "Document Expiry Alert Report": [
        "System Manager",
        "Fleet Manager",
        "Transport Manager",
        "Transport Operator",
    ],
    "Driver Document Expiry Alert Report": [
        "System Manager",
        "Fleet Manager",
        "Transport Manager",
        "Transport Operator",
    ],
}

WORKSPACE_ROLE_RULES = {
    "Transport": [
        "System Manager",
        "Fleet Manager",
        "Transport Manager",
        "Transport Operator",
    ]
}

PERMISSION_FIELDS = (
    "select",
    "read",
    "write",
    "create",
    "delete",
    "submit",
    "cancel",
    "amend",
    "print",
    "email",
    "report",
    "import",
    "export",
    "share",
)


def sync_transport_security():
    ensure_transport_roles()
    sync_doctype_permissions()
    sync_report_roles()
    sync_workspace_roles()
    frappe.clear_cache()


def ensure_transport_roles():
    for role_name, values in ROLE_DEFINITIONS.items():
        if frappe.db.exists("Role", role_name):
            role = frappe.get_doc("Role", role_name)
            changed = False
            for fieldname, value in values.items():
                if role.get(fieldname) != value:
                    role.set(fieldname, value)
                    changed = True
            if changed:
                role.save(ignore_permissions=True)
            continue

        role = frappe.get_doc({"doctype": "Role", "role_name": role_name, **values})
        role.insert(ignore_permissions=True)


def sync_doctype_permissions():
    for doctype, role_rules in DOCTYPE_ROLE_RULES.items():
        for role, rights in role_rules.items():
            ensure_doctype_permission(doctype, role, rights)


def ensure_doctype_permission(doctype, role, rights, permlevel=0):
    existing = frappe.db.get_value(
        "Custom DocPerm",
        {"parent": doctype, "role": role, "permlevel": permlevel, "if_owner": 0},
    )
    if not existing:
        seed_right = "read" if "read" in rights else next(iter(rights))
        add_permission(doctype, role, permlevel, seed_right)

    for fieldname in PERMISSION_FIELDS:
        update_permission_property(
            doctype,
            role,
            permlevel,
            fieldname,
            1 if fieldname in rights else 0,
            validate=False,
        )


def sync_report_roles():
    for report_name, roles in REPORT_ROLE_RULES.items():
        if not frappe.db.exists("Report", report_name):
            continue

        report = frappe.get_doc("Report", report_name)
        current_roles = {row.role for row in report.roles}
        changed = False

        for role in roles:
            if role in current_roles:
                continue
            report.append("roles", {"role": role})
            changed = True

        if changed:
            report.flags.ignore_links = True
            report.save(ignore_permissions=True)


def sync_workspace_roles():
    for workspace_name, roles in WORKSPACE_ROLE_RULES.items():
        if not frappe.db.exists("Workspace", workspace_name):
            continue

        workspace = frappe.get_doc("Workspace", workspace_name)
        current_roles = {row.role for row in workspace.roles}
        changed = False

        for role in roles:
            if role in current_roles:
                continue
            workspace.append("roles", {"role": role})
            changed = True

        if changed:
            workspace.flags.ignore_links = True
            workspace.save(ignore_permissions=True)
