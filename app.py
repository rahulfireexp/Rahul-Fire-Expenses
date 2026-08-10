"""
Rahul Fire Data Entry - Google Sheets Edition
Internet-accessible Flask app using a Google Sheet as the permanent,
free, never-wiped data store. See README_SETUP.md for full setup steps.
"""

import os
import json
import hashlib
import secrets
from functools import wraps
from datetime import datetime, timedelta
from flask import Flask, request, redirect, url_for, send_file, session, render_template_string
from markupsafe import escape

APP_TITLE = "Rahul Fire Data Entry"

MAX_SITES = 20
MAX_EMPLOYEES = 20
PAYMENT_MODES = ["Cash", "Bank Transfer", "UPI", "Cheque", "Credit Card", "Debit Card"]
BILLABLE_OPTIONS = ["Billable", "Non-Billable"]

PURCHASE_HEADERS = ["id","entry_date","purchaser","vendor","invoice_no","amount",
                     "payment_mode","payment_detail","payment_date","site_name","challan_number","notes"]
CHALLAN_HEADERS = ["id","challan_number","challan_date","site_name","vehicle_number","driver_name"]
ITEM_HEADERS = ["id","challan_id","sr_no","description","qty_nos","qty_meters","billable"]

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", secrets.token_hex(32))
app.permanent_session_lifetime = timedelta(days=30)

# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------
APP_USERNAME = os.environ.get("APP_USERNAME", "admin")
APP_PASSWORD_HASH = os.environ.get("APP_PASSWORD_HASH", "")


def hash_password(password, salt=None):
    if salt is None:
        salt = secrets.token_hex(16)
    h = hashlib.sha256((salt + password).encode()).hexdigest()
    return f"{salt}${h}"


def verify_password(password, stored):
    try:
        salt, h = stored.split("$")
        return hashlib.sha256((salt + password).encode()).hexdigest() == h
    except Exception:
        return False


LOGIN_PAGE = """
<!DOCTYPE html><html><head><title>Login - """ + APP_TITLE + """</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;background:#f7f7f7;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;}
.box{background:#fff;padding:30px;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.15);width:320px;}
h2{color:#7A1F1F;text-align:center;}
input{width:100%;padding:8px;margin:8px 0;border:1px solid #ccc;border-radius:4px;box-sizing:border-box;}
button{width:100%;padding:10px;background:#7A1F1F;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:15px;}
.error{color:#b00020;text-align:center;margin-bottom:10px;}
.warn{color:#a05a00;text-align:center;margin-bottom:10px;font-size:12px;}
</style></head><body>
<div class="box">
  <h2>""" + APP_TITLE + """</h2>
  {% if not configured %}<p class="warn">No password configured yet. Set APP_USERNAME and APP_PASSWORD_HASH
  environment variables on your hosting platform.</p>{% endif %}
  {% if error %}<p class="error">{{ error }}</p>{% endif %}
  <form method="post">
    <input type="text" name="username" placeholder="Username" required autofocus>
    <input type="password" name="password" placeholder="Password" required>
    <button type="submit">Login</button>
  </form>
</div></body></html>
"""


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    configured = bool(APP_PASSWORD_HASH)
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        if configured and username == APP_USERNAME and verify_password(password, APP_PASSWORD_HASH):
            session["logged_in"] = True
            session.permanent = True
            return redirect(url_for("dashboard"))
        error = "Invalid username or password."
    return render_template_string(LOGIN_PAGE, error=error, configured=configured)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper


# ---------------------------------------------------------------------------
# Google Sheets connection (gspread)
# ---------------------------------------------------------------------------
_gs_client = None
_sheet = None


def get_sheet():
    """Lazily connect to Google Sheets using service account credentials
    stored in the GOOGLE_CREDENTIALS_JSON environment variable."""
    global _gs_client, _sheet
    if _sheet is not None:
        return _sheet

    import gspread
    from google.oauth2.service_account import Credentials

    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    sheet_name = os.environ.get("GOOGLE_SHEET_NAME", "Rahul Fire Data Entry - Database")
    if not creds_json:
        raise RuntimeError("GOOGLE_CREDENTIALS_JSON environment variable is not set.")

    creds_dict = json.loads(creds_json)
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    _gs_client = gspread.authorize(creds)
    _sheet = _gs_client.open(sheet_name)
    return _sheet


def get_ws(tab_name):
    return get_sheet().worksheet(tab_name)


def read_all_records(tab_name):
    """Returns list of dicts, one per data row, using header row as keys."""
    ws = get_ws(tab_name)
    return ws.get_all_records()


def next_id(tab_name):
    records = read_all_records(tab_name)
    ids = [int(r["id"]) for r in records if str(r.get("id", "")).strip().isdigit()]
    return (max(ids) + 1) if ids else 1


def append_row(tab_name, headers, row_dict):
    ws = get_ws(tab_name)
    row = [row_dict.get(h, "") for h in headers]
    ws.append_row(row, value_input_option="USER_ENTERED")


def find_row_index_by_id(tab_name, row_id):
    """Returns the 1-based sheet row index (including header) for a given id, or None."""
    ws = get_ws(tab_name)
    ids_col = ws.col_values(1)  # 'id' is always column A
    for idx, val in enumerate(ids_col, start=1):
        if idx == 1:
            continue
        if str(val).strip() == str(row_id):
            return idx
    return None


def update_row_by_id(tab_name, headers, row_id, row_dict):
    ws = get_ws(tab_name)
    idx = find_row_index_by_id(tab_name, row_id)
    if idx is None:
        raise ValueError(f"Row with id={row_id} not found in {tab_name}")
    row = [row_dict.get(h, "") for h in headers]
    ws.update(f"A{idx}:{chr(64+len(headers))}{idx}", [row])


def delete_row_by_id(tab_name, row_id):
    ws = get_ws(tab_name)
    idx = find_row_index_by_id(tab_name, row_id)
    if idx is not None:
        ws.delete_rows(idx)


# ---------------------------------------------------------------------------
# Sites & Employees
# ---------------------------------------------------------------------------
def get_sites():
    return [r["name"] for r in read_all_records("Sites") if r.get("name")]


def get_employees():
    return [r["name"] for r in read_all_records("Employees") if r.get("name")]


def add_site(name):
    sites = get_sites()
    if len(sites) >= MAX_SITES:
        return False, f"Maximum of {MAX_SITES} sites reached."
    if name in sites:
        return False, "A site with that name already exists."
    get_ws("Sites").append_row([name])
    return True, "Site added."


def rename_site(old_name, new_name):
    ws = get_ws("Sites")
    names = ws.col_values(1)
    for idx, val in enumerate(names, start=1):
        if val == old_name:
            ws.update_cell(idx, 1, new_name)
            break
    # Cascade rename into Purchases and Challans
    for tab, headers, col in [("Purchases", PURCHASE_HEADERS, "site_name"),
                               ("Challans", CHALLAN_HEADERS, "site_name")]:
        records = read_all_records(tab)
        ws2 = get_ws(tab)
        col_idx = headers.index(col) + 1
        for i, rec in enumerate(records, start=2):
            if rec.get(col) == old_name:
                ws2.update_cell(i, col_idx, new_name)


def delete_site(name, force=False):
    purchases = [r for r in read_all_records("Purchases") if r.get("site_name") == name]
    challans = [r for r in read_all_records("Challans") if r.get("site_name") == name]
    if (purchases or challans) and not force:
        return False, (f"'{name}' has {len(purchases)} purchase(s) and {len(challans)} challan(s) linked. "
                        f"Use Force Delete to remove it anyway.")
    ws = get_ws("Sites")
    names = ws.col_values(1)
    for idx, val in enumerate(names, start=1):
        if val == name:
            ws.delete_rows(idx)
            break
    return True, f"Site '{name}' deleted."


def add_employee(name):
    emps = get_employees()
    if len(emps) >= MAX_EMPLOYEES:
        return False, f"Maximum of {MAX_EMPLOYEES} employees reached."
    if name in emps:
        return False, "An employee with that name already exists."
    get_ws("Employees").append_row([name])
    return True, "Employee added."


def rename_employee(old_name, new_name):
    ws = get_ws("Employees")
    names = ws.col_values(1)
    for idx, val in enumerate(names, start=1):
        if val == old_name:
            ws.update_cell(idx, 1, new_name)
            break
    records = read_all_records("Purchases")
    ws2 = get_ws("Purchases")
    col_idx = PURCHASE_HEADERS.index("purchaser") + 1
    for i, rec in enumerate(records, start=2):
        if rec.get("purchaser") == old_name:
            ws2.update_cell(i, col_idx, new_name)


def delete_employee(name):
    ws = get_ws("Employees")
    names = ws.col_values(1)
    for idx, val in enumerate(names, start=1):
        if val == name:
            ws.delete_rows(idx)
            break
    return True, f"Employee '{name}' deleted."


# ---------------------------------------------------------------------------
# Purchases
# ---------------------------------------------------------------------------
def get_purchases(site_name=None):
    records = read_all_records("Purchases")
    if site_name and site_name != "ALL SITES":
        records = [r for r in records if r.get("site_name") == site_name]
    records.sort(key=lambda r: (str(r.get("entry_date", "")), int(r.get("id", 0)) if str(r.get("id","")).isdigit() else 0))
    return records


def check_duplicate_purchase(entry_date, vendor, invoice_no, amount, exclude_id=None):
    for r in read_all_records("Purchases"):
        if (str(r.get("entry_date")) == entry_date and r.get("vendor") == vendor and
            r.get("invoice_no") == invoice_no and float(r.get("amount") or 0) == float(amount)):
            if exclude_id is not None and str(r.get("id")) == str(exclude_id):
                continue
            return True
    return False


def insert_purchase(record):
    row_id = next_id("Purchases")
    record["id"] = row_id
    append_row("Purchases", PURCHASE_HEADERS, record)
    return row_id


def update_purchase(row_id, record):
    record["id"] = row_id
    update_row_by_id("Purchases", PURCHASE_HEADERS, row_id, record)


def delete_purchase(row_id):
    delete_row_by_id("Purchases", row_id)


def get_purchase_row(row_id):
    for r in read_all_records("Purchases"):
        if str(r.get("id")) == str(row_id):
            return r
    return None


# ---------------------------------------------------------------------------
# Challans
# ---------------------------------------------------------------------------
import re as _re


def _numeric_key(challan_no):
    m = _re.search(r"\d+", str(challan_no or ""))
    return int(m.group()) if m else 0


def challan_date_mismatch(challan_number, challan_date, exclude_id=None):
    for r in read_all_records("Challans"):
        if r.get("challan_number") == challan_number:
            if exclude_id is not None and str(r.get("id")) == str(exclude_id):
                continue
            if str(r.get("challan_date")) != challan_date:
                return r.get("challan_date")
    return None


def get_or_create_challan(challan_number, challan_date, site_name, vehicle_number, driver_name):
    records = read_all_records("Challans")
    for r in records:
        if r.get("challan_number") == challan_number:
            update_row_by_id("Challans", CHALLAN_HEADERS, r["id"], {
                "id": r["id"], "challan_number": challan_number, "challan_date": challan_date,
                "site_name": site_name, "vehicle_number": vehicle_number, "driver_name": driver_name})
            return r["id"]
    row_id = next_id("Challans")
    append_row("Challans", CHALLAN_HEADERS, {
        "id": row_id, "challan_number": challan_number, "challan_date": challan_date,
        "site_name": site_name, "vehicle_number": vehicle_number, "driver_name": driver_name})
    return row_id


def update_challan_header(challan_id, challan_date, site_name, vehicle_number, driver_name):
    challan = get_challan(challan_id)
    update_row_by_id("Challans", CHALLAN_HEADERS, challan_id, {
        "id": challan_id, "challan_number": challan["challan_number"], "challan_date": challan_date,
        "site_name": site_name, "vehicle_number": vehicle_number, "driver_name": driver_name})


def get_challan(challan_id):
    for r in read_all_records("Challans"):
        if str(r.get("id")) == str(challan_id):
            return r
    return None


def get_all_challans():
    records = read_all_records("Challans")
    records.sort(key=lambda r: (str(r.get("challan_date","")), _numeric_key(r.get("challan_number")), str(r.get("challan_number",""))))
    return records


def delete_challan(challan_id):
    delete_row_by_id("Challans", challan_id)
    for it in read_all_records("ChallanItems"):
        if str(it.get("challan_id")) == str(challan_id):
            delete_row_by_id("ChallanItems", it["id"])


def item_duplicate_in_challan(challan_id, description, qty_nos, qty_meters, exclude_item_id=None):
    for it in read_all_records("ChallanItems"):
        if (str(it.get("challan_id")) == str(challan_id) and it.get("description") == description and
            float(it.get("qty_nos") or 0) == float(qty_nos) and float(it.get("qty_meters") or 0) == float(qty_meters)):
            if exclude_item_id is not None and str(it.get("id")) == str(exclude_item_id):
                continue
            return True
    return False


def next_sr_no(challan_id):
    items = [it for it in read_all_records("ChallanItems") if str(it.get("challan_id")) == str(challan_id)]
    srs = [int(it["sr_no"]) for it in items if str(it.get("sr_no","")).isdigit()]
    return (max(srs) + 1) if srs else 1


def insert_challan_item(challan_id, description, qty_nos, qty_meters, billable):
    sr = next_sr_no(challan_id)
    if sr > 35:
        raise ValueError("This challan already has 35 items (maximum allowed).")
    row_id = next_id("ChallanItems")
    append_row("ChallanItems", ITEM_HEADERS, {
        "id": row_id, "challan_id": challan_id, "sr_no": sr, "description": description,
        "qty_nos": qty_nos, "qty_meters": qty_meters, "billable": billable})
    return sr


def update_challan_item(item_id, description, qty_nos, qty_meters, billable):
    item = get_challan_item(item_id)
    update_row_by_id("ChallanItems", ITEM_HEADERS, item_id, {
        "id": item_id, "challan_id": item["challan_id"], "sr_no": item["sr_no"],
        "description": description, "qty_nos": qty_nos, "qty_meters": qty_meters, "billable": billable})


def delete_challan_item(item_id):
    delete_row_by_id("ChallanItems", item_id)


def get_challan_item(item_id):
    for it in read_all_records("ChallanItems"):
        if str(it.get("id")) == str(item_id):
            return it
    return None


def get_items_for_challan(challan_id):
    items = [it for it in read_all_records("ChallanItems") if str(it.get("challan_id")) == str(challan_id)]
    items.sort(key=lambda r: int(r.get("sr_no", 0)) if str(r.get("sr_no","")).isdigit() else 0)
    return items


def get_challan_items_for_site(site_name=None):
    challans = {str(c["id"]): c for c in read_all_records("Challans")}
    items = read_all_records("ChallanItems")
    result = []
    for it in items:
        c = challans.get(str(it.get("challan_id")))
        if not c:
            continue
        if site_name and site_name != "ALL SITES" and c.get("site_name") != site_name:
            continue
        merged = dict(it)
        merged["challan_number"] = c.get("challan_number")
        merged["challan_date"] = c.get("challan_date")
        merged["site_name"] = c.get("site_name")
        merged["vehicle_number"] = c.get("vehicle_number")
        merged["driver_name"] = c.get("driver_name")
        result.append(merged)
    result.sort(key=lambda r: (str(r.get("challan_date","")), _numeric_key(r.get("challan_number")),
                                 str(r.get("challan_number","")), int(r.get("sr_no",0)) if str(r.get("sr_no","")).isdigit() else 0))
    return result


def get_site_summary():
    sites = get_sites()
    purchases = read_all_records("Purchases")
    challans = read_all_records("Challans")
    summary = []
    for s in sites:
        total = sum(float(p.get("amount") or 0) for p in purchases if p.get("site_name") == s)
        n_exp = sum(1 for p in purchases if p.get("site_name") == s)
        n_ch = len({c.get("challan_number") for c in challans if c.get("site_name") == s})
        summary.append({"name": s, "total_expense": total, "n_expenses": n_exp, "n_challans": n_ch})
    return summary


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------
def _safe_sheet_name(name, used):
    base = "Site - " + name
    base = base[:31] if len(base) > 31 else base
    candidate = base
    n = 2
    while candidate in used:
        suffix = f" ({n})"
        candidate = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate


def build_export_workbook():
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="7A1F1F", end_color="7A1F1F", fill_type="solid")
    sub_font = Font(bold=True, color="1F4E78", size=12)
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def style_header_row(ws, row_idx=1):
        for cell in ws[row_idx]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border

    def autosize(ws):
        for col_cells in ws.columns:
            length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 10), 40)

    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Purchase Entries"
    ws1.append(["Date","Purchaser/Employee","Vendor/Details","Invoice No.","Amount (with GST)",
                "Payment Mode","Payment Detail","Payment Date","Site Name","Challan Number","Notes"])
    for r in get_purchases():
        ws1.append([r.get("entry_date"), r.get("purchaser"), r.get("vendor"), r.get("invoice_no"),
                    r.get("amount"), r.get("payment_mode"), r.get("payment_detail"), r.get("payment_date"),
                    r.get("site_name"), r.get("challan_number"), r.get("notes")])
    style_header_row(ws1); autosize(ws1); ws1.freeze_panes = "A2"

    ws2 = wb.create_sheet("Challan Items")
    ws2.append(["Challan Number","Date","Sr No.","Description of Item","Qty (Nos)","Qty (Meters)",
                "Billable (Y/N)","Site Name","Vehicle Number","Driver Name"])
    for r in get_challan_items_for_site():
        ws2.append([r.get("challan_number"), r.get("challan_date"), r.get("sr_no"), r.get("description"),
                    r.get("qty_nos"), r.get("qty_meters"), r.get("billable"), r.get("site_name"),
                    r.get("vehicle_number"), r.get("driver_name")])
    style_header_row(ws2); autosize(ws2); ws2.freeze_panes = "A2"

    used_names = {"Purchase Entries", "Challan Items"}
    for site in get_sites():
        sh = wb.create_sheet(_safe_sheet_name(site, used_names))
        purchases = get_purchases(site)
        items = get_challan_items_for_site(site)
        total = sum(float(p.get("amount") or 0) for p in purchases)
        n_challans = len({it.get("challan_number") for it in items})

        sh.append([f"SITE: {site}"]); sh["A1"].font = Font(bold=True, size=14, color="7A1F1F")
        sh.append([f"Total Purchase/Expense: {total:,.2f}    |    Unique Challans: {n_challans}"])
        sh["A2"].font = sub_font
        sh.append([])
        sh.append(["-- PURCHASE / EXPENSE ENTRIES --"]); sh[sh.max_row][0].font = sub_font
        header_row_1 = sh.max_row + 1
        sh.append(["Date","Purchaser/Employee","Vendor/Details","Invoice No.","Amount (with GST)",
                    "Payment Mode","Payment Detail","Payment Date","Challan Number","Notes"])
        for p in purchases:
            sh.append([p.get("entry_date"), p.get("purchaser"), p.get("vendor"), p.get("invoice_no"),
                       p.get("amount"), p.get("payment_mode"), p.get("payment_detail"), p.get("payment_date"),
                       p.get("challan_number"), p.get("notes")])
        style_header_row(sh, header_row_1)
        sh.append([])
        sh.append(["-- CHALLAN ITEMS --"]); sh[sh.max_row][0].font = sub_font
        header_row_2 = sh.max_row + 1
        sh.append(["Challan Number","Date","Sr No.","Description of Item","Qty (Nos)","Qty (Meters)",
                    "Billable (Y/N)","Vehicle Number","Driver Name"])
        for it in items:
            sh.append([it.get("challan_number"), it.get("challan_date"), it.get("sr_no"), it.get("description"),
                       it.get("qty_nos"), it.get("qty_meters"), it.get("billable"), it.get("vehicle_number"), it.get("driver_name")])
        style_header_row(sh, header_row_2)
        autosize(sh)

    ws_sum = wb.create_sheet("Site Summary")
    ws_sum.append(["Site","Total Expense","No. of Purchase Entries","No. of Unique Challans"])
    for r in get_site_summary():
        ws_sum.append([r["name"], r["total_expense"], r["n_expenses"], r["n_challans"]])
    style_header_row(ws_sum); autosize(ws_sum)

    ws_sites = wb.create_sheet("Sites")
    ws_sites.append(["Site Name"])
    for s in get_sites():
        ws_sites.append([s])
    style_header_row(ws_sites); autosize(ws_sites)

    ws_emps = wb.create_sheet("Employees")
    ws_emps.append(["Employee Name"])
    for e in get_employees():
        ws_emps.append([e])
    style_header_row(ws_emps); autosize(ws_emps)

    return wb


# ---------------------------------------------------------------------------
# HTML layout helpers
# ---------------------------------------------------------------------------
NAV = f"""
<div style="background:#7A1F1F;padding:12px 20px;">
  <a href="/" style="color:#fff;font-weight:bold;text-decoration:none;margin-right:20px;">{APP_TITLE}</a>
  <a href="/purchase" style="color:#fff;text-decoration:none;margin-right:16px;">Purchase</a>
  <a href="/challans" style="color:#fff;text-decoration:none;margin-right:16px;">Challans</a>
  <a href="/sites" style="color:#fff;text-decoration:none;margin-right:16px;">Sites</a>
  <a href="/employees" style="color:#fff;text-decoration:none;margin-right:16px;">Employees</a>
  <a href="/site_view" style="color:#fff;text-decoration:none;margin-right:16px;">Site-wise View</a>
  <a href="/export" style="color:#fff;text-decoration:none;margin-right:16px;">Export to Excel</a>
  <a href="/logout" style="color:#ffd6d6;text-decoration:none;float:right;">Logout</a>
</div>
"""

STYLE = """
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f7f7f7;color:#222;}
.container{padding:20px;max-width:1400px;margin:auto;}
h2{color:#7A1F1F;}
table{border-collapse:collapse;width:100%;background:#fff;margin-top:10px;}
th,td{border:1px solid #ddd;padding:6px 10px;font-size:13px;text-align:left;}
th{background:#7A1F1F;color:#fff;}
tr:nth-child(even){background:#fafafa;}
form.inline{display:inline;}
.card{background:#fff;padding:16px;border-radius:6px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,0.1);}
input,select,textarea{padding:6px;margin:4px 0;border:1px solid #ccc;border-radius:4px;font-family:inherit;}
button{background:#7A1F1F;color:#fff;border:none;padding:7px 14px;border-radius:4px;cursor:pointer;}
button.secondary{background:#888;}
button.danger{background:#b00020;}
button.edit{background:#1F6F4A;}
.msg{padding:10px;background:#e9fdf0;border:1px solid #7ac98e;border-radius:4px;margin-bottom:10px;}
.msg.error{background:#fde9e9;border-color:#e08a8a;}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:8px;}
.actions{white-space:nowrap;}
</style>
"""


def page(title, body, msg=None, msg_type="ok"):
    msg_html = f'<div class="msg {"error" if msg_type=="error" else ""}">{escape(msg)}</div>' if msg else ""
    return f"""<!DOCTYPE html><html><head><title>{title} - {APP_TITLE}</title>{STYLE}</head>
    <body>{NAV}<div class="container"><h2>{title}</h2>{msg_html}{body}</div></body></html>"""


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route("/")
@login_required
def dashboard():
    summary = get_site_summary()
    rows = "".join(f"<tr><td>{escape(r['name'])}</td><td>{r['total_expense']:,.2f}</td>"
                    f"<td>{r['n_expenses']}</td><td>{r['n_challans']}</td></tr>" for r in summary)
    body = f"""<div class="card"><p>Data is stored in your Google Sheet and auto-saved on every change.
    Sorted by Date (Purchase) and Date &rarr; Challan No. &rarr; Sr No. (Challans).</p></div>
    <table><tr><th>Site</th><th>Total Expense</th><th>Purchase Entries</th><th>Unique Challans</th></tr>{rows}</table>"""
    return page("Dashboard", body)


@app.route("/sites", methods=["GET", "POST"])
@login_required
def sites_page():
    msg, msg_type = None, "ok"
    if request.method == "POST":
        action = request.form.get("action")
        if action == "add":
            ok, msg = add_site(request.form.get("name", "").strip())
            msg_type = "ok" if ok else "error"
        elif action == "rename":
            rename_site(request.form.get("old_name"), request.form.get("new_name", "").strip())
            msg = "Site renamed everywhere it is used."
        elif action == "delete":
            ok, msg = delete_site(request.form.get("name"), force=request.form.get("force") == "1")
            msg_type = "ok" if ok else "error"

    sites = get_sites()
    rows = ""
    for s in sites:
        rows += f"""<tr><td>{escape(s)}</td>
          <td><form class="inline" method="post">
            <input type="hidden" name="action" value="rename">
            <input type="hidden" name="old_name" value="{escape(s)}">
            <input type="text" name="new_name" placeholder="New name" required>
            <button type="submit" class="edit">Rename</button></form></td>
          <td class="actions">
            <form class="inline" method="post" onsubmit="return confirm('Delete {escape(s)}?');">
              <input type="hidden" name="action" value="delete"><input type="hidden" name="name" value="{escape(s)}">
              <button type="submit" class="danger">Delete</button></form>
            <form class="inline" method="post" onsubmit="return confirm('Force delete {escape(s)}?');">
              <input type="hidden" name="action" value="delete"><input type="hidden" name="name" value="{escape(s)}">
              <input type="hidden" name="force" value="1"><button type="submit" class="danger">Force Delete</button></form>
          </td></tr>"""
    body = f"""<div class="card"><h3>Add New Site ({len(sites)}/{MAX_SITES} used)</h3>
    <form method="post"><input type="hidden" name="action" value="add">
    <input type="text" name="name" placeholder="New site name" required>
    <button type="submit">Add Site</button></form></div>
    <table><tr><th>Site Name</th><th>Rename</th><th>Delete</th></tr>{rows}</table>"""
    return page("Manage Sites", body, msg, msg_type)


@app.route("/employees", methods=["GET", "POST"])
@login_required
def employees_page():
    msg, msg_type = None, "ok"
    if request.method == "POST":
        action = request.form.get("action")
        if action == "add":
            ok, msg = add_employee(request.form.get("name", "").strip())
            msg_type = "ok" if ok else "error"
        elif action == "rename":
            rename_employee(request.form.get("old_name"), request.form.get("new_name", "").strip())
            msg = "Employee renamed."
        elif action == "delete":
            ok, msg = delete_employee(request.form.get("name"))

    employees = get_employees()
    rows = ""
    for e in employees:
        rows += f"""<tr><td>{escape(e)}</td>
          <td><form class="inline" method="post">
            <input type="hidden" name="action" value="rename">
            <input type="hidden" name="old_name" value="{escape(e)}">
            <input type="text" name="new_name" placeholder="New name" required>
            <button type="submit" class="edit">Rename</button></form></td>
          <td><form class="inline" method="post" onsubmit="return confirm('Delete {escape(e)}?');">
            <input type="hidden" name="action" value="delete"><input type="hidden" name="name" value="{escape(e)}">
            <button type="submit" class="danger">Delete</button></form></td></tr>"""
    body = f"""<div class="card"><h3>Add New Employee ({len(employees)}/{MAX_EMPLOYEES} used)</h3>
    <form method="post"><input type="hidden" name="action" value="add">
    <input type="text" name="name" placeholder="New employee name" required>
    <button type="submit">Add Employee</button></form></div>
    <table><tr><th>Employee Name</th><th>Rename</th><th>Delete</th></tr>{rows}</table>"""
    return page("Manage Employees", body, msg, msg_type)


@app.route("/purchase", methods=["GET", "POST"])
@login_required
def purchase_page():
    msg, msg_type = None, "ok"
    if request.method == "POST":
        action = request.form.get("action")
        if action == "delete":
            delete_purchase(request.form.get("id"))
            msg = "Entry deleted."
        else:
            f = request.form
            date_ = f.get("entry_date", "").strip()
            site = f.get("site_name", "").strip()
            try:
                amount = float(f.get("amount") or 0)
            except ValueError:
                amount = 0.0
            vendor = f.get("vendor", "").strip()
            invoice_no = f.get("invoice_no", "").strip()
            if not date_ or not site:
                msg, msg_type = "Date and Site are required.", "error"
            else:
                if invoice_no and check_duplicate_purchase(date_, vendor, invoice_no, amount) and f.get("confirm_dup") != "1":
                    msg, msg_type = "Possible duplicate found. Submit again to confirm and add anyway.", "error"
                    return page("Purchase Entry", render_purchase_form_and_table(), msg, msg_type)
                insert_purchase({
                    "entry_date": date_, "purchaser": f.get("purchaser","").strip(), "vendor": vendor,
                    "invoice_no": invoice_no, "amount": amount, "payment_mode": f.get("payment_mode","").strip(),
                    "payment_detail": f.get("payment_detail","").strip(), "payment_date": f.get("payment_date","").strip(),
                    "site_name": site, "challan_number": f.get("challan_number","").strip(),
                    "notes": f.get("notes","").strip()})
                msg = "Entry added."
    return page("Purchase Entry", render_purchase_form_and_table(), msg, msg_type)


def render_purchase_form_and_table():
    sites_opts = "".join(f'<option value="{escape(s)}">{escape(s)}</option>' for s in get_sites())
    emp_opts = "".join(f'<option value="{escape(e)}">{escape(e)}</option>' for e in get_employees())
    pm_opts = "".join(f'<option value="{p}">{p}</option>' for p in PAYMENT_MODES)
    today = datetime.today().strftime("%Y-%m-%d")
    form = f"""<div class="card"><h3>New Entry</h3><form method="post">
    <input type="hidden" name="action" value="add"><div class="grid">
    <div>Date<br><input type="date" name="entry_date" value="{today}" required></div>
    <div>Purchaser<br><select name="purchaser">{emp_opts}</select></div>
    <div>Vendor/Details<br><input type="text" name="vendor"></div>
    <div>Invoice No.<br><input type="text" name="invoice_no"></div>
    <div>Amount (with GST)<br><input type="number" step="0.01" name="amount"></div>
    <div>Payment Mode<br><select name="payment_mode">{pm_opts}</select></div>
    <div>Payment Detail<br><input type="text" name="payment_detail"></div>
    <div>Payment Date<br><input type="date" name="payment_date" value="{today}"></div>
    <div>Site Name<br><select name="site_name">{sites_opts}</select></div>
    <div>Challan Number<br><input type="text" name="challan_number"></div>
    <div style="grid-column: span 2;">Notes / Remarks<br><textarea name="notes" rows="2" style="width:100%;"></textarea></div>
    </div><br><button type="submit">Add Entry</button></form></div>"""

    rows = ""
    for r in get_purchases():
        rows += f"""<tr><td>{r.get('id')}</td><td>{escape(str(r.get('entry_date') or ''))}</td>
        <td>{escape(str(r.get('purchaser') or ''))}</td><td>{escape(str(r.get('vendor') or ''))}</td>
        <td>{escape(str(r.get('invoice_no') or ''))}</td><td>{float(r.get('amount') or 0):,.2f}</td>
        <td>{escape(str(r.get('payment_mode') or ''))}</td><td>{escape(str(r.get('payment_detail') or ''))}</td>
        <td>{escape(str(r.get('payment_date') or ''))}</td><td>{escape(str(r.get('site_name') or ''))}</td>
        <td>{escape(str(r.get('challan_number') or ''))}</td><td>{escape(str(r.get('notes') or ''))}</td>
        <td class="actions"><a href="/purchase/{r.get('id')}/edit"><button type="button" class="edit">Edit</button></a>
        <form class="inline" method="post" onsubmit="return confirm('Delete this entry?');">
        <input type="hidden" name="action" value="delete"><input type="hidden" name="id" value="{r.get('id')}">
        <button type="submit" class="danger">Delete</button></form></td></tr>"""
    table = f"""<table><tr><th>ID</th><th>Date</th><th>Purchaser</th><th>Vendor</th><th>Invoice</th>
    <th>Amount</th><th>Mode</th><th>Detail</th><th>Pay Date</th><th>Site</th><th>Challan</th><th>Notes</th><th></th></tr>{rows}</table>"""
    return form + table


@app.route("/purchase/<row_id>/edit", methods=["GET", "POST"])
@login_required
def purchase_edit(row_id):
    row = get_purchase_row(row_id)
    if not row:
        return redirect(url_for("purchase_page"))
    msg, msg_type = None, "ok"
    if request.method == "POST":
        f = request.form
        date_ = f.get("entry_date", "").strip()
        site = f.get("site_name", "").strip()
        try:
            amount = float(f.get("amount") or 0)
        except ValueError:
            amount = 0.0
        if not date_ or not site:
            msg, msg_type = "Date and Site are required.", "error"
        else:
            update_purchase(row_id, {
                "entry_date": date_, "purchaser": f.get("purchaser","").strip(), "vendor": f.get("vendor","").strip(),
                "invoice_no": f.get("invoice_no","").strip(), "amount": amount, "payment_mode": f.get("payment_mode","").strip(),
                "payment_detail": f.get("payment_detail","").strip(), "payment_date": f.get("payment_date","").strip(),
                "site_name": site, "challan_number": f.get("challan_number","").strip(), "notes": f.get("notes","").strip()})
            return redirect(url_for("purchase_page"))
        row = get_purchase_row(row_id)

    sites_opts = "".join(f'<option value="{escape(s)}" {"selected" if s==row.get("site_name") else ""}>{escape(s)}</option>' for s in get_sites())
    emp_opts = "".join(f'<option value="{escape(e)}" {"selected" if e==row.get("purchaser") else ""}>{escape(e)}</option>' for e in get_employees())
    pm_opts = "".join(f'<option value="{p}" {"selected" if p==row.get("payment_mode") else ""}>{p}</option>' for p in PAYMENT_MODES)
    body = f"""<div class="card"><form method="post"><div class="grid">
    <div>Date<br><input type="date" name="entry_date" value="{escape(str(row.get('entry_date') or ''))}" required></div>
    <div>Purchaser<br><select name="purchaser">{emp_opts}</select></div>
    <div>Vendor/Details<br><input type="text" name="vendor" value="{escape(str(row.get('vendor') or ''))}"></div>
    <div>Invoice No.<br><input type="text" name="invoice_no" value="{escape(str(row.get('invoice_no') or ''))}"></div>
    <div>Amount (with GST)<br><input type="number" step="0.01" name="amount" value="{row.get('amount')}"></div>
    <div>Payment Mode<br><select name="payment_mode">{pm_opts}</select></div>
    <div>Payment Detail<br><input type="text" name="payment_detail" value="{escape(str(row.get('payment_detail') or ''))}"></div>
    <div>Payment Date<br><input type="date" name="payment_date" value="{escape(str(row.get('payment_date') or ''))}"></div>
    <div>Site Name<br><select name="site_name">{sites_opts}</select></div>
    <div>Challan Number<br><input type="text" name="challan_number" value="{escape(str(row.get('challan_number') or ''))}"></div>
    <div style="grid-column: span 2;">Notes / Remarks<br><textarea name="notes" rows="2" style="width:100%;">{escape(str(row.get('notes') or ''))}</textarea></div>
    </div><br><button type="submit">Save Changes</button>
    <a href="/purchase"><button type="button" class="secondary">Cancel</button></a></form></div>"""
    return page(f"Edit Purchase Entry #{row_id}", body, msg, msg_type)


@app.route("/challans", methods=["GET", "POST"])
@login_required
def challans_page():
    msg, msg_type = None, "ok"
    if request.method == "POST":
        f = request.form
        challan_no = f.get("challan_number","").strip()
        challan_date = f.get("challan_date","").strip()
        site = f.get("site_name","").strip()
        if not challan_no or not challan_date or not site:
            msg, msg_type = "Challan Number, Date, and Site are required.", "error"
        else:
            mismatch = challan_date_mismatch(challan_no, challan_date)
            if mismatch:
                msg, msg_type = f"Challan '{challan_no}' already exists with date {mismatch}.", "error"
            else:
                cid = get_or_create_challan(challan_no, challan_date, site, f.get("vehicle_number","").strip(), f.get("driver_name","").strip())
                return redirect(url_for("challan_detail", challan_id=cid))

    sites_opts = "".join(f'<option value="{escape(s)}">{escape(s)}</option>' for s in get_sites())
    today = datetime.today().strftime("%Y-%m-%d")
    form = f"""<div class="card"><h3>Start / Load a Challan</h3><form method="post"><div class="grid">
    <div>Challan Number<br><input type="text" name="challan_number" required></div>
    <div>Date<br><input type="date" name="challan_date" value="{today}" required></div>
    <div>Site<br><select name="site_name">{sites_opts}</select></div>
    <div>Vehicle Number<br><input type="text" name="vehicle_number"></div>
    <div>Driver Name<br><input type="text" name="driver_name"></div></div>
    <br><button type="submit">Start / Load Challan</button></form></div>"""

    rows = ""
    for c in get_all_challans():
        n_items = len(get_items_for_challan(c["id"]))
        rows += f"""<tr><td><a href="/challans/{c['id']}">{escape(str(c['challan_number']))}</a></td>
        <td>{escape(str(c['challan_date']))}</td><td>{escape(str(c['site_name']))}</td>
        <td>{escape(str(c.get('vehicle_number') or ''))}</td><td>{escape(str(c.get('driver_name') or ''))}</td><td>{n_items}</td>
        <td class="actions"><a href="/challans/{c['id']}/edit"><button type="button" class="edit">Edit</button></a>
        <form class="inline" method="post" action="/challans/{c['id']}/delete" onsubmit="return confirm('Delete this challan and items?');">
        <button type="submit" class="danger">Delete</button></form></td></tr>"""
    table = f"""<table><tr><th>Challan No.</th><th>Date</th><th>Site</th><th>Vehicle</th><th>Driver</th><th>Items</th><th></th></tr>{rows}</table>"""
    return page("Challans", form + table, msg, msg_type)


@app.route("/challans/<challan_id>/delete", methods=["POST"])
@login_required
def challan_delete_route(challan_id):
    delete_challan(challan_id)
    return redirect(url_for("challans_page"))


@app.route("/challans/<challan_id>/edit", methods=["GET", "POST"])
@login_required
def challan_edit(challan_id):
    challan = get_challan(challan_id)
    if not challan:
        return redirect(url_for("challans_page"))
    msg, msg_type = None, "ok"
    if request.method == "POST":
        f = request.form
        challan_date = f.get("challan_date","").strip()
        site = f.get("site_name","").strip()
        mismatch = challan_date_mismatch(challan["challan_number"], challan_date, exclude_id=challan_id)
        if mismatch:
            msg, msg_type = f"Another challan with the same number already uses date {mismatch}.", "error"
        else:
            update_challan_header(challan_id, challan_date, site, f.get("vehicle_number","").strip(), f.get("driver_name","").strip())
            return redirect(url_for("challan_detail", challan_id=challan_id))
        challan = get_challan(challan_id)

    sites_opts = "".join(f'<option value="{escape(s)}" {"selected" if s==challan["site_name"] else ""}>{escape(s)}</option>' for s in get_sites())
    body = f"""<div class="card"><p><b>Challan Number:</b> {escape(str(challan['challan_number']))} (fixed)</p>
    <form method="post"><div class="grid">
    <div>Date<br><input type="date" name="challan_date" value="{escape(str(challan['challan_date']))}" required></div>
    <div>Site<br><select name="site_name">{sites_opts}</select></div>
    <div>Vehicle Number<br><input type="text" name="vehicle_number" value="{escape(str(challan.get('vehicle_number') or ''))}"></div>
    <div>Driver Name<br><input type="text" name="driver_name" value="{escape(str(challan.get('driver_name') or ''))}"></div>
    </div><br><button type="submit">Save Changes</button>
    <a href="/challans/{challan_id}"><button type="button" class="secondary">Cancel</button></a></form></div>"""
    return page(f"Edit Challan {challan['challan_number']}", body, msg, msg_type)


@app.route("/challans/<challan_id>", methods=["GET", "POST"])
@login_required
def challan_detail(challan_id):
    challan = get_challan(challan_id)
    if not challan:
        return redirect(url_for("challans_page"))
    msg, msg_type = None, "ok"
    if request.method == "POST":
        action = request.form.get("action")
        if action == "delete_item":
            delete_challan_item(request.form.get("item_id"))
            msg = "Item deleted."
        else:
            f = request.form
            desc = f.get("description","").strip()
            try:
                qty_nos = float(f.get("qty_nos") or 0)
                qty_m = float(f.get("qty_meters") or 0)
            except ValueError:
                qty_nos, qty_m = 0.0, 0.0
            billable = f.get("billable","").strip()
            if not desc or billable not in BILLABLE_OPTIONS:
                msg, msg_type = "Description and Billable are required.", "error"
            elif item_duplicate_in_challan(challan_id, desc, qty_nos, qty_m):
                msg, msg_type = "An item with the same Description and Quantity already exists in THIS challan.", "error"
            else:
                try:
                    insert_challan_item(challan_id, desc, qty_nos, qty_m, billable)
                    msg = "Item added."
                except ValueError as e:
                    msg, msg_type = str(e), "error"

    items = get_items_for_challan(challan_id)
    rows = ""
    for it in items:
        rows += f"""<tr><td>{it['sr_no']}</td><td>{escape(str(it['description']))}</td><td>{it['qty_nos']}</td>
        <td>{it['qty_meters']}</td><td>{escape(str(it['billable']))}</td>
        <td class="actions"><a href="/challans/{challan_id}/items/{it['id']}/edit"><button type="button" class="edit">Edit</button></a>
        <form class="inline" method="post" onsubmit="return confirm('Delete this item?');">
        <input type="hidden" name="action" value="delete_item"><input type="hidden" name="item_id" value="{it['id']}">
        <button type="submit" class="danger">Delete</button></form></td></tr>"""

    bill_opts = "".join(f'<option value="{b}">{b}</option>' for b in BILLABLE_OPTIONS)
    body = f"""<div class="card"><p><b>Challan:</b> {escape(str(challan['challan_number']))} &nbsp;
    <b>Date:</b> {escape(str(challan['challan_date']))} &nbsp; <b>Site:</b> {escape(str(challan['site_name']))} &nbsp;
    <b>Vehicle:</b> {escape(str(challan.get('vehicle_number') or '-'))} &nbsp;
    <b>Driver:</b> {escape(str(challan.get('driver_name') or '-'))} &nbsp; ({len(items)}/35 items) &nbsp;
    <a href="/challans/{challan_id}/edit"><button type="button" class="edit">Edit Challan Header</button></a></p></div>
    <div class="card"><h3>Add Item</h3><form method="post"><div class="grid">
    <div>Description<br><input type="text" name="description" required></div>
    <div>Qty (Nos)<br><input type="number" step="0.01" name="qty_nos"></div>
    <div>Qty (Meters)<br><input type="number" step="0.01" name="qty_meters"></div>
    <div>Billable<br><select name="billable">{bill_opts}</select></div></div>
    <br><button type="submit">Add Item</button></form></div>
    <table><tr><th>Sr No.</th><th>Description</th><th>Qty(Nos)</th><th>Qty(M)</th><th>Billable</th><th></th></tr>{rows}</table>
    <p><a href="/challans">&larr; Back to all challans</a></p>"""
    return page(f"Challan {challan['challan_number']}", body, msg, msg_type)


@app.route("/challans/<challan_id>/items/<item_id>/edit", methods=["GET", "POST"])
@login_required
def challan_item_edit(challan_id, item_id):
    item = get_challan_item(item_id)
    if not item:
        return redirect(url_for("challan_detail", challan_id=challan_id))
    msg, msg_type = None, "ok"
    if request.method == "POST":
        f = request.form
        desc = f.get("description","").strip()
        try:
            qty_nos = float(f.get("qty_nos") or 0)
            qty_m = float(f.get("qty_meters") or 0)
        except ValueError:
            qty_nos, qty_m = 0.0, 0.0
        billable = f.get("billable","").strip()
        if not desc or billable not in BILLABLE_OPTIONS:
            msg, msg_type = "Description and Billable are required.", "error"
        elif item_duplicate_in_challan(challan_id, desc, qty_nos, qty_m, exclude_item_id=item_id):
            msg, msg_type = "Another item with the same Description and Quantity already exists in this challan.", "error"
        else:
            update_challan_item(item_id, desc, qty_nos, qty_m, billable)
            return redirect(url_for("challan_detail", challan_id=challan_id))
        item = get_challan_item(item_id)

    bill_opts = "".join(f'<option value="{b}" {"selected" if b==item["billable"] else ""}>{b}</option>' for b in BILLABLE_OPTIONS)
    body = f"""<div class="card"><p><b>Sr No.:</b> {item['sr_no']} (fixed)</p><form method="post"><div class="grid">
    <div>Description<br><input type="text" name="description" value="{escape(str(item['description']))}" required></div>
    <div>Qty (Nos)<br><input type="number" step="0.01" name="qty_nos" value="{item['qty_nos']}"></div>
    <div>Qty (Meters)<br><input type="number" step="0.01" name="qty_meters" value="{item['qty_meters']}"></div>
    <div>Billable<br><select name="billable">{bill_opts}</select></div></div>
    <br><button type="submit">Save Changes</button>
    <a href="/challans/{challan_id}"><button type="button" class="secondary">Cancel</button></a></form></div>"""
    return page(f"Edit Item (Sr No. {item['sr_no']})", body, msg, msg_type)


@app.route("/site_view")
@login_required
def site_view():
    site = request.args.get("site", "ALL SITES")
    sites_opts = "".join(f'<option value="{escape(s)}" {"selected" if s==site else ""}>{escape(s)}</option>' for s in ["ALL SITES"] + get_sites())
    pc_rows = "".join(
        f"<tr><td>{escape(str(r.get('entry_date') or ''))}</td><td>{escape(str(r.get('purchaser') or ''))}</td>"
        f"<td>{escape(str(r.get('vendor') or ''))}</td><td>{escape(str(r.get('invoice_no') or ''))}</td>"
        f"<td>{float(r.get('amount') or 0):,.2f}</td><td>{escape(str(r.get('site_name') or ''))}</td>"
        f"<td>{escape(str(r.get('challan_number') or ''))}</td><td>{escape(str(r.get('notes') or ''))}</td></tr>"
        for r in get_purchases(site))
    ci_rows = "".join(
        f"<tr><td>{escape(str(r.get('challan_number')))}</td><td>{escape(str(r.get('challan_date')))}</td>"
        f"<td>{r.get('sr_no')}</td><td>{escape(str(r.get('description')))}</td><td>{r.get('qty_nos')}</td>"
        f"<td>{r.get('qty_meters')}</td><td>{escape(str(r.get('billable')))}</td><td>{escape(str(r.get('site_name')))}</td></tr>"
        for r in get_challan_items_for_site(site))
    total = sum(float(r.get("amount") or 0) for r in get_purchases(site)) if site != "ALL SITES" else None
    total_html = f"<p><b>Total Expense for {escape(site)}: {total:,.2f}</b></p>" if total is not None else ""
    body = f"""<form method="get"><label>Select Site: </label>
    <select name="site" onchange="this.form.submit()">{sites_opts}</select></form>{total_html}
    <h3>Purchase Entries</h3><table><tr><th>Date</th><th>Purchaser</th><th>Vendor</th><th>Invoice</th>
    <th>Amount</th><th>Site</th><th>Challan</th><th>Notes</th></tr>{pc_rows}</table>
    <h3>Challan Items</h3><table><tr><th>Challan No.</th><th>Date</th><th>Sr No.</th><th>Description</th>
    <th>Qty(Nos)</th><th>Qty(M)</th><th>Billable</th><th>Site</th></tr>{ci_rows}</table>"""
    return page("Site-wise View", body)


@app.route("/export")
@login_required
def export_route():
    wb = build_export_workbook()
    tmp_path = "/tmp/Fire_Safety_Export.xlsx"
    wb.save(tmp_path)
    return send_file(tmp_path, as_attachment=True, download_name="Fire_Safety_Export.xlsx")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
