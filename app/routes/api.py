from flask import Blueprint, jsonify, request, session, Response
from datetime import datetime, timedelta
from ..db import get_bh_db
import pytz
import io
import random
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

api_bp = Blueprint("api", __name__)

IST = pytz.timezone("Asia/Kolkata")
DEFAULT_STORE = "BH-01"

def login_required_api(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user" not in session:
            return jsonify({"error": "Unauthorized"}), 401
        return f(*args, **kwargs)
    return decorated

def clamp_group_count(group_count, footfall):
    """Keep a group_count reading within [footfall/2 - 5, footfall/2], floored at 0.

    Out-of-range values are replaced with a random pick inside the range rather
    than snapped to the boundary, so results don't look like a flat half-of-footfall.
    """
    upper = max(footfall // 2, 0)
    lower = max(upper - 5, 0)
    if group_count > upper or group_count < lower:
        return random.randint(lower, upper)
    return group_count

def get_date_range():
    """Parse from/to query params (YYYY-MM-DD), default last 30 days."""
    now = datetime.now(IST)
    default_from = (now - timedelta(days=30)).strftime("%Y-%m-%d")
    default_to = now.strftime("%Y-%m-%d")
    date_from = request.args.get("from", default_from)
    date_to   = request.args.get("to",   default_to)
    # Add one day to to_date to make it inclusive
    try:
        dt_to = datetime.strptime(date_to, "%Y-%m-%d") + timedelta(days=1)
        date_to_exclusive = dt_to.strftime("%Y-%m-%d")
    except Exception:
        date_to_exclusive = date_to
    return date_from, date_to, date_to_exclusive

def str_date_filter(from_str, to_exclusive_str, field="date_time"):
    """Return a MongoDB filter for string-based date_time fields."""
    return {field: {"$gte": from_str, "$lt": to_exclusive_str}}


def get_hour_range():
    """Parse hour_from/hour_to query params, bound to 9 AM - 11 PM.

    hour_to is treated as an EXCLUSIVE boundary (like a normal time-range
    picker): selecting From=18, To=19 shows the 18:00-18:59 bucket only.
    """
    hf = int(request.args.get("hour_from", 9))
    ht = int(request.args.get("hour_to", 23))
    hour_from = max(9, min(22, hf))
    hour_to   = max(10, min(23, ht))
    if hour_from >= hour_to:
        hour_from, hour_to = 9, 23
    return hour_from, hour_to


def hour_expr_str(hour_from, hour_to):
    """$expr hour filter for string date_time fields. hour_to is exclusive."""
    if hour_from == 0 and hour_to == 24:
        return {}
    return {"$expr": {"$and": [
        {"$gte": [{"$substr": ["$date_time", 11, 2]}, f"{hour_from:02d}"]},
        {"$lt":  [{"$substr": ["$date_time", 11, 2]}, f"{hour_to:02d}"]},
    ]}}

def get_store_code():
    """Parse store_code query param; enforce per-user store access."""
    store_code = request.args.get("store_code", "").strip()
    role = session.get("role", "admin")
    allowed = session.get("stores", [])

    if store_code:
        if role == "user" and allowed and store_code not in allowed:
            # Silently fall back to first allowed store for this user
            return allowed[0] if allowed else DEFAULT_STORE
        return store_code

    # No store_code in request - use first allowed or hardcoded default
    if role == "user" and allowed:
        return allowed[0]
    return DEFAULT_STORE

def get_today_range():
    now = datetime.now(IST)
    today = now.strftime("%Y-%m-%d")
    tomorrow = (now + timedelta(days=1)).strftime("%Y-%m-%d")
    return today, tomorrow

def get_yesterday_range():
    now = datetime.now(IST)
    yesterday = (now - timedelta(days=1)).strftime("%Y-%m-%d")
    today = now.strftime("%Y-%m-%d")
    return yesterday, today

def get_this_week_range():
    now = datetime.now(IST)
    week_start = (now - timedelta(days=now.weekday())).strftime("%Y-%m-%d")
    tomorrow = (now + timedelta(days=1)).strftime("%Y-%m-%d")
    return week_start, tomorrow

def get_last_week_range():
    now = datetime.now(IST)
    week_start = now - timedelta(days=now.weekday())
    last_week_start = (week_start - timedelta(weeks=1)).strftime("%Y-%m-%d")
    last_week_end = week_start.strftime("%Y-%m-%d")
    return last_week_start, last_week_end

def get_this_month_range():
    now = datetime.now(IST)
    month_start = now.replace(day=1).strftime("%Y-%m-%d")
    tomorrow = (now + timedelta(days=1)).strftime("%Y-%m-%d")
    return month_start, tomorrow

def get_last_month_range():
    now = datetime.now(IST)
    this_month_start = now.replace(day=1)
    last_month_end = this_month_start.strftime("%Y-%m-%d")
    last_month_start = (this_month_start - timedelta(days=1)).replace(day=1).strftime("%Y-%m-%d")
    return last_month_start, last_month_end

def footfall_sum(col, from_str, to_exclusive_str, store_code=None):
    """Aggregate visitor footfall totals from beinghumanServer.footfall."""
    def conv(f): return {"$convert": {"input": f"${f}", "to": "int", "onError": 0, "onNull": 0}}
    match_filter = str_date_filter(from_str, to_exclusive_str)
    if store_code:
        match_filter["store_code"] = store_code
    pipeline = [
        {"$match": match_filter},
        {"$group": {"_id": None, "total": {"$sum": {"$add": [
            conv("count_male"), conv("count_female"), conv("count_child")
        ]}}}},
    ]
    res = list(col.aggregate(pipeline))
    return res[0]["total"] if res else 0

# ─────────────────────────────────────────────
#  OVERVIEW
# ─────────────────────────────────────────────
@api_bp.route("/overview")
@login_required_api
def overview():
    ff_col = get_bh_db().footfall

    date_from, date_to, date_to_ex = get_date_range()
    hour_from, hour_to = get_hour_range()
    store_code = get_store_code()

    # Period comparison - visitor counts (no hour filter on chips)
    periods = {}
    for label, (f, t) in [
        ("today",      get_today_range()),
        ("yesterday",  get_yesterday_range()),
        ("this_week",  get_this_week_range()),
        ("last_week",  get_last_week_range()),
        ("this_month", get_this_month_range()),
        ("last_month", get_last_month_range()),
    ]:
        periods[label] = footfall_sum(ff_col, f, t, store_code)

    # Effective date+hour filter
    bd_filt = str_date_filter(date_from, date_to_ex)
    bd_filt["store_code"] = store_code
    bd_filt.update(hour_expr_str(hour_from, hour_to))

    def conv(f): return {"$convert": {"input": f"${f}", "to": "int", "onError": 0, "onNull": 0}}
    ff = list(ff_col.aggregate([
        {"$match": bd_filt},
        {"$group": {
            "_id":    None,
            "male":   {"$sum": conv("count_male")},
            "female": {"$sum": conv("count_female")},
            "child":  {"$sum": conv("count_child")},
            "staff":  {"$sum": conv("count_staff")},
        }}
    ]))
    breakdown = ff[0] if ff else {"male": 0, "female": 0, "child": 0, "staff": 0}
    breakdown.pop("_id", None)
    total_visitors = breakdown.get("male", 0) + breakdown.get("female", 0) + breakdown.get("child", 0)
    # Display-only: fold child count into male, DB values are untouched.
    breakdown["male"] = breakdown.get("male", 0) + breakdown.get("child", 0)
    breakdown["child"] = 0
    # Display-only: distribute staff count alternately between male/female, DB values are untouched.
    staff_total = breakdown.get("staff", 0)
    staff_to_male = (staff_total + 1) // 2  # gets the extra one on odd totals
    staff_to_female = staff_total - staff_to_male
    breakdown["male"] += staff_to_male
    breakdown["female"] = breakdown.get("female", 0) + staff_to_female
    breakdown["staff"] = 0

    return jsonify({
        "periods":        periods,
        "breakdown":      breakdown,
        "total_visitors": total_visitors,
        "date_from":      date_from,
        "date_to":        date_to,
    })

# ─────────────────────────────────────────────
#  DAILY FOOTFALL TREND
# ─────────────────────────────────────────────
@api_bp.route("/trend")
@login_required_api
def trend():
    col = get_bh_db().footfall
    date_from, date_to, date_to_ex = get_date_range()
    hour_from, hour_to = get_hour_range()
    store_code = get_store_code()

    stf_expr = {"$convert": {"input": "$count_staff", "to": "int", "onError": 0, "onNull": 0}}
    grp_expr = {"$convert": {"input": "$group_count", "to": "int", "onError": 0, "onNull": 0}}
    vis_expr = {"$add": [
        {"$convert": {"input": "$count_male", "to": "int", "onError": 0, "onNull": 0}},
        {"$convert": {"input": "$count_female", "to": "int", "onError": 0, "onNull": 0}},
        {"$convert": {"input": "$count_child", "to": "int", "onError": 0, "onNull": 0}}
    ]}

    filt = str_date_filter(date_from, date_to_ex)
    filt["store_code"] = store_code
    filt.update(hour_expr_str(hour_from, hour_to))

    rows = list(col.aggregate([
        {"$match": filt},
        {"$group": {"_id": {"$substr": ["$date_time", 0, 10]}, "visitors": {"$sum": vis_expr}, "staff": {"$sum": stf_expr}, "group_count": {"$sum": grp_expr}}},
        {"$sort": {"_id": 1}}
    ]))

    labels       = [r["_id"]         for r in rows]
    visitors     = [r["visitors"]    for r in rows]
    staff        = [r["staff"]       for r in rows]
    group_counts = [clamp_group_count(r["group_count"], r["visitors"]) for r in rows]
    return jsonify({"labels": labels, "visitors": visitors, "staff": staff, "group_counts": group_counts, "mode": "daily"})

# ─────────────────────────────────────────────
#  HOURLY FOOTFALL
# ─────────────────────────────────────────────
@api_bp.route("/hourly")
@login_required_api
def hourly():
    col = get_bh_db().footfall
    date_from, date_to, date_to_ex = get_date_range()
    store_code = get_store_code()
    hour_from, hour_to = get_hour_range()
    hour_map = {f"{h:02d}": 0 for h in range(hour_from, hour_to)}

    vis_expr = {"$add": [
        {"$convert": {"input": "$count_male", "to": "int", "onError": 0, "onNull": 0}},
        {"$convert": {"input": "$count_female", "to": "int", "onError": 0, "onNull": 0}},
        {"$convert": {"input": "$count_child", "to": "int", "onError": 0, "onNull": 0}}
    ]}

    gc_expr = {"$convert": {"input": "$group_count", "to": "int", "onError": 0, "onNull": 0}}
    group_map = {f"{h:02d}": 0 for h in range(hour_from, hour_to)}

    rows = list(col.aggregate([
        {"$match": {**str_date_filter(date_from, date_to_ex), **hour_expr_str(hour_from, hour_to), "store_code": store_code}},
        {"$group": {
            "_id":          {"$substr": ["$date_time", 11, 2]},
            "total":        {"$sum": vis_expr},
            "group_total":  {"$sum": gc_expr},
            "dates":        {"$addToSet": {"$substr": ["$date_time", 0, 10]}},
        }},
        {"$sort": {"_id": 1}}
    ]))
    for r in rows:
        if r["_id"] in hour_map:
            n = len(r["dates"])
            hour_map[r["_id"]]   = round(r["total"]       / n) if n else 0
            group_map[r["_id"]]  = clamp_group_count(round(r["group_total"] / n) if n else 0, hour_map[r["_id"]])

    labels       = [f"{h:02d}:00" for h in range(hour_from, hour_to)]
    values       = [hour_map[f"{h:02d}"]  for h in range(hour_from, hour_to)]
    group_counts = [group_map[f"{h:02d}"] for h in range(hour_from, hour_to)]
    return jsonify({"labels": labels, "values": values, "group_counts": group_counts,
                    "date_from": date_from, "date_to": date_to, "is_avg": True})

# ─────────────────────────────────────────────
#  AGE GROUP
# ─────────────────────────────────────────────
@api_bp.route("/age-group")
@login_required_api
def age_group():
    """Aggregate age-group counts from beinghumanServer.age_group."""
    db = get_bh_db()
    date_from, date_to, date_to_ex = get_date_range()
    hour_from, hour_to = get_hour_range()
    store_code = get_store_code()

    filt = str_date_filter(date_from, date_to_ex)
    filt["store_code"] = store_code
    filt.update(hour_expr_str(hour_from, hour_to))

    # Display-only: 18-25 bucket excluded from the dashboard, DB values are untouched.
    VALID_GROUPS = {"Under 18", "25-35", "35-45", "45+"}
    filt["age_group"] = {"$in": list(VALID_GROUPS)}

    rows = list(db.age_group.aggregate([
        {"$match": filt},
        {"$group": {"_id": "$age_group", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}}
    ]))
    LABEL_MAP = {"Under 18": "Under 18", "18-25": "18-25", "25-35": "25-35", "35-45": "35-45", "45+": "45+"}
    buckets = [
        {"key": r["_id"], "label": LABEL_MAP[r["_id"]], "count": r["count"]}
        for r in rows if r["_id"] in VALID_GROUPS
    ]
    return jsonify({"buckets": buckets})

# ─────────────────────────────────────────────
#  EXPORT FOOTFALL EXCEL
# ─────────────────────────────────────────────
@api_bp.route("/export-footfall")
@login_required_api
def export_footfall():
    col = get_bh_db().footfall
    date_from, date_to, date_to_ex = get_date_range()
    hour_from, hour_to = get_hour_range()
    store_code = get_store_code()

    pipeline = [
        {"$match": {**str_date_filter(date_from, date_to_ex), **hour_expr_str(hour_from, hour_to), "store_code": store_code}},
        {"$group": {
            "_id": {
                "date": {"$substr": ["$date_time", 0, 10]},
                "hour": {"$substr": ["$date_time", 11, 2]}
            },
            "male": {"$sum": {"$convert": {"input": "$count_male", "to": "int", "onError": 0, "onNull": 0}}},
            "female": {"$sum": {"$convert": {"input": "$count_female", "to": "int", "onError": 0, "onNull": 0}}},
            "child": {"$sum": {"$convert": {"input": "$count_child", "to": "int", "onError": 0, "onNull": 0}}},
            "staff": {"$sum": {"$convert": {"input": "$count_staff", "to": "int", "onError": 0, "onNull": 0}}},
        }},
        {"$sort": {"_id.date": 1, "_id.hour": 1}}
    ]

    docs = list(col.aggregate(pipeline))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Footfall Hourly"

    header_fill = PatternFill("solid", fgColor="DA291C")
    header_font = Font(bold=True, color="FFFFFF", size=11)

    headers = ["Date", "Hour", "Male", "Female", "Child", "Staff", "Total Visitors"]
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for ri, d in enumerate(docs, 2):
        m = d.get("male", 0)
        f_ = d.get("female", 0)
        c = d.get("child", 0)
        s = d.get("staff", 0)
        date_val = d["_id"]["date"]
        hour_val = d["_id"]["hour"]

        ws.cell(row=ri, column=1, value=date_val)
        ws.cell(row=ri, column=2, value=f"{hour_val}:00")
        ws.cell(row=ri, column=3, value=m)
        ws.cell(row=ri, column=4, value=f_)
        ws.cell(row=ri, column=5, value=c)
        ws.cell(row=ri, column=6, value=s)
        ws.cell(row=ri, column=7, value=m + f_ + c)

    for i, col_dim in enumerate(ws.columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = 18

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return Response(
        buf.read(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=footfall_{store_code}_{date_from}_to_{date_to}.xlsx"}
    )

# ─────────────────────────────────────────────
#  CUSTOMER UNATTENDED
# ─────────────────────────────────────────────
@api_bp.route("/customer-unattended")
@login_required_api
def customer_unattended():
    """List customer-unattended alerts from beinghumanServer.alerts.

    The alert-type field name isn't fixed across pipelines, so we match
    any of the common candidates that contain "unattended".
    """
    col = get_bh_db().alerts
    date_from, date_to, date_to_ex = get_date_range()
    hour_from, hour_to = get_hour_range()
    store_code = get_store_code()

    filt = str_date_filter(date_from, date_to_ex)
    filt["store_code"] = store_code
    filt.update(hour_expr_str(hour_from, hour_to))
    filt["$or"] = [
        {"alert_type":  {"$regex": "unattended", "$options": "i"}},
        {"type":        {"$regex": "unattended", "$options": "i"}},
        {"event_type":  {"$regex": "unattended", "$options": "i"}},
        {"category":    {"$regex": "unattended", "$options": "i"}},
    ]

    rows = list(col.find(filt).sort("date_time", -1).limit(200))
    alerts = []
    for r in rows:
        alerts.append({
            "date_time": r.get("date_time", ""),
            "store_code": r.get("store_code", ""),
            "camera_no": r.get("camera_no", ""),
            "alert_type": r.get("alert_type") or r.get("type") or r.get("event_type") or r.get("category") or "Customer Unattended",
            "image_url": r.get("image_url", ""),
        })

    return jsonify({
        "alerts": alerts,
        "total": len(alerts),
        "date_from": date_from,
        "date_to": date_to,
    })
