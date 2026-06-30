#!/usr/bin/env python3
"""
Fetch BigQuery iOS user counts + Sentry hang/crash data for the current month,
then generate both outputs in a single run:

  1. sentry_data.xlsx              — Dashboard + Consolidation Excel workbook
  2. consolidation_report/index.html — HTML AQS consolidation report (all brands)

Usage:
    export SENTRY_AUTH_TOKEN=sntryu_...
    python3 fetch_sentry_data.py
"""

import urllib.request
import urllib.parse
import json
import os
import datetime
import sys
import warnings
import calendar

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.comments import Comment
except ImportError:
    raise SystemExit("ERROR: run  pip install openpyxl  first.")

try:
    warnings.filterwarnings("ignore")
    from google.cloud import bigquery as bq_client
    BQ_AVAILABLE = True
except ImportError:
    BQ_AVAILABLE = False

try:
    import gspread
    from google.auth import default as _gauth_default
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

SPREADSHEET_ID = "1IgBzRjH8XdStzD3L5-IyIg2A2_T4IjheChuhIm6drrc"
WORKSHEET_GID  = 138720290   # gid= from the URL fragment

TOKEN   = os.environ.get("SENTRY_AUTH_TOKEN", "")
ORG     = "delivery-hero-pm"
PROJECT = "4506937839976448"

if not TOKEN:
    raise SystemExit("ERROR: set SENTRY_AUTH_TOKEN environment variable first.")

TODAY      = datetime.date.today()
END_DATE   = TODAY - datetime.timedelta(days=1)          # yesterday
START_DATE = TODAY.replace(day=1)                        # 1st of current month

START    = START_DATE.strftime("%Y-%m-%dT00:00:00.000")
END      = END_DATE.strftime("%Y-%m-%dT23:59:59.999")
BQ_START = START_DATE.strftime("%Y-%m-%d")
BQ_END   = END_DATE.strftime("%Y-%m-%d")

HANGS_QUERY   = '!user.id:*-*-*-*-* app.in_foreground:True "*App hang* detected*"'
CRASHES_QUERY = "level:fatal handled:no"

HEADERS = {"Authorization": f"Bearer {TOKEN}"}

# Firebase Performance — brand display name → BigQuery table suffix
FIREBASE_BRANDS = [
    ("Foodpanda",     "com_logistics_rider_foodpanda_IOS"),
    ("Foodora",       "com_logistics_rider_foodora_IOS"),
    ("Talabat",       "com_logistics_rider_talabat_IOS"),
    ("pedidosya",     "com_logistics_rider_pedidosya_IOS"),
    ("HungerStation", "com_logistics_rider_hungerstation_IOS"),
    ("Yemeksepeti",   "com_logistics_rider_yemeksepeti_IOS"),
    ("Glovo",         "com_logistics_rider_glovo_IOS"),
    ("Woowa",         "com_logistics_rider_woowabros_IOS"),
    ("efood",         "com_logistics_rider_efood_IOS"),
    ("Foody",         "com_logistics_rider_foody_IOS"),
]

# Derived from FIREBASE_BRANDS for use in Firebase fetch queries
FIREBASE_TABLES = [t for _, t in FIREBASE_BRANDS]

# ── Brand definitions (shared by Excel and HTML outputs) ──────────────────────

BRANDS = [
    {"name": "Foodpanda",     "bq_key": "foodpanda",     "sentry_key": "foodpanda",     "app_size": 76.1, "asti": 4.43, "stti": 0.95, "riders": 78203},
    {"name": "Foodora",       "bq_key": "foodora",       "sentry_key": "foodora",       "app_size": 80.3, "asti": 4.43, "stti": 0.95, "riders": 14946},
    {"name": "Talabat",       "bq_key": "talabat",       "sentry_key": "talabat",       "app_size": 76.0, "asti": 4.43, "stti": 0.95, "riders": 16988},
    {"name": "pedidosya",     "bq_key": "pedidosya",     "sentry_key": "pedidosya",     "app_size": 66.6, "asti": 4.43, "stti": 0.95, "riders": 20641},
    {"name": "HungerStation", "bq_key": "hungerstation", "sentry_key": "hungerstation", "app_size": 76.0, "asti": 4.43, "stti": 0.95, "riders": 10630},
    {"name": "Yemeksepeti",   "bq_key": "yemeksepeti",   "sentry_key": "yemeksepeti",   "app_size": 76.4, "asti": 4.43, "stti": 0.95, "riders": 2186},
    {"name": "Glovo",         "bq_key": "glovo",         "sentry_key": "glovo",         "app_size": 76.2, "asti": 4.43, "stti": 0.95, "riders": 46896},
    {"name": "Woowa",         "bq_key": "woowabros",     "sentry_key": "woowa",         "app_size": 75.9, "asti": 4.43, "stti": 0.95, "riders": 731},
    {"name": "efood",         "bq_key": "efood",         "sentry_key": "efood",         "app_size": 66.4, "asti": 4.43, "stti": 0.95, "riders": 5258},
    {"name": "Foody",         "bq_key": "foody",         "sentry_key": "foody",         "app_size": 66.3, "asti": 4.43, "stti": 0.95, "riders": 729},
]

WEIGHTS      = [0.3966, 0.0758, 0.0861, 0.1047, 0.0539, 0.0111, 0.2378, 0.0037, 0.0267, 0.0037]
RIDER_COUNTS = [78203, 14946, 16988, 20641, 10630, 2186, 46896, 731, 5258, 729]

# ── AQS config ─────────────────────────────────────────────────────────────────

AQS_CONFIG = {
    "cfu":      {"baseline": 99.7, "target": 99.9, "weight": 35},
    "hang":     {"baseline": 99.7, "target": 99.9, "weight": 25},
    "app_size": {"baseline": 75,   "target": 60,   "weight": 2},
    "asti":     {"baseline": 4,    "target": 2,    "weight": 10},
    "stti":     {"baseline": 1.5,  "target": 0.5,  "weight": 10},
    "frozen":   {"baseline": 3,    "target": 1,    "weight": 13},
    "skipped":  {"baseline": 2,    "target": 1,    "weight": 5},
}

# ── Excel style constants ──────────────────────────────────────────────────────

BQ_HEADER_FILL    = PatternFill("solid", fgColor="BDD7EE")  # pastel sky blue
CRASH_HEADER_FILL = PatternFill("solid", fgColor="FFADAD")  # pastel rose
HANG_HEADER_FILL  = PatternFill("solid", fgColor="C9B8F0")  # pastel lavender
SETTINGS_FILL     = PatternFill("solid", fgColor="C6EFCE")  # pastel mint green
DASH_HEADER_FILL  = PatternFill("solid", fgColor="B8D4E8")  # pastel steel blue
EVEN_ROW_FILL     = PatternFill("solid", fgColor="EEF4FB")  # very light pastel blue
SUMMARY_LBL_FILL  = PatternFill("solid", fgColor="FDEBD0")  # pastel peach
SUMMARY_VAL_FILL  = PatternFill("solid", fgColor="D5F5E3")  # pastel sage green
BOLD = Font(bold=True)

_THIN   = Side(style="thin",   color="BFBFBF")
_MEDIUM = Side(style="medium", color="808080")
THIN_BOX   = Border(left=_THIN,   right=_THIN,   top=_THIN,   bottom=_THIN)
MEDIUM_BOX = Border(left=_MEDIUM, right=_MEDIUM, top=_MEDIUM, bottom=_MEDIUM)

# Raw data column positions — must match the SUMIFS formulas in the dashboard:
#   D =SUMIFS(V:V,  U:U, "*"&$B$10&"*", T:T,  C{row})  → BQ at T(20)-V(22)
#   E =SUMIFS(X:X,  Z:Z, "*"&$B$10&"*", Y:Y,  C{row})  → Crash at X(24)-Z(26)
#   K =SUMIFS(AB:AB,AD:AD,"*"&$B$10&"*",AC:AC,C{row})  → Hang at AB(28)-AD(30)
BQ_START_COL    = 20   # T
CRASH_START_COL = 24   # X
HANG_START_COL  = 28   # AB


# ── Data fetching ──────────────────────────────────────────────────────────────

def fetch_bigquery():
    if not BQ_AVAILABLE:
        print("WARNING: google-cloud-bigquery not installed, skipping BQ fetch.", file=sys.stderr)
        return []
    query = f"""
        SELECT partition_date as dt, appId, count(distinct clientId) as user_count
        FROM `fulfillment-dwh-production.curated_data_shared_coredata_tracking.perseus_events_rider_app`
        WHERE partition_date BETWEEN '{BQ_START}' AND '{BQ_END}'
          AND platform = 'iOS'
        GROUP BY ALL
        ORDER BY appId, dt ASC
    """
    try:
        client = bq_client.Client(project="logistics-rider-staging")
        rows = list(client.query(query))
        return [
            {
                "dt":         str(row.dt),
                "appId":      row.appId or "",
                "user_count": row.user_count,
            }
            for row in rows
        ]
    except Exception as exc:
        print(f"WARNING: BigQuery fetch failed, skipping BQ data: {exc}", file=sys.stderr)
        return []


def fetch_firebase_frames():
    """Return {"frozen": float, "skipped": float} aggregated across ALL brands.

    Runs a single query that UNIONs all brand tables — same structure as the
    reference SQL — and returns one row.  The same values are applied to every
    brand row in the Consolidation sheet.

    Date window: 1st of current month (inclusive) → yesterday (inclusive).
    """
    if not BQ_AVAILABLE:
        print("WARNING: google-cloud-bigquery not available, skipping Firebase frames fetch.",
              file=sys.stderr)
        return {}

    selects = "\n  UNION ALL".join(
        f"""
  SELECT
    app_display_version, app_build_version, os_version, device_name, country,
    event_name, event_type, event_timestamp,
    trace_info.duration_us,
    trace_info.screen_info,
    trace_info.metric_info
  FROM `logistics-54934.firebase_performance.{table}`
  WHERE TIMESTAMP_TRUNC(_PARTITIONTIME, DAY) >= TIMESTAMP('{BQ_START}')
    AND TIMESTAMP_TRUNC(_PARTITIONTIME, DAY) <= TIMESTAMP('{BQ_END}')"""
        for _, table in FIREBASE_BRANDS
    )

    query = f"""
WITH fb_performance AS ({selects}
),
frames_grader AS (
  SELECT
    event_name,
    event_type,
    duration_us,
    screen_info.slow_frame_ratio,
    screen_info.frozen_frame_ratio,
    CONCAT(event_name, CAST(event_timestamp AS STRING), app_build_version, country) AS event_id,
    CASE
      WHEN COUNT(CONCAT(event_name, CAST(event_timestamp AS STRING), app_build_version, country)) > 0
        THEN 1 ELSE 0
    END AS total_event_count,
    CASE WHEN COUNTIF(screen_info.slow_frame_ratio   > 0.5)   > 0 THEN 1 ELSE 0 END AS count_skipped_greater50,
    CASE WHEN COUNTIF(screen_info.frozen_frame_ratio > 0.001) > 0 THEN 1 ELSE 0 END AS count_frozen_greater1
  FROM fb_performance
  WHERE screen_info IS NOT NULL
    AND CHAR_LENGTH(event_name) > 9
  GROUP BY 1, 2, 3, 4, 5, 6
),
events_totals AS (
  SELECT
    SUM(total_event_count)       AS events_count,
    SUM(count_skipped_greater50) AS skipped_frames_count,
    SUM(count_frozen_greater1)   AS frozen_frames_count
  FROM frames_grader
),
frames_percentages_calc AS (
  SELECT
    (skipped_frames_count / events_count) * 100 AS skipped_frames_percentage,
    (frozen_frames_count  / events_count) * 100 AS forzen_frames_percentage
  FROM events_totals
)
SELECT
  ROUND(skipped_frames_percentage, 2) AS skipped_frames_percentage,
  ROUND(forzen_frames_percentage,  2) AS forzen_frames_percentage,
  ROUND(100 - ((0.2 * skipped_frames_percentage) + (0.8 * forzen_frames_percentage)), 2) AS performance_score
FROM frames_percentages_calc
"""
    try:
        client = bq_client.Client(project="logistics-rider-staging")
        rows = list(client.query(query))
        if not rows:
            print("WARNING: Firebase frames query returned no rows.", file=sys.stderr)
            return {}
        row = rows[0]
        result = {
            "frozen":  float(row.forzen_frames_percentage),
            "skipped": float(row.skipped_frames_percentage),
        }
        print(f"Firebase frames: frozen={result['frozen']}%  skipped={result['skipped']}%")
        return result
    except Exception as exc:
        print(f"ERROR fetching Firebase frames: {exc}", file=sys.stderr)
        return {}


def fetch_discover(query, environment=None, start=None, end=None):
    rows = []
    cursor = None
    _start = start or START
    _end   = end   or END

    while True:
        params = [
            ("project",  PROJECT),
            ("query",    query),
            ("start",    _start),
            ("end",      _end),
            ("field",    "timestamp.to_day"),
            ("field",    "Brand"),
            ("field",    "count_unique(user)"),
            ("sort",     "timestamp.to_day"),
            ("dataset",  "errors"),
            ("per_page", "50"),
        ]
        if environment:
            params.append(("environment", environment))
        if cursor:
            params.append(("cursor", cursor))

        url = (f"https://sentry.io/api/0/organizations/{ORG}/events/?"
               + urllib.parse.urlencode(params))
        req = urllib.request.Request(url, headers=HEADERS)
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = json.loads(resp.read())
                link = resp.headers.get("Link", "")
        except Exception as e:
            print(f"ERROR fetching Discover: {e}", file=sys.stderr)
            break

        rows.extend(data.get("data", []))

        next_cursor = None
        for part in link.split(","):
            if 'rel="next"' in part and 'results="true"' in part:
                for seg in part.split(";"):
                    seg = seg.strip()
                    if seg.startswith("<"):
                        next_cursor = urllib.parse.parse_qs(
                            urllib.parse.urlparse(seg[1:-1]).query
                        ).get("cursor", [None])[0]
        if not next_cursor:
            break
        cursor = next_cursor

    return rows


def fetch_discover_max(query, environment=None, start=None, end=None):
    """Run the Discover query twice and return rows with the max count_unique(user)
    per (Brand, day) pair, so intermittent Sentry under-counts are corrected."""
    run1 = fetch_discover(query, environment=environment, start=start, end=end)
    run2 = fetch_discover(query, environment=environment, start=start, end=end)

    best = {}  # (brand, day) → (max_count, row)
    for rows in (run1, run2):
        for row in rows:
            brand = (row.get("Brand") or "")
            day   = (row.get("timestamp.to_day") or "")[:10]
            count = int(row.get("count_unique(user)", 0) or 0)
            key   = (brand.lower(), day)
            if key not in best or count > best[key][0]:
                best[key] = (count, row)

    result = []
    for (_, _), (count, row) in best.items():
        merged = dict(row)
        merged["count_unique(user)"] = count
        result.append(merged)
    return result


def fetch_discover_per_brand(query, environment=None):
    """Run a separate Discover query per brand with an explicit Brand filter.
    Merges all results into a single list. More reliable than a single query
    since Sentry can silently drop or zero-out brand rows in grouped results."""
    all_rows = []
    for brand in BRANDS:
        skey = brand["sentry_key"]
        brand_query = f'{query} Brand:*{skey}*'
        rows = fetch_discover(brand_query, environment=environment)
        total = sum(int(r.get("count_unique(user)", 0) or 0) for r in rows)
        print(f"    {skey}: {total} users across {len(rows)} day-rows")
        all_rows.extend(rows)
    return all_rows


def shape_rows(raw, user_col):
    rows = [
        {
            user_col:           int(row.get("count_unique(user)", 0)),
            "day":              row.get("timestamp.to_day", "")[:10],
            "timestamp.to_day": row.get("timestamp.to_day", ""),
            "Brand":            row.get("Brand", ""),
        }
        for row in raw
    ]
    return sorted(rows, key=lambda r: r["timestamp.to_day"])


def aggregate_by_brand(rows, user_col="count_unique(user)"):
    """Aggregate user counts from Sentry rows by brand (substring match on sentry_key).

    Accepts either raw Discover rows (user_col="count_unique(user)") or
    shaped rows (user_col="CRASH_USERS" / "HANG_USERS").
    Returns dict: {sentry_key_lower: total_count} for each brand in BRANDS.
    """
    totals = {}
    for brand in BRANDS:
        skey = brand["sentry_key"].lower()
        total = 0
        for row in rows:
            brand_field = (row.get("Brand") or "").lower()
            if skey in brand_field:
                total += int(row.get(user_col, 0) or 0)
        totals[skey] = total
    return totals


def backfill_zero_rows(rows, user_col, query, environment=None):
    """For any row where user_col is 0, re-query Sentry for that day to get the real count."""
    zero_days = {row["day"] for row in rows if row[user_col] == 0}
    for day in sorted(zero_days):
        print(f"  Re-querying {user_col} for {day} (row had 0 count)...")
        day_raw = fetch_discover(
            query, environment=environment,
            start=f"{day}T00:00:00.000",
            end=f"{day}T23:59:59.999",
        )
        day_counts = {}
        for r in day_raw:
            brand = (r.get("Brand") or "").lower()
            day_counts[brand] = day_counts.get(brand, 0) + int(r.get("count_unique(user)", 0) or 0)
        for row in rows:
            if row["day"] == day and row[user_col] == 0:
                brand = (row.get("Brand") or "").lower()
                row[user_col] = day_counts.get(brand, 0)


# ── AQS formula ────────────────────────────────────────────────────────────────

def aqs_score(value, col_key):
    """Compute AQS score for a single metric value."""
    cfg = AQS_CONFIG[col_key]
    baseline = cfg["baseline"]
    target   = cfg["target"]
    weight   = cfg["weight"]
    score = min(100, max(0, (((value - baseline) / (target - baseline)) * 50) + 50))
    return score * weight / 100


# ── Color coding (for HTML report) ────────────────────────────────────────────

def color_cfu(v):
    if v >= 99.9:
        return "#E8F5EA"
    elif v >= 99.7:
        return "#FFF8E1"
    return "#FFE4E6"


def color_hang(v):
    return color_cfu(v)  # same thresholds


def color_app_size(v):
    if v <= 60:
        return "#E8F5EA"
    elif v <= 75:
        return "#FFF8E1"
    return "#FFE4E6"


def color_asti(v):
    if v <= 2:
        return "#E8F5EA"
    elif v <= 4:
        return "#FFF8E1"
    return "#FFE4E6"


def color_stti(v):
    if v <= 0.5:
        return "#E8F5EA"
    elif v <= 1.5:
        return "#FFF8E1"
    return "#FFE4E6"


def color_frozen(v):
    if v <= 1:
        return "#E8F5EA"
    elif v <= 3:
        return "#FFF8E1"
    return "#FFE4E6"


def color_skipped(v):
    if v <= 1:
        return "#E8F5EA"
    elif v <= 2:
        return "#FFF8E1"
    return "#FFE4E6"


COLOR_FNS = {
    "cfu":      color_cfu,
    "hang":     color_hang,
    "app_size": color_app_size,
    "asti":     color_asti,
    "stti":     color_stti,
    "frozen":   color_frozen,
    "skipped":  color_skipped,
}

# Pastel brand colors — soft, muted tones
BRAND_COLORS = [
    "#FFE8EA",  # Foodpanda    — Soft Rose
    "#E6FAF0",  # Foodora      — Soft Mint
    "#E4F1FF",  # Talabat      — Soft Sky Blue
    "#FEFEE8",  # pedidosya    — Soft Lemon
    "#F2E8F6",  # HungerStation — Soft Lavender
    "#E4F9F8",  # Yemeksepeti  — Soft Turquoise
    "#FEFEF2",  # Glovo        — Soft Cream
    "#DFF8F7",  # Woowa        — Soft Aqua
    "#EEE8E2",  # efood        — Soft Taupe
    "#EAE4FA",  # Foody        — Soft Purple
]


# ── HTML report generation ─────────────────────────────────────────────────────

def generate_html_report(bq_users_by_brand, crash_by_brand, hang_by_brand, firebase_data):
    """Generate consolidation_report/index.html from pre-fetched data.

    Args:
        bq_users_by_brand: dict {appid_lower: total_user_count}
        crash_by_brand:    dict {sentry_key_lower: total_crash_users}
        hang_by_brand:     dict {sentry_key_lower: total_hang_users}
        firebase_data:     dict {"frozen": float, "skipped": float} (may be empty)
    """
    OUT_DIR = "consolidation_report"

    fetch_date = TODAY.strftime("%-d %b %Y")
    date_range = f"{START_DATE.strftime('%-d %b')} – {END_DATE.strftime('%-d %b %Y')}"

    # ── Per-brand metrics ──────────────────────────────────────────────────────
    # Crash/hang counts from Sentry are summed daily unique users (user-days).
    # BQ user counts are also summed daily (user-days) — same unit, ratio is valid.
    # Fallback uses riders × days to approximate user-days when BQ is unavailable.
    days_in_period = (END_DATE - START_DATE).days + 1

    brand_metrics = []
    for brand in BRANDS:
        skey = brand["sentry_key"].lower()
        bkey = brand["bq_key"].lower()

        # Sum BQ users for this brand (substring match on appId)
        users = 0
        for appid, count in bq_users_by_brand.items():
            if bkey in appid:
                users += count

        crashes = crash_by_brand.get(skey, 0)
        hangs   = hang_by_brand.get(skey, 0)

        if users > 0:
            cfu  = max(0.0, min(100.0, int((1 - crashes / users) * 10000) / 100))
            hang = max(0.0, min(100.0, int((1 - hangs   / users) * 10000) / 100))
        else:
            # Scale riders by days to match the user-days unit of crash/hang counts
            fallback_user_days = brand["riders"] * days_in_period
            cfu  = max(0.0, min(100.0, int((1 - crashes / fallback_user_days) * 10000) / 100))
            hang = max(0.0, min(100.0, int((1 - hangs   / fallback_user_days) * 10000) / 100))

        frozen  = firebase_data.get("frozen",  0.58)
        skipped = firebase_data.get("skipped", 1.2)

        brand_metrics.append({
            "name":     brand["name"],
            "riders":   brand["riders"],
            "users":    users,
            "crashes":  crashes,
            "hangs":    hangs,
            "cfu":      cfu,
            "hang":     hang,
            "app_size": brand["app_size"],
            "asti":     brand["asti"],
            "stti":     brand["stti"],
            "frozen":   frozen,
            "skipped":  skipped,
            "fallback": users == 0,
        })

    print("\nBrand metrics:")
    for m in brand_metrics:
        flag = " (fallback)" if m["fallback"] else ""
        print(f"  {m['name']:15s}  users={m['users']:>8,}  crashes={m['crashes']:>5}  hangs={m['hangs']:>5}"
              f"  cfu={m['cfu']:.2f}%  hang={m['hang']:.2f}%{flag}")

    # ── Weighted averages ──────────────────────────────────────────────────────
    def weighted_avg(key):
        values = [m[key] for m in brand_metrics]
        return int(sum(v * w for v, w in zip(values, WEIGHTS)) * 100) / 100

    avg_metrics = {
        "cfu":      weighted_avg("cfu"),
        "hang":     weighted_avg("hang"),
        "app_size": weighted_avg("app_size"),
        "asti":     weighted_avg("asti"),
        "stti":     weighted_avg("stti"),
        "frozen":   weighted_avg("frozen"),
        "skipped":  weighted_avg("skipped"),
    }

    # ── AQS scores ────────────────────────────────────────────────────────────
    aqs_scores = {k: round(aqs_score(avg_metrics[k], k), 4) for k in AQS_CONFIG}
    final_aqs  = round(sum(aqs_scores.values()), 2)

    print(f"\nWeighted averages: {avg_metrics}")
    print(f"AQS scores:        {aqs_scores}")
    print(f"FINAL AQS:         {final_aqs}")

    # ── HTML helpers ──────────────────────────────────────────────────────────
    def td(value, col_key, fmt=None, extra_style=""):
        """Return a <td> with background color coding and formatted value."""
        color = COLOR_FNS[col_key](value)
        text  = fmt(value) if fmt else str(value)
        style = f"background:{color};{extra_style}"
        return f'<td style="{style}">{text}</td>'

    def fmt_pct(v):
        return f"{v:.2f}%"

    def fmt_mb(v):
        return f"{v}"

    def fmt_s(v):
        return f"{v}"

    # ── Brand data rows — each brand gets its own pastel row color ────────────
    brand_rows_html = ""
    for i, m in enumerate(brand_metrics):
        row_color = BRAND_COLORS[i % len(BRAND_COLORS)]
        fallback_note = (' <span style="font-size:10px;color:#6b7280;" title="BQ unavailable — using fallback">*</span>'
                         if m["fallback"] else "")
        cell_style = f'background:{row_color};'
        brand_rows_html += f"""
      <tr style="background:{row_color};">
        <td style="{cell_style}font-weight:600;white-space:nowrap;">{m['name']}{fallback_note}</td>
        <td style="{cell_style}text-align:right;">{fmt_pct(m['cfu'])}</td>
        <td style="{cell_style}text-align:right;">{fmt_pct(m['hang'])}</td>
        <td style="{cell_style}text-align:right;">{fmt_mb(m['app_size'])}</td>
        <td style="{cell_style}text-align:right;">{fmt_s(m['asti'])}</td>
        <td style="{cell_style}text-align:right;">{fmt_s(m['stti'])}</td>
        <td style="{cell_style}text-align:right;">{fmt_pct(m['frozen'])}</td>
        <td style="{cell_style}text-align:right;">{fmt_pct(m['skipped'])}</td>
        <td style="{cell_style}text-align:right;">{m['riders']:,}</td>
      </tr>"""

    # ── AQS score row cells ────────────────────────────────────────────────────
    aqs_col_order = ["cfu", "hang", "app_size", "asti", "stti", "frozen", "skipped"]
    aqs_cells_html = "".join(
        f'<td style="text-align:right;">{aqs_scores[k]:.4f}</td>'
        for k in aqs_col_order
    )

    bq_status_note = ""
    if not BQ_AVAILABLE:
        bq_status_note = """
    <div style="margin-top:12px;padding:10px 16px;background:#FFF3CD;border:1px solid #ffc107;
                border-radius:6px;font-size:13px;color:#856404;">
      ⚠ Using fallback data (BigQuery unavailable). Install <code>google-cloud-bigquery</code>
      and authenticate to fetch live data.
    </div>"""

    brand_json = json.dumps([
        {
            "name":     m["name"],
            "color":    BRAND_COLORS[i],
            "cfu":      m["cfu"],
            "hang":     m["hang"],
            "app_size": m["app_size"],
            "asti":     m["asti"],
            "stti":     m["stti"],
            "frozen":   m["frozen"],
            "skipped":  m["skipped"],
            "riders":   m["riders"],
        }
        for i, m in enumerate(brand_metrics)
    ], indent=2)

    bq_available_str = 'True' if BQ_AVAILABLE else 'False'

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>iOS AQS Consolidation Report</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
  <style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            background: #f8fafc; color: #111; padding: 32px; }}
    h1 {{ font-size: 22px; font-weight: 700; margin-bottom: 4px; }}
    .subtitle {{ color: #666; font-size: 13px; margin-bottom: 24px; }}
    .warning-banner {{
      background: #fffbf0; border: 1px solid #f59e0b;
      border-left: 4px solid #f59e0b; border-radius: 6px;
      padding: 12px 16px; font-size: 13px; color: #92400e;
      font-weight: 600; margin-bottom: 16px;
    }}

    /* ── Tabs ── */
    .tabs {{ display: flex; gap: 4px; margin-bottom: 0; border-bottom: 2px solid #e5e7eb; }}
    .tab {{ padding: 9px 20px; font-size: 13px; font-weight: 600; cursor: pointer;
             border-radius: 6px 6px 0 0; border: 1px solid transparent;
             border-bottom: none; color: #6b7280; background: transparent;
             position: relative; top: 2px; transition: color .15s; }}
    .tab:hover {{ color: #111; }}
    .tab.active {{ color: #111; background: #fff; border-color: #e5e7eb;
                   border-bottom-color: #fff; }}
    .panel {{ display: none; background: #fff; border: 1px solid #e5e7eb;
               border-top: none; border-radius: 0 6px 6px 6px; padding: 28px; }}
    .panel.active {{ display: block; }}

    /* ── AQS Table panel ── */
    .score-wrap {{ display: flex; gap: 16px; margin-bottom: 28px; flex-wrap: wrap; }}
    .score-card {{ background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 8px;
                   padding: 14px 24px; min-width: 180px; }}
    .score-card .sc-label {{ font-size: 11px; text-transform: uppercase;
                              letter-spacing: .05em; color: #6b7280; margin-bottom: 4px; }}
    .score-card .sc-value {{ font-size: 32px; font-weight: 700; color: #111; }}
    .score-card .sc-sub   {{ font-size: 11px; color: #9ca3af; margin-top: 2px; }}
    .table-wrap {{ overflow-x: auto; margin-bottom: 28px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
    thead th {{
      text-align: left; padding: 8px 12px;
      border-bottom: 2px solid #e5e7eb; color: #666; font-weight: 600;
      font-size: 11px; text-transform: uppercase; letter-spacing: .05em; white-space: nowrap;
    }}
    thead th.num {{ text-align: right; }}
    tbody tr {{ border-bottom: 1px solid #f3f4f6; }}
    tbody tr:hover {{ background: #f9fafb; }}
    tbody td {{ padding: 9px 12px; vertical-align: middle; white-space: nowrap; text-align: right; }}
    tbody td:first-child {{ text-align: left; font-weight: 500; }}
    tr.row-avg td {{
      background: #f5fbff !important; font-weight: 700;
      border-top: 2px solid #dbeafe; color: #3b82f6;
    }}
    tr.row-avg td:first-child {{ text-align: left; }}
    tr.row-aqs td {{
      background: #f6fef8 !important; font-weight: 700;
      border-top: 2px solid #d1fae5; color: #34d399;
    }}
    tr.row-aqs td:first-child {{ text-align: left; }}

    /* ── Dashboard panel ── */
    .pill-actions {{ display: flex; gap: 8px; margin-bottom: 12px; align-items: center; }}
    .pill-action-btn {{
      font-size: 12px; padding: 4px 14px; border-radius: 6px; cursor: pointer;
      border: 1px solid #d1d5db; background: #f9fafb; color: #374151; font-weight: 500;
    }}
    .pill-action-btn:hover {{ background: #e5e7eb; }}
    .brand-pills {{ display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 24px; }}
    .brand-pill {{
      padding: 5px 14px; border-radius: 20px; font-size: 12px; font-weight: 600;
      cursor: pointer; border: 2px solid rgba(0,0,0,.12); transition: opacity .15s;
      user-select: none;
    }}
    .brand-pill.inactive {{ opacity: .3; }}
    .dash-summary {{ display: flex; gap: 14px; flex-wrap: wrap; margin-bottom: 28px; }}
    .dash-card {{
      background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 8px;
      padding: 14px 20px; min-width: 150px; flex: 1;
    }}
    .dash-card .dc-label {{ font-size: 11px; text-transform: uppercase;
                             letter-spacing: .05em; color: #6b7280; margin-bottom: 4px; }}
    .dash-card .dc-value {{ font-size: 24px; font-weight: 700; color: #111; }}
    .dash-card .dc-sub {{ font-size: 11px; color: #9ca3af; margin-top: 2px; }}
    .charts-row {{ display: flex; gap: 20px; flex-wrap: wrap; margin-bottom: 24px; }}
    .chart-box {{
      flex: 1; min-width: 280px; background: #fff; border: 1px solid #e5e7eb;
      border-radius: 8px; padding: 18px;
    }}
    .chart-box h3 {{ font-size: 12px; font-weight: 600; color: #374151;
                     text-transform: uppercase; letter-spacing: .04em; margin-bottom: 14px; }}
    .chart-box canvas {{ max-width: 100%; height: 240px !important; }}
    .dash-section-title {{
      font-size: 12px; font-weight: 600; color: #374151;
      text-transform: uppercase; letter-spacing: .04em; margin-bottom: 12px;
    }}
    .dash-table-wrap {{ overflow-x: auto; }}
    .dash-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
    .dash-table thead th {{
      text-align: right; padding: 8px 12px; border-bottom: 2px solid #e5e7eb;
      color: #666; font-weight: 600; font-size: 11px;
      text-transform: uppercase; letter-spacing: .05em; white-space: nowrap;
    }}
    .dash-table thead th:first-child {{ text-align: left; }}
    .dash-table tbody tr {{ border-bottom: 1px solid #f3f4f6; }}
    .dash-table tbody tr:hover {{ filter: brightness(.96); }}
    .dash-table tbody td {{ padding: 9px 12px; text-align: right; white-space: nowrap; }}
    .dash-table tbody td:first-child {{ text-align: left; font-weight: 600; }}

    .footer {{ margin-top: 32px; font-size: 11px; color: #9ca3af;
               border-top: 1px solid #e5e7eb; padding-top: 16px; line-height: 1.8; }}
    .footer code {{ background: #f3f4f6; padding: 1px 5px; border-radius: 3px; font-size: 11px; }}
  </style>
</head>
<body>

<h1>iOS AQS Consolidation Report</h1>
<p class="subtitle">
  Period: {date_range} &nbsp;|&nbsp; Fetched: {fetch_date}
</p>

<div class="warning-banner">
  ⚠ App size, ASTI and STTI have the older values, not recently updated. Update these columns manually before sharing. For the accurate AQS score, update these values.
</div>
{bq_status_note}

<div class="tabs" id="tabs">
  <div class="tab active" data-panel="aqs-table">AQS Table</div>
  <div class="tab" data-panel="dashboard">Dashboard</div>
</div>

<!-- ── Panel 1: AQS Table ── -->
<div class="panel active" id="panel-aqs-table">

<div class="table-wrap">
  <table>
    <thead>
      <tr>
        <th>Brand</th>
        <th class="num">Crash Free %</th>
        <th class="num">App Hangs %</th>
        <th class="num">App Size (MB)</th>
        <th class="num">ASTI (s)</th>
        <th class="num">STTI (s)</th>
        <th class="num">Frozen Frames %</th>
        <th class="num">Skipped Frames %</th>
        <th class="num">Riders</th>
      </tr>
    </thead>
    <tbody>
      {brand_rows_html}
      <tr class="row-avg">
        <td>Weighted AVG</td>
        <td>{avg_metrics['cfu']:.2f}%</td>
        <td>{avg_metrics['hang']:.2f}%</td>
        <td>{avg_metrics['app_size']}</td>
        <td>{avg_metrics['asti']}</td>
        <td>{avg_metrics['stti']}</td>
        <td>{avg_metrics['frozen']:.2f}%</td>
        <td>{avg_metrics['skipped']:.2f}%</td>
        <td></td>
      </tr>
      <tr class="row-aqs">
        <td>AQS Score</td>
        {aqs_cells_html}
        <td></td>
      </tr>
    </tbody>
  </table>
</div>

<div class="score-wrap">
  <div class="score-card">
    <div class="sc-label">Final AQS Score</div>
    <div class="sc-value">{final_aqs}</div>
    <div class="sc-sub">{date_range}</div>
  </div>
</div>

<div class="footer">
  <strong>Notes:</strong><br>
  &bull; <strong>Crash Free % / App Hangs %</strong> — computed from Sentry Discover API
    (<code>level:fatal handled:no</code> / <code>!user.id:*-*-*-*-* app.in_foreground:True "*App hang* detected*"</code>)
    against BQ daily user counts for the period.<br>
  &bull; <strong>Frozen Frames % / Skipped Frames %</strong> — single aggregated value from Firebase Performance BQ,
    applied to all brands equally.<br>
  &bull; <strong>App Size, ASTI, STTI</strong> — older values, not recently updated; update manually for accurate AQS score.<br>
  &bull; <strong>Weighted AVG</strong> — SUMPRODUCT of brand values with rider-share weights.<br>
  &bull; <strong>AQS formula</strong> — <code>min(100, max(0, (((value - baseline) / (target - baseline)) * 50) + 50)) * weight%</code><br>
  &bull; BQ_AVAILABLE = <code>{bq_available_str}</code> &nbsp;|&nbsp;
    Fallback frozen=0.58, skipped=1.2 used when Firebase BQ is unavailable.<br>
  &bull; Rows marked with <span style="color:#9ca3af">*</span> used fallback CFU/hang values (BQ user count was zero).
</div>

</div><!-- /panel-aqs-table -->

<!-- ── Panel 2: Dashboard ── -->
<div class="panel" id="panel-dashboard">

  <div class="pill-actions">
    <button class="pill-action-btn" onclick="selectAllBrands()">Select All</button>
    <button class="pill-action-btn" onclick="clearAllBrands()">Clear All</button>
  </div>
  <div class="brand-pills" id="brand-pills"></div>

  <div class="dash-summary" id="dash-summary"></div>

  <div class="charts-row">
    <div class="chart-box">
      <h3>Crash Free % per Brand</h3>
      <canvas id="chart-cfu"></canvas>
    </div>
    <div class="chart-box">
      <h3>App Hangs % per Brand</h3>
      <canvas id="chart-hang"></canvas>
    </div>
  </div>

  <div class="charts-row">
    <div class="chart-box">
      <h3>Frozen &amp; Skipped Frames % per Brand</h3>
      <canvas id="chart-frames"></canvas>
    </div>
    <div class="chart-box">
      <h3>App Size (MB) per Brand</h3>
      <canvas id="chart-appsize"></canvas>
    </div>
  </div>

  <p class="dash-section-title">Brand Detail</p>
  <div class="dash-table-wrap">
    <table class="dash-table">
      <thead>
        <tr>
          <th>Brand</th>
          <th>Crash Free %</th>
          <th>App Hangs %</th>
          <th>App Size (MB)</th>
          <th>ASTI (s)</th>
          <th>STTI (s)</th>
          <th>Frozen Frames %</th>
          <th>Skipped Frames %</th>
          <th>Riders</th>
        </tr>
      </thead>
      <tbody id="dash-table-body"></tbody>
    </table>
  </div>

</div><!-- /panel-dashboard -->

<script>
var BRAND_DATA = {brand_json};

// ── Tab switching ─────────────────────────────────────────────────────────────
document.getElementById('tabs').addEventListener('click', function(e) {{
  var t = e.target.closest('.tab');
  if (!t) return;
  document.querySelectorAll('.tab').forEach(function(x) {{ x.classList.remove('active'); }});
  document.querySelectorAll('.panel').forEach(function(x) {{ x.classList.remove('active'); }});
  t.classList.add('active');
  document.getElementById('panel-' + t.dataset.panel).classList.add('active');
}});

// ── Brand selection state ─────────────────────────────────────────────────────
var selected = BRAND_DATA.map(function() {{ return true; }});
var chartCfu, chartHang, chartFrames, chartAppsize;

function getActiveBrands() {{
  return BRAND_DATA.filter(function(b, i) {{ return selected[i]; }});
}}

function riderWeightedAvg(key) {{
  var active = getActiveBrands();
  if (!active.length) return 0;
  var totalRiders = 0, sum = 0;
  active.forEach(function(b) {{ sum += b[key] * b.riders; totalRiders += b.riders; }});
  return totalRiders ? sum / totalRiders : 0;
}}

// ── Summary cards ─────────────────────────────────────────────────────────────
function renderSummary() {{
  var active = getActiveBrands();
  var cards = [
    {{ key: 'cfu',      label: 'Crash Free %',     fmt: function(v) {{ return v.toFixed(2) + '%'; }} }},
    {{ key: 'hang',     label: 'App Hangs %',       fmt: function(v) {{ return v.toFixed(2) + '%'; }} }},
    {{ key: 'frozen',   label: 'Frozen Frames %',   fmt: function(v) {{ return v.toFixed(2) + '%'; }} }},
    {{ key: 'skipped',  label: 'Skipped Frames %',  fmt: function(v) {{ return v.toFixed(2) + '%'; }} }},
    {{ key: 'app_size', label: 'Avg App Size (MB)',  fmt: function(v) {{ return v.toFixed(1) + ' MB'; }} }},
  ];
  var sub = active.length + ' brand' + (active.length !== 1 ? 's' : '') + ' · rider-weighted';
  document.getElementById('dash-summary').innerHTML = cards.map(function(c) {{
    return '<div class="dash-card">'
      + '<div class="dc-label">' + c.label + '</div>'
      + '<div class="dc-value">' + c.fmt(riderWeightedAvg(c.key)) + '</div>'
      + '<div class="dc-sub">' + sub + '</div></div>';
  }}).join('');
}}

// ── Chart helpers ─────────────────────────────────────────────────────────────
function hexToRgba(hex, a) {{
  var r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
  return 'rgba(' + r + ',' + g + ',' + b + ',' + a + ')';
}}
function darken(hex) {{
  var r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
  return 'rgb(' + Math.max(0,r-60) + ',' + Math.max(0,g-60) + ',' + Math.max(0,b-60) + ')';
}}

function renderCharts() {{
  var active = getActiveBrands();
  var labels = active.map(function(b) {{ return b.name; }});
  var bgColors = active.map(function(b) {{ return hexToRgba(b.color, 0.88); }});
  var bdColors = active.map(function(b) {{ return darken(b.color); }});

  if (chartCfu) chartCfu.destroy();
  chartCfu = new Chart(document.getElementById('chart-cfu'), {{
    type: 'bar',
    data: {{ labels: labels, datasets: [{{
      data: active.map(function(b) {{ return b.cfu; }}),
      backgroundColor: bgColors, borderColor: bdColors, borderWidth: 1, borderRadius: 4
    }}] }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{ legend: {{ display: false }}, tooltip: {{ callbacks: {{ label: function(ctx) {{ return ctx.raw.toFixed(3) + '%'; }} }} }} }},
      scales: {{ y: {{ min: 99, max: 100.1, ticks: {{ callback: function(v) {{ return v + '%'; }} }} }} }},
    }},
  }});

  if (chartHang) chartHang.destroy();
  chartHang = new Chart(document.getElementById('chart-hang'), {{
    type: 'bar',
    data: {{ labels: labels, datasets: [{{
      data: active.map(function(b) {{ return b.hang; }}),
      backgroundColor: bgColors, borderColor: bdColors, borderWidth: 1, borderRadius: 4
    }}] }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{ legend: {{ display: false }}, tooltip: {{ callbacks: {{ label: function(ctx) {{ return ctx.raw.toFixed(3) + '%'; }} }} }} }},
      scales: {{ y: {{ min: 99, max: 100.1, ticks: {{ callback: function(v) {{ return v + '%'; }} }} }} }},
    }},
  }});

  if (chartFrames) chartFrames.destroy();
  chartFrames = new Chart(document.getElementById('chart-frames'), {{
    type: 'bar',
    data: {{
      labels: labels,
      datasets: [
        {{ label: 'Frozen %', data: active.map(function(b) {{ return b.frozen; }}),
           backgroundColor: 'rgba(99,102,241,.75)', borderColor: 'rgb(79,70,229)', borderWidth: 1, borderRadius: 4 }},
        {{ label: 'Skipped %', data: active.map(function(b) {{ return b.skipped; }}),
           backgroundColor: 'rgba(251,146,60,.75)', borderColor: 'rgb(234,88,12)', borderWidth: 1, borderRadius: 4 }},
      ],
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{ legend: {{ display: true, position: 'top' }},
                  tooltip: {{ callbacks: {{ label: function(ctx) {{ return ctx.dataset.label + ': ' + ctx.raw.toFixed(2) + '%'; }} }} }} }},
      scales: {{ y: {{ ticks: {{ callback: function(v) {{ return v + '%'; }} }} }} }},
    }},
  }});

  if (chartAppsize) chartAppsize.destroy();
  chartAppsize = new Chart(document.getElementById('chart-appsize'), {{
    type: 'bar',
    data: {{ labels: labels, datasets: [{{
      data: active.map(function(b) {{ return b.app_size; }}),
      backgroundColor: bgColors, borderColor: bdColors, borderWidth: 1, borderRadius: 4
    }}] }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{ legend: {{ display: false }}, tooltip: {{ callbacks: {{ label: function(ctx) {{ return ctx.raw + ' MB'; }} }} }} }},
      scales: {{ y: {{ min: 50, ticks: {{ callback: function(v) {{ return v + ' MB'; }} }} }} }},
    }},
  }});
}}

// ── Brand detail table ────────────────────────────────────────────────────────
function renderTable() {{
  var active = getActiveBrands();
  var tbody = document.getElementById('dash-table-body');
  if (!active.length) {{
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;color:#9ca3af;padding:20px;">No brands selected</td></tr>';
    return;
  }}
  tbody.innerHTML = active.map(function(b) {{
    return '<tr style="background:' + b.color + ';">'
      + '<td>' + b.name + '</td>'
      + '<td>' + b.cfu.toFixed(2) + '%</td>'
      + '<td>' + b.hang.toFixed(2) + '%</td>'
      + '<td>' + b.app_size + '</td>'
      + '<td>' + b.asti + '</td>'
      + '<td>' + b.stti + '</td>'
      + '<td>' + b.frozen.toFixed(2) + '%</td>'
      + '<td>' + b.skipped.toFixed(2) + '%</td>'
      + '<td>' + b.riders.toLocaleString() + '</td>'
      + '</tr>';
  }}).join('');
}}

function renderAll() {{
  renderSummary();
  renderCharts();
  renderTable();
}}

// ── Brand pills ───────────────────────────────────────────────────────────────
function buildPills() {{
  document.getElementById('brand-pills').innerHTML = BRAND_DATA.map(function(b, i) {{
    return '<div class="brand-pill" id="pill-' + i + '" style="background:' + b.color + ';" onclick="toggleBrand(' + i + ')">' + b.name + '</div>';
  }}).join('');
}}

function toggleBrand(i) {{
  selected[i] = !selected[i];
  document.getElementById('pill-' + i).classList.toggle('inactive', !selected[i]);
  renderAll();
}}

function selectAllBrands() {{
  selected = BRAND_DATA.map(function() {{ return true; }});
  BRAND_DATA.forEach(function(b, i) {{ document.getElementById('pill-' + i).classList.remove('inactive'); }});
  renderAll();
}}

function clearAllBrands() {{
  selected = BRAND_DATA.map(function() {{ return false; }});
  BRAND_DATA.forEach(function(b, i) {{ document.getElementById('pill-' + i).classList.add('inactive'); }});
  renderAll();
}}

buildPills();
renderAll();
</script>

</body>
</html>"""

    os.makedirs(OUT_DIR, exist_ok=True)
    out_path = os.path.join(OUT_DIR, "index.html")
    with open(out_path, "w") as f:
        f.write(html)

    print(f"\nReport written to: {out_path}")


# ── Excel workbook generation ──────────────────────────────────────────────────

def add_consolidation_sheet(wb, firebase_data=None):
    """Second tab: per-brand summary.

    Crash free users (B) and App Hangs (C) are computed directly from the
    Dashboard tab's raw data (T-AF) using SUMIFS, so they always reflect
    whichever date range was fetched — no need to change B10 per brand.

    Frozen Frames (G) and Skipped Frames (H) come from firebase_data
    returned by fetch_firebase_frames(); falls back to static values
    when the BQ fetch is unavailable.
    """
    if firebase_data is None:
        firebase_data = {}
    ws = wb.create_sheet(title="Consolidation", index=0)

    headers = [
        "Brand", "Crash free users", "App Hangs", "App size",
        "ASTI", "STTI", "Frozen Frames", "Skipped Frames",
    ]
    # Column M (13) — Riders count per brand (order matches BRANDS list)
    RIDER_COL  = 13   # M
    # Column N (14) — fixed brand weights by rider share; order matches BRANDS list
    WEIGHT_COL = 14   # N

    # Static per-brand columns: (brand_name, app_size, asti, stti, frozen_fb, skipped_fb)
    # Brand name and numeric columns are derived from the module-level BRANDS list.
    # frozen_fb / skipped_fb are fallback values used when firebase_data is unavailable.
    static_cols = [
        (b["name"], b["app_size"], b["asti"], b["stti"], 0.58, 1.2)
        for b in BRANDS
    ]
    brands = [row[0] for row in static_cols]

    month_days = calendar.monthrange(START_DATE.year, START_DATE.month)[1]

    DATA_START = 2
    DATA_END   = DATA_START + len(brands) - 1   # row 11
    AVG_ROW    = DATA_END + 1                   # row 12
    AQS_ROW    = AVG_ROW + 1                    # row 13
    FINAL_ROW  = AQS_ROW + 2                    # row 15 (one blank gap)

    FINAL_BLUE = PatternFill("solid", fgColor="4472C4")

    # ── Header row ────────────────────────────────────────────────────────────
    for ci, h in enumerate(headers, start=1):
        cell = ws.cell(1, ci, value=h)
        cell.font = BOLD; cell.fill = DASH_HEADER_FILL
        cell.border = THIN_BOX
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    # ── Brand data rows ───────────────────────────────────────────────────────
    for i, (brand, app_size, asti, stti, frozen_fb, skipped_fb) in enumerate(static_cols):
        r = DATA_START + i
        row_fill = EVEN_ROW_FILL if r % 2 == 0 else None

        def styled(col, value=None):
            cell = ws.cell(r, col, value=value)
            cell.border = THIN_BOX
            if row_fill:
                cell.fill = row_fill
            return cell

        styled(1, brand).font = BOLD

        # B: Crash free % — SUMIFS on Dashboard raw data (no B10 dependency)
        # Fallback uses M (riders) × month_days to approximate user-days when BQ is unavailable.
        styled(2).value = (
            f'=TRUNC(MAX(0,MIN(100,(1-SUMIFS(Dashboard!$X:$X,Dashboard!$Z:$Z,"*"&$A{r}&"*")'
            f'/IF(SUMIFS(Dashboard!$V:$V,Dashboard!$U:$U,"*"&$A{r}&"*")>0,'
            f'SUMIFS(Dashboard!$V:$V,Dashboard!$U:$U,"*"&$A{r}&"*"),'
            f'M{r}*{month_days}))*100)),2)'
        )
        # C: Hang free %
        styled(3).value = (
            f'=TRUNC(MAX(0,MIN(100,(1-SUMIFS(Dashboard!$AB:$AB,Dashboard!$AD:$AD,"*"&$A{r}&"*")'
            f'/IF(SUMIFS(Dashboard!$V:$V,Dashboard!$U:$U,"*"&$A{r}&"*")>0,'
            f'SUMIFS(Dashboard!$V:$V,Dashboard!$U:$U,"*"&$A{r}&"*"),'
            f'M{r}*{month_days}))*100)),2)'
        )
        styled(4, app_size)
        styled(5, asti)
        styled(6, stti)

        # G: Frozen Frames — single aggregated value from BQ, same for all brands
        styled(7, firebase_data.get("frozen",  frozen_fb))

        # H: Skipped Frames — single aggregated value from BQ, same for all brands
        styled(8, firebase_data.get("skipped", skipped_fb))

    # ── Column M: Riders count ────────────────────────────────────────────────
    rcol = get_column_letter(RIDER_COL)
    mhdr = ws.cell(1, RIDER_COL, value="Riders count")
    mhdr.font = BOLD; mhdr.fill = SETTINGS_FILL; mhdr.border = THIN_BOX
    mhdr.alignment = Alignment(horizontal="center", wrap_text=True)
    for i, count in enumerate(RIDER_COUNTS):
        cell = ws.cell(DATA_START + i, RIDER_COL, value=count)
        cell.border = THIN_BOX
        if (DATA_START + i) % 2 == 0:
            cell.fill = EVEN_ROW_FILL

    # ── Column N: Weight Rider ID (fixed brand rider-share weights) ───────────
    wcol = get_column_letter(WEIGHT_COL)
    hdr = ws.cell(1, WEIGHT_COL, value="Weight Rider ID")
    hdr.font = BOLD; hdr.fill = SETTINGS_FILL; hdr.border = THIN_BOX
    hdr.alignment = Alignment(horizontal="center", wrap_text=True)
    for i, w in enumerate(WEIGHTS):
        cell = ws.cell(DATA_START + i, WEIGHT_COL, value=w)
        cell.number_format = "0.00%"
        cell.border = THIN_BOX
        if (DATA_START + i) % 2 == 0:
            cell.fill = EVEN_ROW_FILL

    # ── AVG row (row 12) — weighted average using rider-share weights in col N ─
    ws.cell(AVG_ROW, 1, value="AVG").font = BOLD
    ws.cell(AVG_ROW, 1).fill = SUMMARY_LBL_FILL
    ws.cell(AVG_ROW, 1).border = THIN_BOX
    for ci in range(2, 9):
        cl = get_column_letter(ci)
        cell = ws.cell(AVG_ROW, ci)
        cell.value = (f"=TRUNC(SUMPRODUCT({cl}{DATA_START}:{cl}{DATA_END},"
                      f"{wcol}{DATA_START}:{wcol}{DATA_END}),2)")
        cell.font = BOLD; cell.fill = SUMMARY_LBL_FILL; cell.border = THIN_BOX
        cell.number_format = "0.00"   # plain number, no % sign

    # ── AQS row (row 13) ─────────────────────────────────────────────────────
    aqs_formulas = {
        2: f"=ROUND(MIN(100,MAX(0,(((B{AVG_ROW}-99.7)/(99.9-99.7))*50)+50))*0.35,2)",
        3: f"=ROUND(MIN(100,MAX(0,(((C{AVG_ROW}-99.7)/(99.9-99.7))*50)+50))*0.25,2)",
        4: f"=ROUND(MIN(100,MAX(0,(((D{AVG_ROW}-75)/(60-75))*50)+50))*0.02,2)",
        5: f"=ROUND(MIN(100,MAX(0,(((E{AVG_ROW}-4)/(2-4))*50)+50))*0.10,2)",
        6: f"=ROUND(MIN(100,MAX(0,(((F{AVG_ROW}-1.5)/(0.5-1.5))*50)+50))*0.10,2)",
        7: f"=ROUND(MIN(100,MAX(0,(((G{AVG_ROW}-3)/(1-3))*50)+50))*0.13,2)",
        8: f"=ROUND(MIN(100,MAX(0,(((H{AVG_ROW}-2)/(1-2))*50)+50))*0.05,2)",
    }
    ws.cell(AQS_ROW, 1, value="AQS").font = BOLD
    ws.cell(AQS_ROW, 1).fill = SUMMARY_VAL_FILL
    ws.cell(AQS_ROW, 1).border = THIN_BOX
    for ci, val in aqs_formulas.items():
        cell = ws.cell(AQS_ROW, ci, value=val)
        cell.font = BOLD; cell.fill = SUMMARY_VAL_FILL; cell.border = THIN_BOX
        cell.number_format = "0.00"   # plain number, no % sign

    # M13: sum of M column data rows
    m13 = ws.cell(AQS_ROW, 13, value=f"=SUM(M{DATA_START}:M{DATA_END})")
    m13.font = BOLD; m13.fill = SUMMARY_VAL_FILL; m13.border = THIN_BOX

    # N13: sum of weights (validation — should equal 100%)
    n13 = ws.cell(AQS_ROW, WEIGHT_COL, value=f"=SUM({wcol}{DATA_START}:{wcol}{DATA_END})")
    n13.font = BOLD; n13.fill = SETTINGS_FILL; n13.border = THIN_BOX
    n13.number_format = "0.00%"

    # ── FINAL AQS ─────────────────────────────────────────────────────────────
    ws.cell(FINAL_ROW, 1, value="FINAL AQS =")
    ws.cell(FINAL_ROW, 1).font = Font(bold=True, color="FFFFFF")
    ws.cell(FINAL_ROW, 1).fill = FINAL_BLUE
    ws.cell(FINAL_ROW, 1).border = MEDIUM_BOX
    ws.cell(FINAL_ROW, 1).alignment = Alignment(horizontal="right")

    final_cell = ws.cell(FINAL_ROW, 2)
    final_cell.value = f"=ROUND(SUM(B{AQS_ROW}:H{AQS_ROW}),2)"
    final_cell.font = Font(bold=True, color="FFFFFF", size=13)
    final_cell.fill = FINAL_BLUE
    final_cell.border = MEDIUM_BOX
    final_cell.number_format = "0.00"

    # ── Manual-data note ──────────────────────────────────────────────────────
    NOTE_TEXT = ("⚠  App size, ASTI and STTI have the older values, not recently updated. "
                 "Update these columns manually before sharing. "
                 "For the accurate AQS score, update these values.")
    NOTE_FILL = PatternFill("solid", fgColor="FFEB9C")   # amber
    NOTE_FONT = Font(bold=True, color="9C5700")           # dark orange

    NOTE_ROW = FINAL_ROW + 2
    note_cell = ws.cell(NOTE_ROW, 1, value=NOTE_TEXT)
    note_cell.font = NOTE_FONT
    note_cell.fill = NOTE_FILL
    note_cell.border = MEDIUM_BOX
    note_cell.alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=NOTE_ROW, start_column=1,
                   end_row=NOTE_ROW,   end_column=8)
    ws.row_dimensions[NOTE_ROW].height = 30

    # Also attach pop-up comments to the D, E, F column headers
    MANUAL_COMMENT = ("Older values, not recently updated.\n"
                      "Update manually for accurate AQS score.")
    for col_letter in ("D", "E", "F"):
        ws[f"{col_letter}1"].comment = Comment(MANUAL_COMMENT, "Script")

    # ── Column widths & freeze ────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 18
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 17
    ws.column_dimensions[rcol].width = 15  # M — Riders count
    ws.column_dimensions[wcol].width = 16  # N — Weight Rider ID
    ws.row_dimensions[1].height = 40
    ws.freeze_panes = "B2"


def write_excel(bq_rows, hang_rows, crash_rows, path, firebase_data=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dashboard"

    bq_cols    = ["dt", "appId", "user_count"]
    crash_cols = ["CRASH_USERS", "day", "Brand"]
    hang_cols  = ["HANG_USERS",  "day", "Brand"]

    # ── Raw data (T-AD) ────────────────────────────────────────────────────────
    # BQ:    T(20)=dt, U(21)=appId, V(22)=user_count
    # Crash: X(24)=CRASH_USERS, Y(25)=day, Z(26)=Brand
    # Hang:  AB(28)=HANG_USERS, AC(29)=day, AD(30)=Brand

    for ci, name in enumerate(bq_cols, start=BQ_START_COL):
        cell = ws.cell(row=1, column=ci, value=name)
        cell.font = BOLD; cell.fill = BQ_HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for ci, name in enumerate(crash_cols, start=CRASH_START_COL):
        cell = ws.cell(row=1, column=ci, value=name)
        cell.font = BOLD; cell.fill = CRASH_HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for ci, name in enumerate(hang_cols, start=HANG_START_COL):
        cell = ws.cell(row=1, column=ci, value=name)
        cell.font = BOLD; cell.fill = HANG_HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for ri, row in enumerate(bq_rows, start=2):
        for ci, key in enumerate(bq_cols, start=BQ_START_COL):
            ws.cell(row=ri, column=ci, value=row[key])

    for ri, row in enumerate(crash_rows, start=2):
        for ci, key in enumerate(crash_cols, start=CRASH_START_COL):
            ws.cell(row=ri, column=ci, value=row[key])

    for ri, row in enumerate(hang_rows, start=2):
        for ci, key in enumerate(hang_cols, start=HANG_START_COL):
            ws.cell(row=ri, column=ci, value=row[key])

    for ci in range(BQ_START_COL, BQ_START_COL + 3):
        ws.column_dimensions[get_column_letter(ci)].width = 22
    ws.column_dimensions["W"].width = 4   # spacer between BQ and Crash sections
    for ci in range(CRASH_START_COL, CRASH_START_COL + 3):
        ws.column_dimensions[get_column_letter(ci)].width = 22
    ws.column_dimensions["AA"].width = 4  # spacer between Crash and Hang sections
    for ci in range(HANG_START_COL, HANG_START_COL + 3):
        ws.column_dimensions[get_column_letter(ci)].width = 22

    # ── Dashboard (A-N) ────────────────────────────────────────────────────────
    # Settings: column A = label, column B = value, rows 1-10
    settings = [
        ("CFU part in AQS",  60),      # B1
        ("base AQS score",   32.65),   # B2
        ("CFU minimum",      99.5),    # B3
        ("CFU maximum",      99.9),    # B4
        (None,               None),
        ("Hang minimum",     99.5),    # B6
        ("Hang maximum",     99.9),    # B7
        ("Hang part in AQS", 25),      # B8
        (None,               None),
        ("Brand",            "woowa"), # B10 — change to filter by brand
    ]
    for i, (label, value) in enumerate(settings, start=1):
        if label:
            a = ws.cell(row=i, column=1, value=label)
            a.font = BOLD; a.fill = SETTINGS_FILL
            b = ws.cell(row=i, column=2, value=value)
            b.fill = SETTINGS_FILL

    # Brand dropdown in B10
    BRAND_DROPDOWN = ["foodora", "woowa", "foodpanda", "talabat", "efood", "glovo",
                      "hungerstation", "yemek", "foody", "pedidosya"]
    brand_dv = DataValidation(
        type="list",
        formula1='"' + ",".join(BRAND_DROPDOWN) + '"',
        allow_blank=False,
        showDropDown=False,
    )
    ws.add_data_validation(brand_dv)
    brand_dv.add("B10")

    # Column headers (row 1, columns C-N)
    col_headers = {
        3:  "Date",
        4:  "Total users count (ios_User_Count)",
        5:  "Crashed users count (Query)",
        7:  "CFU %",
        11: "Hang user count (Query)",
        12: "Hang free %",
        14: "Projected AQS including hangs",
    }
    for col, header in col_headers.items():
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = BOLD; cell.fill = DASH_HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    # All calendar days of current month in column C (rows 2 to month_days+1)
    month_days = calendar.monthrange(START_DATE.year, START_DATE.month)[1]
    for day in range(1, month_days + 1):
        ws.cell(row=day + 1, column=3, value=START_DATE.replace(day=day).strftime("%Y-%m-%d"))

    DATA_END    = month_days + 1       # last date row (all calendar days filled in col C)
    FORMULA_END = END_DATE.day + 1    # last row with formulas — up to yesterday only

    # Per-day formulas only for rows with actual data (day 1 → yesterday)
    for r in range(2, FORMULA_END + 1):
        ws.cell(r, 4).value  = f'=SUMIFS(V:V,U:U,"*"&$B$10&"*",T:T,C{r})'
        ws.cell(r, 5).value  = f'=SUMIFS(X:X,Z:Z,"*"&$B$10&"*",Y:Y,C{r})'
        ws.cell(r, 7).value  = f'=ROUND(MAX(0,MIN(100,(1-E{r}/IF(D{r}>0,D{r},SUMIFS(Consolidation!$M:$M,Consolidation!$A:$A,"*"&$B$10&"*")/{month_days}))*100)),2)'
        ws.cell(r, 8).value  = f'=MIN(1,(G{r}-$B$3)/($B$4-$B$3))*$B$1'
        ws.cell(r, 11).value = f'=SUMIFS(AB:AB,AD:AD,"*"&$B$10&"*",AC:AC,C{r})'
        ws.cell(r, 12).value = f'=ROUND(MAX(0,MIN(100,100*(1-K{r}/IF(D{r}>0,D{r},SUMIFS(Consolidation!$M:$M,Consolidation!$A:$A,"*"&$B$10&"*")/{month_days})))),2)'
        ws.cell(r, 13).value = f'=MAX(0,(L{r}-$B$6)/($B$7-$B$6)*$B$8)'
        ws.cell(r, 14).value = f'=$B$2+H{r}+M{r}'

    # Summary label row and value row below all calendar dates
    LBL = DATA_END + 1
    VAL = DATA_END + 2
    month_name = START_DATE.strftime("%B")

    summaries = [
        (4,  None,                            f"=AVERAGE(D2:D{FORMULA_END})"),
        (7,  "AVG CFU",                       f'=TRUNC(MAX(0,MIN(100,(1-SUMIFS(X:X,Z:Z,"*"&$B$10&"*")/IF(SUMIFS(V:V,U:U,"*"&$B$10&"*")>0,SUMIFS(V:V,U:U,"*"&$B$10&"*"),SUMIFS(Consolidation!$M:$M,Consolidation!$A:$A,"*"&$B$10&"*")*{month_days}))*100)),2)'),
        (8,  None,                            f"=AVERAGE(H2:H{FORMULA_END})"),
        (9,  f"{month_name} AQS score",       f"=$B$2+H{VAL}"),
        (12, "AVG Hang-free",                 f'=TRUNC(MAX(0,MIN(100,(1-SUMIFS(AB:AB,AD:AD,"*"&$B$10&"*")/IF(SUMIFS(V:V,U:U,"*"&$B$10&"*")>0,SUMIFS(V:V,U:U,"*"&$B$10&"*"),SUMIFS(Consolidation!$M:$M,Consolidation!$A:$A,"*"&$B$10&"*")*{month_days}))*100)),2)'),
        (13, None,                            f"=MAX(0,(L{VAL}-$B$6)/($B$7-$B$6)*$B$8)"),
        (14, "Projected AQS including hangs", f"=$B$2+MAX(0,(G{VAL}-$B$3)/($B$4-$B$3)*($B$1-$B$8))+M{VAL}"),
    ]
    for col, label, formula in summaries:
        if label:
            lc = ws.cell(row=LBL, column=col, value=label)
            lc.font = BOLD
        ws.cell(row=VAL, column=col, value=formula)

    ws.cell(row=VAL, column=4).number_format  = "0"      # rider count — integer
    ws.cell(row=VAL, column=7).number_format  = "0.##"   # AVG CFU % — up to 2 decimal places
    ws.cell(row=VAL, column=8).number_format  = "0.00"   # AVG CFU AQS (hidden helper)
    ws.cell(row=VAL, column=9).number_format  = "0.00"   # month AQS score
    ws.cell(row=VAL, column=12).number_format = "0.##"   # AVG Hang-free % — up to 2 decimal places
    ws.cell(row=VAL, column=13).number_format = "0.00"   # Hang-free AQS (hidden helper)
    ws.cell(row=VAL, column=14).number_format = "0.00"   # Projected AQS

    # ── Borders, alternating rows & summary colors ───────────────────────────────

    # Settings region A1:B10
    for r in range(1, 11):
        for c in (1, 2):
            ws.cell(r, c).border = THIN_BOX

    # Dashboard header row C1:N1
    for c in range(3, 15):
        ws.cell(1, c).border = THIN_BOX

    # Data rows C-N: alternating row fill + thin borders
    for r in range(2, DATA_END + 1):
        for c in range(3, 15):
            cell = ws.cell(r, c)
            cell.border = THIN_BOX
            if r % 2 == 0:
                cell.fill = EVEN_ROW_FILL

    # Cols A-B for rows 11-DATA_END: thin border only (empty settings area)
    for r in range(11, DATA_END + 1):
        for c in (1, 2):
            ws.cell(r, c).border = THIN_BOX

    # Summary label row
    for c in range(4, 15):
        cell = ws.cell(LBL, c)
        cell.fill = SUMMARY_LBL_FILL
        cell.border = MEDIUM_BOX

    # Summary value row
    for c in range(4, 15):
        cell = ws.cell(VAL, c)
        cell.fill = SUMMARY_VAL_FILL
        cell.border = MEDIUM_BOX

    # Raw data section (T-AF): header + data borders with alternating rows
    max_raw = max(len(bq_rows), len(crash_rows), len(hang_rows)) + 1
    raw_cols = (list(range(BQ_START_COL, BQ_START_COL + 3)) +
                list(range(CRASH_START_COL, CRASH_START_COL + 3)) +
                list(range(HANG_START_COL, HANG_START_COL + 3)))
    for r in range(1, max_raw + 1):
        for c in raw_cols:
            cell = ws.cell(r, c)
            cell.border = THIN_BOX
            if r > 1 and r % 2 == 0:
                cell.fill = EVEN_ROW_FILL

    # Freeze the header row
    ws.freeze_panes = "C2"

    # Dashboard column widths
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 4   # spacer
    ws.column_dimensions["G"].width = 12
    ws.column_dimensions["H"].hidden = True   # CFU AQS Score — hidden helper for col N
    ws.column_dimensions["I"].width = 22
    ws.column_dimensions["J"].width = 4   # spacer
    ws.column_dimensions["K"].width = 24
    ws.column_dimensions["L"].width = 14
    ws.column_dimensions["M"].hidden = True   # Hang-free AQS — hidden helper for col N
    ws.column_dimensions["N"].width = 30
    ws.row_dimensions[1].height = 45

    add_consolidation_sheet(wb, firebase_data=firebase_data)

    # Make Consolidation the first visible tab when the workbook is opened
    wb.active = wb.worksheets[0]  # index 0 = Consolidation (inserted first)

    wb.save(path)
    print(f"Wrote {path}  ({len(bq_rows)} BQ rows, {len(crash_rows)} crash rows, {len(hang_rows)} hang rows)")


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    print(f"Fetching data for {BQ_START} → {BQ_END}\n")

    print("Fetching BigQuery iOS user counts...")
    bq_rows = fetch_bigquery()

    print("Fetching Sentry hangs (per-brand queries)...")
    raw_hang_rows = fetch_discover_per_brand(HANGS_QUERY)
    hang_rows = shape_rows(raw_hang_rows, "HANG_USERS")
    backfill_zero_rows(hang_rows, "HANG_USERS", HANGS_QUERY)

    print("Fetching Sentry crashes (per-brand queries)...")
    raw_crash_rows = fetch_discover_per_brand(CRASHES_QUERY, environment="production")
    crash_rows = shape_rows(raw_crash_rows, "CRASH_USERS")
    backfill_zero_rows(crash_rows, "CRASH_USERS", CRASHES_QUERY, environment="production")

    print("Fetching Firebase Performance frames...")
    firebase_data = fetch_firebase_frames()

    # Aggregate per-brand totals for HTML report (no extra network calls)
    bq_users_by_brand = {}
    for row in bq_rows:
        key = (row["appId"] or "").lower()
        bq_users_by_brand[key] = bq_users_by_brand.get(key, 0) + row["user_count"]

    crash_by_brand = aggregate_by_brand(crash_rows, "CRASH_USERS")
    hang_by_brand  = aggregate_by_brand(hang_rows,  "HANG_USERS")

    os.makedirs("consolidation_report", exist_ok=True)
    write_excel(bq_rows, hang_rows, crash_rows, "consolidation_report/sentry_data.xlsx", firebase_data=firebase_data)

    generate_html_report(bq_users_by_brand, crash_by_brand, hang_by_brand, firebase_data)


if __name__ == "__main__":
    main()
