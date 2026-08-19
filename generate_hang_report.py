#!/usr/bin/env python3
"""
Sentry App Hang Analysis Report Generator — Rider iOS (Multi-Month)
=====================================================================
Fetches deduplicated AppHang issue counts from Sentry per month and generates
a self-contained HTML dashboard with a comparison overview and per-month detail.

Usage:
    export SENTRY_AUTH_TOKEN=sntryu_...
    python3 generate_hang_report.py

Output:
    hang_report/index.html
"""

import urllib.request
import urllib.parse
import json
import os
import time
import datetime

try:
    from google.cloud import bigquery as _bq
    BQ_AVAILABLE = True
except ImportError:
    BQ_AVAILABLE = False
    print("WARNING: google-cloud-bigquery not installed — device class fleet totals will be skipped.")

# ── Config ────────────────────────────────────────────────────────────────────

TOKEN   = os.environ.get("SENTRY_AUTH_TOKEN", "")
ORG     = "delivery-hero-pm"
PROJECT = "4506937839976448"
OUT_DIR = "hang_report"

# Pin specific releases to compare in the By Release tab.
# Each entry needs a version string and the Sentry dist (build number).
# Set to [] to auto-fetch the most recent releases from Sentry instead.
PINNED_RELEASES = [
    {"version": "4.2630.1", "dist": "1026"},
    {"version": "4.2631.1", "dist": "1032"},
    {"version": "4.2632.4", "dist": "1040"},
    {"version": "4.2633.1", "dist": "1044"},
    {"version": "4.2634.1", "dist": "1049"},
]

# Release Weekly uses title-based search rather than mechanism filter.
RELEASE_WEEKLY_QUERY = 'is:unresolved "*App Hang* detected*"'

if not TOKEN:
    raise SystemExit("ERROR: set SENTRY_AUTH_TOKEN environment variable first.")

# ── Months to analyze ─────────────────────────────────────────────────────────

TODAY = datetime.date.today()

def add_months(d, delta):
    """Return the first day of the month shifted by delta months."""
    month_index = d.year * 12 + (d.month - 1) + delta
    year = month_index // 12
    month = month_index % 12 + 1
    return datetime.date(year, month, 1)


def get_months():
    """Return current month + previous 2 months (3-month rolling window)."""
    month_count = 3
    current_month_start = TODAY.replace(day=1)
    start_month = add_months(current_month_start, -(month_count - 1))

    months = []
    for i in range(month_count):
        start_dt = add_months(start_month, i)
        end_dt = add_months(start_dt, 1)

        # Current month is partial and should include data up to today.
        if start_dt.year == TODAY.year and start_dt.month == TODAY.month:
            end_dt = TODAY + datetime.timedelta(days=1)

        months.append({
            "label": start_dt.strftime("%B %Y"),
            "short": start_dt.strftime("%b %Y"),
            "start": start_dt.strftime("%Y-%m-%dT00:00:00"),
            "end":   end_dt.strftime("%Y-%m-%dT00:00:00"),
        })

    return months


MONTHS = get_months()

# ── Weeks to analyze for weekly trend tab ─────────────────────────────────────

def get_last_4_weeks():
    """
    Return the latest 4 weekly trend windows.

    Weeks are Monday-Sunday. If the current week is not complete yet,
    the latest bucket is Monday-today and is marked as partial.
    """
    weeks = []

    current_week_start = TODAY - datetime.timedelta(days=TODAY.weekday())
    current_week_end = TODAY

    # Include current week as the latest bucket. It is partial unless today is Sunday.
    for i in range(4):
        start_dt = current_week_start - datetime.timedelta(days=i * 7)
        if i == 0:
            end_dt = current_week_end
            is_partial = TODAY.weekday() != 6
        else:
            end_dt = start_dt + datetime.timedelta(days=6)
            is_partial = False

        suffix = " (partial)" if is_partial else ""
        weeks.append({
            "label": f"{start_dt.strftime('%b %d')} - {end_dt.strftime('%b %d')}{suffix}",
            "short": f"{start_dt.strftime('%b %d')}",
            "start": start_dt.strftime("%Y-%m-%dT00:00:00"),
            "end":   (end_dt + datetime.timedelta(days=1)).strftime("%Y-%m-%dT00:00:00"),
            "partial": is_partial,
        })

    return list(reversed(weeks))


WEEKS = get_last_4_weeks()

# ── Exclusion patterns ────────────────────────────────────────────────────────

NX = {
    "mapbox":   "!stack.package:*Mapbox* !message:*Mapbox*",
    "naver":    "!stack.function:*NavigationMapView*",
    "webkit":   "!stack.package:*WebKit* !stack.package:*WebCore* !message:*WKWebView*",
    "firebase": "!message:*FIRCLS* !message:*FireApp*",
    "sentry":   "!stack.package:*Sentry* !message:*SentryAppHang*",
    "keyboard": "!stack.function:*UIKeyboard*",
}

def excl(*keys):
    return " ".join(NX[k] for k in keys)

BASE_QUERY = 'is:unresolved "*App Hang* detected*"'
# Discover events API doesn't support is:unresolved (issue-level filter).
DISCOVER_BASE_QUERY = '"*App Hang* detected*"'

# ── Category definitions ──────────────────────────────────────────────────────

CATEGORIES = [
    {
        "key":      "mapbox",
        "label":    "Mapbox",
        "color":    "#FFB3BA", # Pastel Rose
        "filters":  ["stack.package:*Mapbox*", "stack.function:*Mapbox*"],
        "excl":     "",
        "link_filter": "stack.package:*Mapbox*",
        "culprits": [
            "MetalView.draw", "MetalView.nextDrawable", "MapboxMap.init",
            "MapView.commonInit", "MBMStyleManager", "InfoButtonOrnament.init",
            "MapboxLocationProvider.locationManager",
            "MapboxSpeechSynthesizer.init", "MapboxSpeechSynthesizer.speak",
        ],
    },
    {
        "key":      "naver",
        "label":    "Naver",
        "color":    "#BAFFC9", # Pastel Mint
        "filters":  [],
        "excl":     excl("mapbox"),
        "link_filter": "",
        "culprits": [
            "NMFMapView.init", "NMFCameraUpdate",
        ],
    },
    {
        "key":      "webkit",
        "label":    "WebKit / WebView",
        "color":    "#BAE1FF", # Pastel Sky Blue
        "filters":  ["stack.package:*WebKit*", "stack.function:*WKWebView*"],
        "excl":     excl("mapbox", "naver"),
        "link_filter": "stack.package:*WebKit*",
        "culprits": [
            "WebViewController.webView.getter", "WKWebView.load",
            "WKWebView.evaluateJavaScript", "WebCore::FrameLoader",
        ],
    },
    {
        "key":      "firebase",
        "label":    "Firebase / Crashlytics",
        "color":    "#FFFFBA", # Pastel Lemon
        "filters":  [
            "message:*FIRCLSFileLoop*", "message:*FIRCLSProcess*",
            "message:*FireApp*", "stack.function:*FIRCLS*", "stack.function:*FireApp*",
        ],
        "excl":     excl("mapbox", "naver", "webkit") + " !stack.function:*AppleLocationProvider*",
        "link_filter": "message:*FIRCLS*",
        "culprits": [
            "FIRCLSFileLoopWithWriteBlock", "FIRCLSProcessRecordAllThreads",
            "FireAppService.init",
        ],
    },
    {
        "key":      "keyboard",
        "label":    "Keyboard",
        "color":    "#E0BBE4", # Pastel Lavender
        "filters":  ["stack.function:*UIKeyboard*"],
        "excl":     excl("mapbox", "naver", "webkit", "firebase"),
        "link_filter": "stack.function:*UIKeyboard*",
        "culprits": [],
    },
    {
        "key":      "location",
        "label":    "Location",
        "color":    "#B2F2EF", # Pastel Turquoise
        "filters":  ["stack.package:*CoreLocation*", "stack.function:*CLLocation*", "stack.function:*CLClientCreateWithBundleIdentifierAndPathWithWebsiteOnSilo*"],
        "excl":     excl("mapbox", "naver", "webkit", "firebase", "sentry", "keyboard"),
        "link_filter": "stack.package:*CoreLocation*",
        "culprits": [
            "AppleLocationProvider.init", "LocationPermission",
            "CLLocationManager.startUpdatingLocation", "LocationServiceImpl.scheduleTimer",
        ],
    },
    {
        "key":      "audio",
        "label":    "Audio / Media",
        "color":    "#FDFDCC", # Pastel Cream
        "filters":  [
            "stack.package:*AVFoundation*", "stack.function:*AVAudio*",
            "stack.function:*AudioPlayer*",
        ],
        "excl":     excl("mapbox", "naver", "webkit", "firebase", "sentry", "keyboard"),
        "link_filter": "stack.package:*AVFoundation*",
        "culprits": [
            "AudioPlayerImpl.registerRemoteSyncProvider",
            "AVAudioSession.setCategory", "AVAudioSession.setActive",
        ],
    },
    {
        "key":      "camera",
        "label":    "Camera / AR",
        "color":    "#97EDEA", # Pastel Aqua
        "filters":  [
            "stack.package:*ARKit*", "stack.function:*ARSession*",
            "stack.function:*CameraProvider*", "stack.function:*QRScanCameraOverlayView*",
        ],
        "excl":     excl("mapbox", "naver", "webkit", "firebase", "sentry", "keyboard"),
        "link_filter": "stack.function:*QRScanCameraOverlayView*",
        "culprits": [
            "QRScanCameraOverlayView",
            "ARView.init", "ARPhotoManagerImpl.convertPixelBufferToUIImage",
            "ARSession.run", "CameraProviderImpl",
        ],
    },
    {
        "key":      "storage",
        "label":    "Storage",
        "color":    "#D3C0B0", # Pastel Taupe
        "filters":  ["stack.package:*CoreData*", "stack.function:*PersistentData*"],
        "excl":     excl("mapbox", "naver", "webkit", "firebase", "sentry", "keyboard"),
        "link_filter": "stack.package:*CoreData*",
        "culprits": [
            "PersistentDataActorImpl.__allocating_init", "PersistentDataActorImpl.init",
        ],
    },
]

# ── Sentry fetch helpers ───────────────────────────────────────────────────────

def fetch_issues(query, start, end, cursor=None):
    params = {
        "project": PROJECT,
        "query":   query,
        "start":   start,
        "end":     end,
        "limit":   100,
    }
    if cursor:
        params["cursor"] = cursor
    url = (f"https://sentry.io/api/0/organizations/{ORG}/issues/?"
           + urllib.parse.urlencode(params))
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {TOKEN}"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read())
        link = resp.headers.get("Link", "")
        next_cursor = None
        for part in link.split(","):
            if 'rel="next"' in part and 'results="true"' in part:
                for seg in part.split(";"):
                    seg = seg.strip()
                    if seg.startswith("<"):
                        next_cursor = urllib.parse.parse_qs(
                            urllib.parse.urlparse(seg[1:-1]).query
                        ).get("cursor", [None])[0]
        return data, next_cursor


def fetch_all(query, start, end):
    """Fetch all pages for a query, returning {issue_id: user_count}."""
    issues = {}
    cursor = None
    while True:
        batch, cursor = fetch_issues(query, start, end, cursor)
        for issue in batch:
            iid = issue["id"]
            u = issue.get("userCount") or 0
            if iid not in issues or issues[iid] < u:
                issues[iid] = u
        if not cursor:
            break
        time.sleep(0.4)
    return issues


def count_unique_users(query, start, end):
    """Count unique users via Discover /events/ API — SDK 9.9.0 raw totals."""
    params = urllib.parse.urlencode([
        ("project", PROJECT),
        ("query",   query),
        ("start",   start),
        ("end",     end),
        ("field",   "count_unique(user)"),
        ("dataset", "discover"),
    ])
    url = f"https://sentry.io/api/0/organizations/{ORG}/events/?{params}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {TOKEN}"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read())
    rows = data.get("data", [])
    return int(rows[0].get("count_unique(user)", 0)) if rows else 0


# ── Device class helpers ──────────────────────────────────────────────────────
# Uses Sentry's native device.class field: high | medium | low

DEVICE_TIERS = [
    {"key": "high",   "label": "High-end",  "color": "#6366f1"},
    {"key": "medium", "label": "Mid-range", "color": "#f59e0b"},
    {"key": "low",    "label": "Low-end",   "color": "#94a3b8"},
]


def fetch_bq_device_class_totals():
    """
    Distinct iOS riders per device class for yesterday.
    Returns {"high": N, "medium": N, "low": N} or all zeros if BQ is unavailable.
    """
    if not BQ_AVAILABLE:
        return {"high": 0, "medium": 0, "low": 0}
    query = """
        SELECT
          CASE
            WHEN NOT REGEXP_CONTAINS(device.mobileDeviceMarketingName, r'iPhone') THEN 'medium'
            -- SE 3rd gen has A15 chip (= iPhone 13 class) → high
            WHEN REGEXP_CONTAINS(device.mobileDeviceMarketingName, r'iPhone SE.*3rd') THEN 'high'
            -- iPhone X era (X, XR, XS, XS Max): no clean model number → medium
            WHEN REGEXP_CONTAINS(device.mobileDeviceMarketingName, r'iPhone X') THEN 'medium'
            -- Matches Sentry device.class: high = A15+ (iPhone 13+), medium = A13–A14 (11–12), low = A12 and below
            WHEN SAFE_CAST(
              SUBSTR(REGEXP_REPLACE(device.mobileDeviceMarketingName, r'[^0-9]', ''), 1, 2) AS INT64
            ) >= 13 THEN 'high'
            WHEN SAFE_CAST(
              SUBSTR(REGEXP_REPLACE(device.mobileDeviceMarketingName, r'[^0-9]', ''), 1, 2) AS INT64
            ) >= 11 THEN 'medium'
            ELSE 'low'
          END AS device_class,
          COUNT(DISTINCT clientId) AS daily_count
        FROM `fulfillment-dwh-production.curated_data_shared_coredata_tracking.perseus_events_rider_app`
        WHERE partition_date = DATE_SUB(CURRENT_DATE(), INTERVAL 1 DAY)
          AND platform = 'iOS'
        GROUP BY device_class
    """
    try:
        client = _bq.Client(project="logistics-rider-staging")
        rows = list(client.query(query).result())
        totals = {"high": 0, "medium": 0, "low": 0}
        for row in rows:
            if row.device_class in totals:
                totals[row.device_class] = row.daily_count
        return totals
    except Exception as e:
        print(f"WARNING: BigQuery device class fetch failed: {e}")
        return {"high": 0, "medium": 0, "low": 0}


def rel_filter_str(rel):
    """Build Sentry filter for a release entry (handles single dist or multiple dists)."""
    dists = rel.get("dists")
    d = rel.get("dist")
    v = rel.get("version")
    if dists:
        return "(" + " OR ".join(f"dist:{dd}" for dd in dists) + ")"
    elif d:
        return f"dist:{d}"
    return f"release:{v}"


def fetch_recent_releases(limit=8):
    """Return PINNED_RELEASES if configured, otherwise fetch from Sentry API."""
    if PINNED_RELEASES:
        result = []
        for r in PINNED_RELEASES:
            entry = {"version": r["version"], "label": r["version"], "short": r["version"]}
            if r.get("dists"):
                entry["dist"] = None
                entry["dists"] = r["dists"]
            else:
                entry["dist"] = r.get("dist")
            result.append(entry)
        return result
    params = urllib.parse.urlencode([
        ("project", PROJECT),
        ("limit",   str(limit)),
        ("start",   WEEKS[0]["start"]),
        ("end",     WEEKS[-1]["end"]),
        ("orderby", "date"),
    ])
    url = f"https://sentry.io/api/0/organizations/{ORG}/releases/?{params}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {TOKEN}"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read())
    except Exception as e:
        print(f"WARNING: Could not fetch releases: {e}")
        return []
    return [{"version": r["version"], "dist": None, "label": r["version"], "short": r["version"]}
            for r in data if r.get("version")]


RELEASES            = fetch_recent_releases()
RELEASE_WINDOW_START = (TODAY - datetime.timedelta(days=90)).strftime("%Y-%m-%dT00:00:00")
RELEASE_WINDOW_END   = TODAY.strftime("%Y-%m-%dT00:00:00")


# ── Run analysis for each month ───────────────────────────────────────────────

fetch_date = TODAY.strftime("%-d %b %Y")
all_month_results = []

for month in MONTHS:
    print(f"\n── {month['label']} ({month['start'][:10]} → {month['end'][:10]}) ──")

    assigned = set()
    results  = []

    for cat in CATEGORIES:
        cat_issues = {}
        for filt in cat["filters"]:
            q = f"{BASE_QUERY} {cat['excl']} {filt}".strip()
            try:
                fetched = fetch_all(q, month["start"], month["end"])
                for iid, u in fetched.items():
                    if iid not in cat_issues or cat_issues[iid] < u:
                        cat_issues[iid] = u
                time.sleep(0.5)
            except Exception as e:
                print(f"  ERROR {cat['label']} / {filt}: {e}")
                time.sleep(3)

        new_issues = {iid: u for iid, u in cat_issues.items() if iid not in assigned}
        already    = len(cat_issues) - len(new_issues)
        cat_users  = sum(new_issues.values())
        assigned.update(new_issues.keys())

        results.append({**cat, "users": cat_users})
        print(f"  {cat['label']:30s} {cat_users:6} users  "
              f"({len(new_issues)} issues, {already} absorbed by higher priority)")

    total_unique = sum(r["users"] for r in results)
    print(f"  {'TOTAL':30s} {total_unique:6} users")

    # Build Sentry link queries
    cumulative_excl = ""
    for r in results:
        r["link_query"] = f'{BASE_QUERY} {cumulative_excl} {r["link_filter"]}'.strip()
        nx_key = r["key"] if r["key"] in NX else None
        if nx_key:
            cumulative_excl = (cumulative_excl + " " + NX[nx_key]).strip()

    all_month_results.append({
        "month":   month,
        "results": results,
        "total":   total_unique,
    })



# ── Run release analysis for by-release tab ──────────────────────────────────
# Total hangs per release — no category breakdown.
# Uses Discover API for true per-dist unique user counts.

all_release_results = []

for rel in RELEASES:
    v          = rel["version"]
    rel_filter = rel_filter_str(rel)
    link_query = f'is:unresolved "*App Hang* detected*" {rel_filter}'.strip()
    discover_q = f'"*App Hang* detected*" {rel_filter}'.strip()
    dists_str  = ",".join(rel["dists"]) if rel.get("dists") else rel.get("dist")
    print(f"\n── Release {v} (dist {dists_str}) ──")
    try:
        users = count_unique_users(discover_q, RELEASE_WINDOW_START, RELEASE_WINDOW_END)
        time.sleep(0.5)
    except Exception as e:
        print(f"  ERROR: {e}")
        users = 0
    print(f"  {users} unique users")
    all_release_results.append({
        "release":    rel,
        "users":      users,
        "link_query": link_query,
    })


# ── Run release weekly analysis ───────────────────────────────────────────────

RELEASE_WEEKLY_RELEASES = RELEASES
all_release_weekly = []

for rel in RELEASE_WEEKLY_RELEASES:
    v          = rel["version"]
    rel_filter = rel_filter_str(rel)
    link_query = f'{RELEASE_WEEKLY_QUERY} {rel_filter}'.strip()
    print(f"\n── Release weekly: {v} ──")

    week_users = []
    for week in WEEKS:
        q = f'"*App Hang* detected*" {rel_filter}'.strip()
        try:
            users = count_unique_users(q, week["start"], week["end"])
        except Exception as e:
            print(f"  ERROR {week['label']}: {e}")
            users = 0
        week_users.append(users)
        print(f"  {week['label']:35s} {users:6} unique users")
        time.sleep(0.5)

    all_release_weekly.append({
        "release":    rel,
        "week_users": week_users,
        "link_query": link_query,
    })


# ── Run categories by release analysis ───────────────────────────────────────
# 90-day window to capture full release lifetime.
# Uses cumulative exclusions matching the monthly analysis priority order.

RELEASE_CAT_WINDOW_START = (TODAY - datetime.timedelta(days=90)).strftime("%Y-%m-%dT00:00:00")
RELEASE_CAT_WINDOW_END   = TODAY.strftime("%Y-%m-%dT00:00:00")
all_release_cat_results = []

for rel in RELEASES:
    v          = rel["version"]
    rel_filter = rel_filter_str(rel)
    dists_str  = ",".join(rel["dists"]) if rel.get("dists") else rel.get("dist")
    print(f"\n── Categories by Release: {v} (dist {dists_str}) ──")

    cumulative_excl = ""
    results = []
    for cat in CATEGORIES:
        link_filter = cat["link_filter"]
        link_query  = f'{BASE_QUERY} {rel_filter} {cumulative_excl} {link_filter}'.strip()
        discover_q  = f'{DISCOVER_BASE_QUERY} {rel_filter} {cumulative_excl} {link_filter}'.strip()
        try:
            users = count_unique_users(discover_q, RELEASE_CAT_WINDOW_START, RELEASE_CAT_WINDOW_END)
        except Exception as e:
            print(f"  ERROR {cat['label']}: {e}")
            users = 0
        print(f"  {cat['label']:30s} {users:6} unique users")
        results.append({**cat, "users": users, "link_query": link_query})
        nx_key = cat["key"] if cat["key"] in NX else None
        if nx_key:
            cumulative_excl = (cumulative_excl + " " + NX[nx_key]).strip()
        time.sleep(0.5)

    total = sum(r["users"] for r in results)
    print(f"  {'TOTAL':30s} {total:6} users")
    all_release_cat_results.append({
        "release": rel,
        "results": results,
        "total":   total,
    })


# ── Overall hang-impacted riders per device class ─────────────────────────────
# Yesterday only; single query per device class — no category breakdown.
# Uses Sentry's native device.class field (high | medium | low).

YESTERDAY           = (TODAY - datetime.timedelta(days=1))
DEVICE_WINDOW_START = YESTERDAY.strftime("%Y-%m-%dT00:00:00")
DEVICE_WINDOW_END   = TODAY.strftime("%Y-%m-%dT00:00:00")
device_class_sentry = {}

print(f"\n── Device Class: overall hang-impacted riders ({YESTERDAY}) ──")
for tier in DEVICE_TIERS:
    discover_q = f'{DISCOVER_BASE_QUERY} device.class:{tier["key"]}'.strip()
    link_q     = f'{BASE_QUERY} device.class:{tier["key"]}'.strip()
    try:
        users = count_unique_users(discover_q, DEVICE_WINDOW_START, DEVICE_WINDOW_END)
        time.sleep(0.3)
    except Exception as e:
        print(f"  ERROR {tier['key']}: {e}")
        users = 0
    device_class_sentry[tier["key"]] = {"users": users, "query": link_q}
    print(f"  {tier['key']:<10} {users:,} riders impacted")

# Fetch total iOS riders per device class from BigQuery (yesterday's partition).
print(f"\n── BigQuery: iOS riders per device class ({YESTERDAY}) ──")
bq_device_totals = fetch_bq_device_class_totals()
print(f"  {'Class':<10} {'Total riders':>14} {'Hang-impacted':>14} {'Impact rate':>12}")
print(f"  {'-'*52}")
for tier in DEVICE_TIERS:
    total    = bq_device_totals[tier["key"]]
    impacted = device_class_sentry[tier["key"]]["users"]
    rate     = f"{impacted / total * 100:.2f}%" if total > 0 else "—"
    print(f"  {tier['key']:<10} {total:>14,} {impacted:>14,} {rate:>12}")


# ── Hang-impacted riders by country for current month ─────────────────────────
CURRENT_MONTH        = MONTHS[-1]
COUNTRY_WINDOW_START = CURRENT_MONTH["start"]
COUNTRY_WINDOW_END   = CURRENT_MONTH["end"]


def fetch_bq_country_totals():
    """iOS rider count per country_code from BigQuery — same window as Sentry hang data."""
    if not BQ_AVAILABLE:
        return {}
    month_start = COUNTRY_WINDOW_START[:10]
    month_end   = COUNTRY_WINDOW_END[:10]
    query = f"""
        SELECT
          country_code,
          COUNT(DISTINCT rider_id) AS rider_count
        FROM `fulfillment-dwh-production.cl.rider_devices` d
        WHERE d.created_date >= "{month_start}"
          AND d.created_date < "{month_end}"
          AND device.operating_system LIKE "%iOS%"
        GROUP BY country_code
        ORDER BY rider_count DESC
    """
    try:
        client = _bq.Client(project="logistics-rider-staging")
        rows = list(client.query(query).result())
        return {row.country_code: row.rider_count for row in rows if row.country_code}
    except Exception as e:
        print(f"WARNING: BigQuery country totals fetch failed: {e}")
        return {}


print(f"\n── BigQuery: iOS riders per country ({CURRENT_MONTH['label']}) ──")
bq_country_totals = fetch_bq_country_totals()
print(f"  {len(bq_country_totals)} countries found in BigQuery")

print(f"\n── Country breakdown: hang-impacted riders ({CURRENT_MONTH['label']}) ──")
print(f"  {'Country':<6} {'Total riders':>14} {'Hang-impacted':>14} {'Impact rate':>12}")
print(f"  {'-'*50}")
country_breakdown = []
for code, total_riders in sorted(bq_country_totals.items(), key=lambda x: -x[1]):
    code_lower = code.lower()
    discover_q = f'{DISCOVER_BASE_QUERY} country:{code_lower}'
    link_q     = f'{BASE_QUERY} country:{code_lower}'
    try:
        users = count_unique_users(discover_q, COUNTRY_WINDOW_START, COUNTRY_WINDOW_END)
        time.sleep(0.3)
    except Exception as e:
        print(f"  ERROR {code_lower}: {e}")
        users = 0
    rate = f"{users / total_riders * 100:.2f}%" if total_riders > 0 else "—"
    print(f"  {code_lower:<6} {total_riders:>14,} {users:>14,} {rate:>12}")
    country_breakdown.append({
        "country_code": code_lower,
        "users":        users,
        "total_riders": total_riders,
        "query":        link_q,
    })

country_breakdown.sort(key=lambda x: -x["users"])


# ── Build JS data ─────────────────────────────────────────────────────────────

months_js = json.dumps([
    {
        "label":  m["month"]["label"],
        "short":  m["month"]["short"],
        "total":  m["total"],
        "start":  m["month"]["start"],
        "end":    m["month"]["end"],
        "categories": [
            {
                "key":      r["key"],
                "label":    r["label"],
                "color":    r["color"],
                "users":    r["users"],
                "query":    r["link_query"],
                "culprits": r["culprits"],
            }
            for r in m["results"]
        ],
    }
    for m in all_month_results
], indent=2)


weeks_js = json.dumps([
    {
        "label": w["label"],
        "short": w["short"],
        "start": w["start"],
        "end":   w["end"],
    }
    for w in WEEKS
], indent=2)


releases_js = json.dumps([
    {
        "label": r["release"]["label"],
        "short": r["release"]["short"],
        "users": r["users"],
        "start": RELEASE_WINDOW_START,
        "end":   RELEASE_WINDOW_END,
        "query": r["link_query"],
    }
    for r in all_release_results
], indent=2)


release_weekly_js = json.dumps([
    {
        "label":      r["release"]["label"],
        "short":      r["release"]["short"],
        "week_users": r["week_users"],
        "link_query": r["link_query"],
    }
    for r in all_release_weekly
], indent=2)


release_categories_js = json.dumps([
    {
        "label":  r["release"]["label"],
        "short":  r["release"]["short"],
        "total":  r["total"],
        "start":  RELEASE_CAT_WINDOW_START,
        "end":    RELEASE_CAT_WINDOW_END,
        "categories": [
            {
                "key":      cat["key"],
                "label":    cat["label"],
                "color":    cat["color"],
                "users":    cat["users"],
                "query":    cat["link_query"],
                "culprits": cat.get("culprits", []),
            }
            for cat in r["results"]
        ],
    }
    for r in all_release_cat_results
], indent=2)


device_class_js = json.dumps([
    {
        "key":            tier["key"],
        "label":          tier["label"],
        "color":          tier["color"],
        "impacted_users": device_class_sentry[tier["key"]]["users"],
        "query":          device_class_sentry[tier["key"]]["query"],
        "total_riders":   bq_device_totals[tier["key"]],
    }
    for tier in DEVICE_TIERS
], indent=2)


country_js = json.dumps({
    "month": CURRENT_MONTH["label"],
    "start": COUNTRY_WINDOW_START,
    "end":   COUNTRY_WINDOW_END,
    "countries": country_breakdown,
}, indent=2)

month_tabs_html = "\n".join(
    f'  <div class="tab" data-panel="month-{i}">{m["short"]}{" (partial)" if i == len(MONTHS) - 1 else ""}</div>'
    for i, m in enumerate(MONTHS)
)
month_panels_html = "\n".join(
    f'<div class="panel" id="panel-month-{i}"></div>'
    for i in range(len(MONTHS))
)
month_summary_text = " · ".join(m["short"] for m in MONTHS)
current_month_note = f'{MONTHS[-1]["short"]} is a partial month (data up to {fetch_date}).'

# ── Generate HTML ─────────────────────────────────────────────────────────────

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Sentry AppHang Analysis — Rider iOS</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
  <style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            background: #f8fafc; color: #111; padding: 32px; }}
    h1 {{ font-size: 22px; font-weight: 700; margin-bottom: 4px; }}
    .subtitle {{ color: #666; font-size: 13px; margin-bottom: 24px; }}

    /* ── Tabs ── */
    .tabs {{ display: flex; gap: 4px; margin-bottom: 0; border-bottom: 2px solid #e5e7eb; }}
    .tab {{ padding: 9px 20px; font-size: 13px; font-weight: 600; cursor: pointer;
             border-radius: 6px 6px 0 0; border: 1px solid transparent;
             border-bottom: none; color: #6b7280; background: transparent;
             position: relative; top: 2px; transition: color .15s; }}
    .tab:hover {{ color: #111; }}
    .tab.active {{ color: #111; background: #fff; border-color: #e5e7eb;
                   border-bottom-color: #fff; }}

    /* ── Panels ── */
    .panel {{ display: none; background: #fff; border: 1px solid #e5e7eb;
               border-top: none; border-radius: 0 6px 6px 6px; padding: 28px; }}
    .panel.active {{ display: block; }}

    /* ── Overview panel ── */
    .overview-totals {{ display: flex; gap: 16px; margin-bottom: 28px; flex-wrap: wrap; }}
    .total-card {{ background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 8px;
                   padding: 14px 20px; min-width: 180px; }}
    .total-card .month-name {{ font-size: 11px; text-transform: uppercase;
                                letter-spacing: .05em; color: #6b7280; margin-bottom: 4px; }}
    .total-card .month-total {{ font-size: 26px; font-weight: 700; color: #111; }}
    .total-card .month-sub {{ font-size: 11px; color: #9ca3af; margin-top: 2px; }}

    .overview-chart-wrap {{ margin-bottom: 28px; }}
    .overview-chart-wrap canvas {{ max-width: 100%; height: 320px !important; }}

    .compare-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
    .compare-table thead th {{ text-align: left; padding: 8px 12px;
      border-bottom: 2px solid #e5e7eb; color: #666; font-weight: 600;
      font-size: 11px; text-transform: uppercase; letter-spacing: .05em; white-space: nowrap; }}
    .compare-table thead th.num {{ text-align: right; }}
    .compare-table tbody tr {{ border-bottom: 1px solid #f3f4f6; }}
    .compare-table tbody tr:hover {{ background: #f9fafb; }}
    .compare-table td {{ padding: 9px 12px; vertical-align: middle; }}
    .compare-table td.num {{ text-align: right; font-variant-numeric: tabular-nums; }}
    .compare-table td.cell-link {{ cursor: pointer; }}
    .compare-table td.cell-link:hover {{ background: #eff6ff; color: #2563eb; text-decoration: underline; }}
    .td-group {{ display: flex; align-items: center; gap: 8px; font-weight: 500; white-space: nowrap; }}
    .delta-pos {{ color: #dc2626; font-size: 11px; }}
    .delta-neg {{ color: #16a34a; font-size: 11px; }}
    .delta-neu {{ color: #9ca3af; font-size: 11px; }}

    /* ── Month detail panel ── */
    .layout {{ display: flex; gap: 40px; align-items: flex-start; flex-wrap: wrap; }}
    .chart-wrap {{ width: 300px; flex-shrink: 0; }}
    .chart-wrap canvas {{ width: 300px !important; height: 300px !important; }}
    .legend {{ margin-top: 14px; }}
    .legend-item {{ display: flex; align-items: center; gap: 8px; font-size: 12px;
                    margin-bottom: 6px; cursor: pointer; }}
    .legend-item:hover .legend-label {{ text-decoration: underline; }}
    .swatch {{ width: 12px; height: 12px; border-radius: 2px; flex-shrink: 0; }}
    .legend-label {{ color: #111; }}
    .legend-pct {{ color: #666; margin-left: auto; font-variant-numeric: tabular-nums; }}
    .table-wrap {{ flex: 1; min-width: 420px; overflow-x: auto; }}
    table.detail {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
    table.detail thead th {{ text-align: left; padding: 8px 12px;
      border-bottom: 2px solid #e5e7eb; color: #666; font-weight: 600;
      font-size: 11px; text-transform: uppercase; letter-spacing: .05em;
      cursor: pointer; white-space: nowrap; user-select: none; }}
    table.detail thead th:hover {{ color: #111; }}
    table.detail thead th.sorted-asc::after {{ content: " ↑"; }}
    table.detail thead th.sorted-desc::after {{ content: " ↓"; }}
    table.detail tbody tr {{ border-bottom: 1px solid #f3f4f6; cursor: pointer; }}
    table.detail tbody tr:hover {{ background: #f9fafb; }}
    table.detail td {{ padding: 10px 12px; vertical-align: top; }}
    .td-pct {{ font-variant-numeric: tabular-nums; color: #374151; }}
    .td-users {{ font-variant-numeric: tabular-nums; font-weight: 600; }}
    .td-culprits {{ color: #6b7280; font-size: 11px;
                    font-family: "SFMono-Regular", Menlo, monospace; line-height: 1.6; }}
    .td-culprits span {{ display: inline-block; background: #f3f4f6; border-radius: 4px;
                          padding: 1px 6px; margin: 1px 2px 1px 0; }}

    .weekly-chart-wrap {{ background: #fff; border: 1px solid #e5e7eb; border-radius: 8px;
                         padding: 18px; margin-bottom: 24px; }}
    .weekly-chart-wrap canvas {{ max-width: 100%; height: 360px !important; }}
    .weekly-summary {{ display: flex; gap: 16px; margin-bottom: 20px; flex-wrap: wrap; }}
    .weekly-card {{ background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 8px;
                    padding: 12px 16px; min-width: 150px; }}
    .weekly-card .week-name {{ font-size: 11px; text-transform: uppercase; letter-spacing: .05em;
                               color: #6b7280; margin-bottom: 4px; }}
    .weekly-card .week-total {{ font-size: 22px; font-weight: 700; }}

    .footer {{ margin-top: 32px; font-size: 11px; color: #9ca3af;
               border-top: 1px solid #e5e7eb; padding-top: 16px; }}
  </style>
</head>
<body>

<h1>AppHang Analysis — Rider iOS</h1>
<p class="subtitle">
  Filter: <code>is:unresolved "*App Hang* detected*"</code> &nbsp;|&nbsp;
  Data fetched: {fetch_date} &nbsp;|&nbsp;
  Months: {month_summary_text}
</p>

<div class="tabs" id="tabs">
  <div class="tab active" data-panel="overview">Overview</div>
{month_tabs_html}
  <div class="tab" data-panel="releases">By Release</div>
  <div class="tab" data-panel="country">By Country</div>
  <div class="tab" data-panel="release-categories">Categories by Release</div>
  <div class="tab" data-panel="device-class">Device Class</div>
</div>

<div class="panel active" id="panel-overview">
  <div class="overview-totals" id="overview-totals"></div>
  <div class="overview-chart-wrap">
    <canvas id="bar-chart"></canvas>
  </div>
  <table class="compare-table" id="compare-table">
    <thead id="compare-thead">
      <tr id="compare-header-row">
        <th>Category</th>
      </tr>
    </thead>
    <tbody id="compare-tbody"></tbody>
  </table>
</div>

{month_panels_html}

<div class="panel" id="panel-release-weekly">
  <p style="font-size:13px;color:#666;margin-bottom:20px;">
    Total AppHang riders impacted per week per release — overall hang numbers, no category breakdown &nbsp;|&nbsp;
    Click any cell to open Sentry for that release + week
  </p>
  <div class="weekly-chart-wrap">
    <canvas id="release-weekly-chart"></canvas>
  </div>
  <table class="compare-table" id="release-weekly-table">
    <thead>
      <tr id="release-weekly-header-row">
        <th>Release</th>
      </tr>
    </thead>
    <tbody id="release-weekly-tbody"></tbody>
  </table>
</div>

<div class="panel" id="panel-releases">
  <p style="font-size:13px;color:#666;margin-bottom:20px;">
    AppHang category breakdown per release — last 90 days &nbsp;|&nbsp;
    Click any cell to open the matching Sentry query for that release
  </p>
  <div class="weekly-summary" id="release-summary"></div>
  <div class="weekly-chart-wrap">
    <canvas id="release-chart"></canvas>
  </div>
  <table class="compare-table" id="release-table">
    <thead>
      <tr id="release-header-row">
        <th>Category</th>
      </tr>
    </thead>
    <tbody id="release-tbody"></tbody>
  </table>
</div>

<div class="panel" id="panel-release-categories">
  <p style="font-size:13px;color:#666;margin-bottom:20px;">
    AppHang categories per release — last 90 days &nbsp;|&nbsp;
    Click any cell to open the matching Sentry query
  </p>
  <div class="weekly-summary" id="release-categories-summary"></div>
  <div class="weekly-chart-wrap">
    <canvas id="release-categories-chart"></canvas>
  </div>
  <table class="compare-table" id="release-categories-table">
    <thead>
      <tr id="release-categories-header-row">
        <th>Category</th>
      </tr>
    </thead>
    <tbody id="release-categories-tbody"></tbody>
  </table>
</div>

<div class="panel" id="panel-device-class">
  <p style="font-size:13px;color:#666;margin-bottom:16px;">
    Overall hang-impacted riders per device class — last 30 days &nbsp;|&nbsp;
    <strong>High-end:</strong> iPhone 13+ (A15+) &nbsp;·&nbsp;
    <strong>Mid-range:</strong> iPhone 11–12 &amp; X/XR/XS (A11–A14) &nbsp;·&nbsp;
    <strong>Low-end:</strong> iPhone 10 &amp; older &nbsp;|&nbsp;
    Riders from BigQuery (yesterday) · Hang-impacted from Sentry device.class (yesterday) · Click count to open Sentry
  </p>

  <div class="weekly-summary" id="device-class-fleet-summary"></div>

  <div class="weekly-chart-wrap">
    <canvas id="device-class-chart"></canvas>
  </div>
  <table class="compare-table" id="device-class-table">
    <thead>
      <tr>
        <th>Device Class</th>
        <th class="num">Avg Daily Riders</th>
        <th class="num">Hang-impacted</th>
        <th class="num">Impact Rate</th>
      </tr>
    </thead>
    <tbody id="device-class-tbody"></tbody>
  </table>
</div>

<div class="panel" id="panel-country">
  <p style="font-size:13px;color:#666;margin-bottom:16px;">
    Hang-impacted riders by <code>country_code</code> — {CURRENT_MONTH['label']} (partial, data up to {fetch_date}) &nbsp;|&nbsp;
    Total riders from BigQuery (<code>rider_devices</code>, iOS, {COUNTRY_WINDOW_START[:10]} – {COUNTRY_WINDOW_END[:10]}) &nbsp;|&nbsp;
    Click any row to open Sentry filtered by that country
  </p>
  <div style="background:#fff;border:1px solid #e5e7eb;border-radius:8px;padding:18px;margin-bottom:24px;overflow:auto;">
    <canvas id="country-chart" style="max-width:100%;height:600px !important;"></canvas>
  </div>
  <table class="compare-table" id="country-table">
    <thead>
      <tr>
        <th>Country Code</th>
        <th class="num">Total iOS Riders</th>
        <th class="num">Hang-impacted</th>
        <th class="num">Impact Rate</th>
      </tr>
    </thead>
    <tbody id="country-tbody"></tbody>
  </table>
</div>

<div class="footer">
  <strong>Notes:</strong>
  <br>• <strong>Counts are deduplicated</strong> — each issue counted in exactly one category (highest priority wins).
     Priority: Mapbox › Naver › WebKit › Firebase › Sentry SDK › Keyboard › Location › Audio › Camera › Storage.
  <br>• <strong>Filter:</strong> <code>"*App Hang* detected*"</code> — matches all AppHang events reported by the Sentry SDK.
  <br>• <strong>Naver</strong> matched via <code>stack.function:*NavigationMapView*</code> (app-level wrapper).
  <br>• <strong>Firebase</strong> is statically linked — matched via <code>message:</code> contains for known culprits (<code>FIRCLS*</code>, <code>FireApp*</code>).
  <br>• <strong>{current_month_note}</strong> Earlier months in this report are complete months.
  <br>• <strong>Device Class tab:</strong> both hang-impacted riders (Sentry) and fleet totals (BigQuery) are for yesterday only — giving a true daily impact rate.
  <br>• <strong>Delta arrows:</strong> ▲ = more riders impacted (worse), ▼ = fewer (better).
  <br>• <strong>Release 4.2625.3 is excluded</strong> — it only rolled out to ~5% of riders and is not comparable to full-rollout releases.
</div>

<script>
const ORG        = "{ORG}";
const PROJECT_ID = "{PROJECT}";
const BASE_URL   = `https://${{ORG}}.sentry.io/issues/`;

const MONTHS              = {months_js};
const WEEKS               = {weeks_js};
const RELEASES            = {releases_js};
const RELEASE_WEEKLY      = {release_weekly_js};
const RELEASE_CATEGORIES  = {release_categories_js};
const DEVICE_CLASS        = {device_class_js};
const COUNTRY_DATA        = {country_js};

function sentryURL(row, month) {{
  const p = new URLSearchParams({{
    environment: "production",
    project: PROJECT_ID,
    query: row.query,
    start: month.start,
    end:   month.end,
  }});
  return BASE_URL + "?" + p.toString();
}}

// ── Tab switching ────────────────────────────────────────────────────────────
document.querySelectorAll(".tab").forEach(tab => {{
  tab.addEventListener("click", () => {{
    document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
    document.querySelectorAll(".panel").forEach(p => p.classList.remove("active"));
    tab.classList.add("active");
    document.getElementById("panel-" + tab.dataset.panel).classList.add("active");
  }});
}});

// ── Overview: total cards ────────────────────────────────────────────────────
const totalsEl = document.getElementById("overview-totals");
MONTHS.forEach(m => {{
  const card = document.createElement("div");
  card.className = "total-card";
  card.innerHTML = `
    <div class="month-name">${{m.label}}</div>
    <div class="month-total">${{m.total.toLocaleString()}}</div>
    <div class="month-sub">unique riders impacted</div>`;
  totalsEl.appendChild(card);
}});

// ── Overview: grouped bar chart ──────────────────────────────────────────────
(function() {{
  const catKeys   = MONTHS[0].categories.map(c => c.key);
  const catLabels = MONTHS[0].categories.map(c => c.label);
  const catColors = MONTHS[0].categories.map(c => c.color);

  const datasets = MONTHS.map((m, mi) => {{
    const opacity = mi === 0 ? "cc" : mi === 1 ? "99" : "55";
    return {{
      label: m.short,
      data: m.categories.map(c => c.users),
      backgroundColor: catColors.map(col => col + (mi === 0 ? "" : mi === 1 ? "bb" : "77")),
      borderColor: catColors.map(col => col),
      borderWidth: 1,
      borderRadius: 3,
    }};
  }});

  const ctx = document.getElementById("bar-chart").getContext("2d");
  new Chart(ctx, {{
    type: "bar",
    data: {{ labels: catLabels, datasets }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{ position: "top", labels: {{ font: {{ size: 12 }} }} }},
        tooltip: {{
          callbacks: {{
            label: ctx => ` ${{ctx.dataset.label}}: ${{ctx.parsed.y.toLocaleString()}} riders`
          }}
        }}
      }},
      scales: {{
        x: {{ ticks: {{ font: {{ size: 11 }} }}, grid: {{ display: false }} }},
        y: {{ ticks: {{ font: {{ size: 11 }} }}, beginAtZero: true,
               title: {{ display: true, text: "Riders impacted", font: {{ size: 11 }} }} }}
      }}
    }}
  }});
}})();

// ── Overview: comparison table (dynamic — works for any number of months) ────
(function() {{
  // Build header row dynamically
  const headerRow = document.getElementById("compare-header-row");
  MONTHS.forEach(m => {{
    const th = document.createElement("th");
    th.className = "num";
    th.textContent = m.short;
    headerRow.appendChild(th);
  }});
  // Delta columns: each consecutive pair
  for (let i = 1; i < MONTHS.length; i++) {{
    const th = document.createElement("th");
    th.className = "num";
    th.textContent = `${{MONTHS[i].short}} vs ${{MONTHS[i-1].short}}`;
    headerRow.appendChild(th);
  }}

  const tbody = document.getElementById("compare-tbody");
  const cats  = MONTHS[0].categories;

  function deltaHTML(a, b) {{
    if (a === 0 && b === 0) return '<span class="delta-neu">—</span>';
    const diff = b - a;
    const pct  = a === 0 ? "∞" : Math.abs(Math.round(diff / a * 100)) + "%";
    if (diff > 0) return `<span class="delta-pos">▲ +${{diff.toLocaleString()}} (${{pct}})</span>`;
    if (diff < 0) return `<span class="delta-neg">▼ ${{diff.toLocaleString()}} (${{pct}})</span>`;
    return '<span class="delta-neu">± 0</span>';
  }}

  cats.forEach((cat, ci) => {{
    const vals = MONTHS.map(m => m.categories[ci].users);
    const tr = document.createElement("tr");

    // Category label cell
    const tdLabel = document.createElement("td");
    tdLabel.innerHTML = `<div class="td-group">
      <span class="td-swatch" style="background:${{cat.color}}"></span>${{cat.label}}
    </div>`;
    tr.appendChild(tdLabel);

    // One clickable cell per month
    MONTHS.forEach((m, mi) => {{
      const td = document.createElement("td");
      td.className = "num cell-link";
      td.title = `Open ${{cat.label}} in Sentry — ${{m.label}}`;
      td.textContent = vals[mi].toLocaleString();
      td.addEventListener("click", () =>
        window.open(sentryURL(m.categories[ci], m), "_blank"));
      tr.appendChild(td);
    }});

    // Delta cells for consecutive month pairs
    for (let i = 1; i < MONTHS.length; i++) {{
      const td = document.createElement("td");
      td.className = "num";
      td.innerHTML = deltaHTML(vals[i-1], vals[i]);
      tr.appendChild(td);
    }}

    tbody.appendChild(tr);
  }});

  // Totals row
  const totals = MONTHS.map(m => m.total);
  const tr = document.createElement("tr");
  tr.style.cssText = "font-weight:700; border-top:2px solid #e5e7eb;";
  const tdTotalLabel = document.createElement("td");
  tdTotalLabel.textContent = "Total";
  tr.appendChild(tdTotalLabel);
  totals.forEach(t => {{
    const td = document.createElement("td");
    td.className = "num";
    td.textContent = t.toLocaleString();
    tr.appendChild(td);
  }});
  for (let i = 1; i < totals.length; i++) {{
    const td = document.createElement("td");
    td.className = "num";
    td.innerHTML = deltaHTML(totals[i-1], totals[i]);
    tr.appendChild(td);
  }}
  tbody.appendChild(tr);
}})();

// ── Release Weekly tab ──────────────────────────────────────────────────────
(function() {{
  if (!RELEASE_WEEKLY.length || !WEEKS.length) return;

  const COLORS = ["#6366f1", "#f59e0b", "#10b981", "#ef4444", "#8b5cf6"];

  function deltaHTML(a, b) {{
    if (a === 0 && b === 0) return '<span class="delta-neu">—</span>';
    const diff = b - a;
    const pct = a === 0 ? "∞" : Math.abs(Math.round(diff / a * 100)) + "%";
    if (diff > 0) return `<span class="delta-pos">▲ +${{diff.toLocaleString()}} (${{pct}})</span>`;
    if (diff < 0) return `<span class="delta-neg">▼ ${{diff.toLocaleString()}} (${{pct}})</span>`;
    return '<span class="delta-neu">± 0</span>';
  }}

  // Table header
  const headerRow = document.getElementById("release-weekly-header-row");
  WEEKS.forEach(w => {{
    const th = document.createElement("th");
    th.className = "num";
    th.textContent = w.label;
    headerRow.appendChild(th);
  }});
  for (let i = 1; i < WEEKS.length; i++) {{
    const th = document.createElement("th");
    th.className = "num";
    th.textContent = `${{WEEKS[i].short}} vs ${{WEEKS[i-1].short}}`;
    headerRow.appendChild(th);
  }}

  // Table body
  const tbody = document.getElementById("release-weekly-tbody");
  RELEASE_WEEKLY.forEach((rel, ri) => {{
    const tr = document.createElement("tr");
    const color = COLORS[ri % COLORS.length];

    const tdLabel = document.createElement("td");
    tdLabel.innerHTML = `<div class="td-group">
      <span style="display:inline-block;width:12px;height:12px;border-radius:50%;
                   background:${{color}};flex-shrink:0"></span>${{rel.label}}
    </div>`;
    tr.appendChild(tdLabel);

    WEEKS.forEach((w, wi) => {{
      const td = document.createElement("td");
      td.className = "num cell-link";
      td.title = `Open ${{rel.label}} in Sentry — ${{w.label}}`;
      td.textContent = rel.week_users[wi].toLocaleString();
      td.addEventListener("click", () =>
        window.open(sentryURL({{ query: rel.link_query }}, w), "_blank"));
      tr.appendChild(td);
    }});

    for (let i = 1; i < WEEKS.length; i++) {{
      const td = document.createElement("td");
      td.className = "num";
      td.innerHTML = deltaHTML(rel.week_users[i-1], rel.week_users[i]);
      tr.appendChild(td);
    }}
    tbody.appendChild(tr);
  }});

  // Total Hangs row
  const weekTotals = WEEKS.map((_, wi) =>
    RELEASE_WEEKLY.reduce((sum, rel) => sum + rel.week_users[wi], 0));
  const totalRow = document.createElement("tr");
  totalRow.style.cssText = "font-weight:700; border-top:2px solid #e5e7eb;";
  const tdTotalLabel = document.createElement("td");
  tdTotalLabel.textContent = "Total Hangs";
  totalRow.appendChild(tdTotalLabel);
  weekTotals.forEach(t => {{
    const td = document.createElement("td");
    td.className = "num";
    td.textContent = t.toLocaleString();
    totalRow.appendChild(td);
  }});
  for (let i = 1; i < WEEKS.length; i++) {{
    const td = document.createElement("td");
    td.className = "num";
    td.innerHTML = deltaHTML(weekTotals[i-1], weekTotals[i]);
    totalRow.appendChild(td);
  }}
  tbody.appendChild(totalRow);

  // Line chart
  const datasets = RELEASE_WEEKLY.map((rel, ri) => ({{
    label:           rel.label,
    data:            rel.week_users,
    borderColor:     COLORS[ri % COLORS.length],
    backgroundColor: COLORS[ri % COLORS.length],
    tension:         0.3,
    fill:            false,
    pointRadius:     5,
    pointHoverRadius: 7,
  }}));

  new Chart(document.getElementById("release-weekly-chart").getContext("2d"), {{
    type: "line",
    data: {{
      labels:   WEEKS.map(w => w.label),
      datasets: datasets,
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{ position: "top", labels: {{ font: {{ size: 12 }} }} }},
        tooltip: {{
          callbacks: {{
            label: ctx => ` ${{ctx.dataset.label}}: ${{ctx.parsed.y.toLocaleString()}} riders`
          }}
        }}
      }},
      scales: {{
        y: {{ beginAtZero: true,
               title: {{ display: true, text: "Riders impacted", font: {{ size: 11 }} }} }}
      }}
    }}
  }});
}})();

// ── By Release tab ──────────────────────────────────────────────────────────
(function() {{
  if (!RELEASES.length) {{
    document.getElementById("panel-releases").innerHTML =
      '<p style="color:#9ca3af;padding:24px;">No releases found in the last 90 days.</p>';
    return;
  }}

  // Summary cards
  const summary = document.getElementById("release-summary");
  RELEASES.forEach(r => {{
    const card = document.createElement("div");
    card.className = "weekly-card";
    card.innerHTML = `
      <div class="week-name">${{r.label}}</div>
      <div class="week-total">${{r.users.toLocaleString()}}</div>
      <div class="month-sub">unique riders impacted</div>`;
    summary.appendChild(card);
  }});

  function deltaHTML(a, b) {{
    if (a === 0 && b === 0) return '<span class="delta-neu">—</span>';
    const diff = b - a;
    const pct = a === 0 ? "∞" : Math.abs(Math.round(diff / a * 100)) + "%";
    if (diff > 0) return `<span class="delta-pos">▲ +${{diff.toLocaleString()}} (${{pct}})</span>`;
    if (diff < 0) return `<span class="delta-neg">▼ ${{diff.toLocaleString()}} (${{pct}})</span>`;
    return '<span class="delta-neu">± 0</span>';
  }}

  // Table header
  const headerRow = document.getElementById("release-header-row");
  RELEASES.forEach(r => {{
    const th = document.createElement("th");
    th.className = "num";
    th.textContent = r.short;
    headerRow.appendChild(th);
  }});
  for (let i = 1; i < RELEASES.length; i++) {{
    const th = document.createElement("th");
    th.className = "num";
    th.textContent = `${{RELEASES[i].short}} vs ${{RELEASES[i-1].short}}`;
    headerRow.appendChild(th);
  }}

  // Table body — single total row
  const tbody = document.getElementById("release-tbody");
  const tr = document.createElement("tr");
  const tdLabel = document.createElement("td");
  tdLabel.innerHTML = '<div class="td-group"><strong>Total Hangs</strong></div>';
  tr.appendChild(tdLabel);
  RELEASES.forEach(r => {{
    const td = document.createElement("td");
    td.className = "num cell-link";
    td.title = `Open AppHangs in Sentry — ${{r.label}}`;
    td.textContent = r.users.toLocaleString();
    td.addEventListener("click", () => window.open(sentryURL(r, r), "_blank"));
    tr.appendChild(td);
  }});
  for (let i = 1; i < RELEASES.length; i++) {{
    const td = document.createElement("td");
    td.className = "num";
    td.innerHTML = deltaHTML(RELEASES[i-1].users, RELEASES[i].users);
    tr.appendChild(td);
  }}
  tbody.appendChild(tr);

  // Simple bar chart
  const HANG_COLOR = "#BAE1FF";
  new Chart(document.getElementById("release-chart").getContext("2d"), {{
    type: "bar",
    data: {{
      labels:   RELEASES.map(r => r.label),
      datasets: [{{
        label:           "AppHangs",
        data:            RELEASES.map(r => r.users),
        backgroundColor: RELEASES.map((_, i) =>
          i === RELEASES.length - 1 ? HANG_COLOR : HANG_COLOR + "99"),
        borderColor:     HANG_COLOR,
        borderWidth:     1,
        borderRadius:    4,
      }}]
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{ callbacks: {{
          label: ctx => ` ${{ctx.parsed.y.toLocaleString()}} riders impacted`
        }} }}
      }},
      scales: {{
        x: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 11 }} }} }},
        y: {{ beginAtZero: true,
               title: {{ display: true, text: "Riders impacted", font: {{ size: 11 }} }} }}
      }}
    }}
  }});
}})();

// ── Categories by Release tab ─────────────────────────────────────────────────
(function() {{
  if (!RELEASE_CATEGORIES.length) {{
    document.getElementById("panel-release-categories").innerHTML =
      '<p style="color:#9ca3af;padding:24px;">No data available — re-run the report generator.</p>';
    return;
  }}

  const summary = document.getElementById("release-categories-summary");
  RELEASE_CATEGORIES.forEach(r => {{
    const card = document.createElement("div");
    card.className = "weekly-card";
    card.innerHTML = `
      <div class="week-name">${{r.label}}</div>
      <div class="week-total">${{r.total.toLocaleString()}}</div>
      <div class="month-sub">unique riders impacted</div>`;
    summary.appendChild(card);
  }});

  const catColors = RELEASE_CATEGORIES[0].categories.map(c => c.color);
  const catLabels = RELEASE_CATEGORIES[0].categories.map(c => c.label);

  const datasets = RELEASE_CATEGORIES.map((r, ri) => ({{
    label: r.short,
    data: r.categories.map(c => c.users),
    backgroundColor: catColors.map(col => col + (ri === RELEASE_CATEGORIES.length - 1 ? "" : "99")),
    borderColor: catColors.map(col => col),
    borderWidth: 1,
    borderRadius: 3,
  }}));

  new Chart(document.getElementById("release-categories-chart").getContext("2d"), {{
    type: "bar",
    data: {{ labels: catLabels, datasets }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{ position: "top", labels: {{ font: {{ size: 12 }} }} }},
        tooltip: {{
          callbacks: {{
            label: ctx => ` ${{ctx.dataset.label}}: ${{ctx.parsed.y.toLocaleString()}} riders`
          }}
        }}
      }},
      scales: {{
        x: {{ ticks: {{ font: {{ size: 11 }} }}, grid: {{ display: false }} }},
        y: {{ ticks: {{ font: {{ size: 11 }} }}, beginAtZero: true,
               title: {{ display: true, text: "Riders impacted", font: {{ size: 11 }} }} }}
      }}
    }}
  }});

  function deltaHTML(a, b) {{
    if (a === 0 && b === 0) return '<span class="delta-neu">—</span>';
    const diff = b - a;
    const pct  = a === 0 ? "∞" : Math.abs(Math.round(diff / a * 100)) + "%";
    if (diff > 0) return `<span class="delta-pos">▲ +${{diff.toLocaleString()}} (${{pct}})</span>`;
    if (diff < 0) return `<span class="delta-neg">▼ ${{diff.toLocaleString()}} (${{pct}})</span>`;
    return '<span class="delta-neu">± 0</span>';
  }}

  const headerRow = document.getElementById("release-categories-header-row");
  RELEASE_CATEGORIES.forEach(r => {{
    const th = document.createElement("th");
    th.className = "num";
    th.textContent = r.short;
    headerRow.appendChild(th);
  }});
  for (let i = 1; i < RELEASE_CATEGORIES.length; i++) {{
    const th = document.createElement("th");
    th.className = "num";
    th.textContent = `${{RELEASE_CATEGORIES[i].short}} vs ${{RELEASE_CATEGORIES[i-1].short}}`;
    headerRow.appendChild(th);
  }}

  const tbody = document.getElementById("release-categories-tbody");
  const cats = RELEASE_CATEGORIES[0].categories;

  cats.forEach((cat, ci) => {{
    const vals = RELEASE_CATEGORIES.map(r => r.categories[ci].users);
    const tr = document.createElement("tr");

    const tdLabel = document.createElement("td");
    tdLabel.innerHTML = `<div class="td-group">
      <span class="td-swatch" style="background:${{cat.color}}"></span>${{cat.label}}
    </div>`;
    tr.appendChild(tdLabel);

    RELEASE_CATEGORIES.forEach((r, ri) => {{
      const td = document.createElement("td");
      td.className = "num cell-link";
      td.title = `Open ${{cat.label}} in Sentry — ${{r.label}}`;
      td.textContent = vals[ri].toLocaleString();
      td.addEventListener("click", () =>
        window.open(sentryURL(r.categories[ci], r), "_blank"));
      tr.appendChild(td);
    }});

    for (let i = 1; i < RELEASE_CATEGORIES.length; i++) {{
      const td = document.createElement("td");
      td.className = "num";
      td.innerHTML = deltaHTML(vals[i-1], vals[i]);
      tr.appendChild(td);
    }}

    tbody.appendChild(tr);
  }});

  const totals = RELEASE_CATEGORIES.map(r => r.total);
  const tr = document.createElement("tr");
  tr.style.cssText = "font-weight:700; border-top:2px solid #e5e7eb;";
  const tdTotalLabel = document.createElement("td");
  tdTotalLabel.textContent = "Total";
  tr.appendChild(tdTotalLabel);
  totals.forEach(t => {{
    const td = document.createElement("td");
    td.className = "num";
    td.textContent = t.toLocaleString();
    tr.appendChild(td);
  }});
  for (let i = 1; i < totals.length; i++) {{
    const td = document.createElement("td");
    td.className = "num";
    td.innerHTML = deltaHTML(totals[i-1], totals[i]);
    tr.appendChild(td);
  }}
  tbody.appendChild(tr);
}})();

// ── Month detail panels ──────────────────────────────────────────────────────
MONTHS.forEach((month, mi) => {{
  const panel = document.getElementById(`panel-month-${{mi}}`);
  const canvasId  = `pie-${{mi}}`;
  const legendId  = `legend-${{mi}}`;
  const tbodyId   = `tbody-${{mi}}`;

  panel.innerHTML = `
    <p style="font-size:13px;color:#666;margin-bottom:20px;">
      ${{month.label}} &nbsp;|&nbsp;
      Total unique riders impacted: <strong>${{month.total.toLocaleString()}}</strong> &nbsp;|&nbsp;
      Click any row or legend item to open in Sentry
    </p>
    <div class="layout">
      <div class="chart-wrap">
        <canvas id="${{canvasId}}"></canvas>
        <div class="legend" id="${{legendId}}"></div>
      </div>
      <div class="table-wrap">
        <table class="detail" id="tbl-${{mi}}">
          <thead><tr>
            <th data-col="label">Group</th>
            <th data-col="pct" class="sorted-desc">% Active</th>
            <th data-col="users">Riders</th>
            <th data-col="culprits">Representative Culprits</th>
          </tr></thead>
          <tbody id="${{tbodyId}}"></tbody>
        </table>
      </div>
    </div>`;

  const total = month.total;
  function pct(u) {{ return total ? (u / total * 100).toFixed(1) + "%" : "0.0%"; }}

  let sortCol = "pct", sortDir = -1;

  function sorted(data) {{
    return [...data].sort((a, b) => {{
      if (sortCol === "label") return sortDir * a.label.localeCompare(b.label);
      if (sortCol === "culprits") return sortDir * (a.culprits.length - b.culprits.length);
      return sortDir * (a.users - b.users);
    }});
  }}

  function renderLegend(data) {{
    const leg = document.getElementById(legendId);
    if (!leg) return;
    leg.innerHTML = "";
    sorted(data).filter(r => r.users > 0).forEach(row => {{
      const item = document.createElement("div");
      item.className = "legend-item";
      item.innerHTML = `<span class="swatch" style="background:${{row.color}}"></span>
        <span class="legend-label">${{row.label}}</span>
        <span class="legend-pct">${{pct(row.users)}}</span>`;
      item.addEventListener("click", () => window.open(sentryURL(row, month), "_blank"));
      leg.appendChild(item);
    }});
  }}

  function renderTable(data) {{
    const tbody = document.getElementById(tbodyId);
    if (!tbody) return;
    tbody.innerHTML = "";
    sorted(data).forEach(row => {{
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><div class="td-group">
          <span class="td-swatch" style="background:${{row.color}}"></span>${{row.label}}
        </div></td>
        <td class="td-pct">${{pct(row.users)}}</td>
        <td class="td-users">${{row.users.toLocaleString()}}</td>
        <td class="td-culprits">
          ${{row.culprits.length
            ? row.culprits.map(c => `<span>${{c}}</span>`).join("")
            : '<span style="color:#d1d5db">—</span>'}}
        </td>`;
      tr.addEventListener("click", () => window.open(sentryURL(row, month), "_blank"));
      tbody.appendChild(tr);
    }});
  }}

  // Sortable headers — must wait until panel is in DOM
  requestAnimationFrame(() => {{
    document.querySelectorAll(`#tbl-${{mi}} thead th[data-col]`).forEach(th => {{
      th.addEventListener("click", () => {{
        const col = th.dataset.col;
        if (sortCol === col) {{ sortDir *= -1; }} else {{ sortCol = col; sortDir = col === "label" ? 1 : -1; }}
        document.querySelectorAll(`#tbl-${{mi}} thead th`).forEach(h =>
          h.classList.remove("sorted-asc", "sorted-desc"));
        th.classList.add(sortDir === 1 ? "sorted-asc" : "sorted-desc");
        renderTable(month.categories);
        renderLegend(month.categories);
      }});
    }});
  }});

  // Pie chart
  const visible = month.categories.filter(r => r.users > 0);
  requestAnimationFrame(() => {{
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    new Chart(canvas.getContext("2d"), {{
      type: "doughnut",
      data: {{
        labels: visible.map(r => r.label),
        datasets: [{{
          data: visible.map(r => r.users),
          backgroundColor: visible.map(r => r.color),
          borderWidth: 2, borderColor: "#fff"
        }}]
      }},
      options: {{
        responsive: false,
        plugins: {{
          legend: {{ display: false }},
          tooltip: {{ callbacks: {{
            label: c => ` ${{c.label}}: ${{c.parsed.toLocaleString()}} riders (${{(c.parsed/total*100).toFixed(1)}}%)`
          }} }}
        }},
        onClick: (evt, elements) => {{
          if (elements.length) window.open(sentryURL(visible[elements[0].index], month), "_blank");
        }}
      }}
    }});
    renderLegend(month.categories);
    renderTable(month.categories);
  }});
}});

// ── Device Class Distribution tab ────────────────────────────────────────────
(function() {{
  if (!DEVICE_CLASS.length) {{
    document.getElementById("panel-device-class").innerHTML =
      '<p style="color:#9ca3af;padding:24px;">No device class data available — re-run the report generator.</p>';
    return;
  }}

  function openSentry(tier) {{
    const p = new URLSearchParams({{
      environment: "production",
      project: PROJECT_ID,
      query: tier.query,
      start: "{DEVICE_WINDOW_START}",
      end:   "{DEVICE_WINDOW_END}",
    }});
    window.open(BASE_URL + "?" + p.toString(), "_blank");
  }}

  // ── Fleet summary cards ────────────────────────────────────────────────────
  const fleetSummary = document.getElementById("device-class-fleet-summary");
  DEVICE_CLASS.forEach(tier => {{
    const total    = tier.total_riders;
    const impacted = tier.impacted_users;
    const rate     = total > 0 ? (impacted / total * 100).toFixed(2) + "%" : "—";
    const card = document.createElement("div");
    card.className = "weekly-card";
    card.innerHTML = `
      <div class="week-name" style="color:${{tier.color}};font-weight:700">${{tier.label}}</div>
      <div style="font-size:11px;color:#6b7280;margin-top:6px;line-height:1.8">
        <span style="display:block">Avg daily riders: <strong>${{total > 0 ? total.toLocaleString() : "—"}}</strong></span>
        <span style="display:block">Hang-impacted: <strong>${{impacted.toLocaleString()}}</strong></span>
        <span style="display:block;margin-top:4px;font-size:14px;font-weight:700;color:#111">
          ${{rate}} hang rate
        </span>
      </div>`;
    fleetSummary.appendChild(card);
  }});

  // ── Simple 3-bar chart ─────────────────────────────────────────────────────
  new Chart(document.getElementById("device-class-chart").getContext("2d"), {{
    type: "bar",
    data: {{
      labels: DEVICE_CLASS.map(t => t.label),
      datasets: [{{
        label: "Hang-impacted riders",
        data: DEVICE_CLASS.map(t => t.impacted_users),
        backgroundColor: DEVICE_CLASS.map(t => t.color),
        borderColor: DEVICE_CLASS.map(t => t.color),
        borderWidth: 1,
        borderRadius: 4,
      }}],
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{
          callbacks: {{
            label: ctx => {{
              const tier = DEVICE_CLASS[ctx.dataIndex];
              const rate = tier.total_riders > 0
                ? ` (${{(ctx.parsed.y / tier.total_riders * 100).toFixed(2)}}% of fleet)`
                : "";
              return ` ${{ctx.parsed.y.toLocaleString()}} riders impacted${{rate}}`;
            }}
          }}
        }}
      }},
      scales: {{
        x: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 12 }} }} }},
        y: {{ beginAtZero: true,
               title: {{ display: true, text: "Riders impacted", font: {{ size: 11 }} }} }}
      }},
      onClick: (evt, elements) => {{
        if (elements.length) openSentry(DEVICE_CLASS[elements[0].index]);
      }}
    }}
  }});

  // ── 3-row table: Device Class | Total Riders | Hang-impacted | Impact Rate ─
  const tbody = document.getElementById("device-class-tbody");
  DEVICE_CLASS.forEach(tier => {{
    const total    = tier.total_riders;
    const impacted = tier.impacted_users;
    const rate     = total > 0 ? (impacted / total * 100).toFixed(2) + "%" : "—";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><div class="td-group">
        <span class="swatch" style="background:${{tier.color}}"></span>
        <span style="font-weight:600;color:${{tier.color}}">${{tier.label}}</span>
      </div></td>
      <td class="num">${{total > 0 ? total.toLocaleString() : "—"}}</td>
      <td class="num cell-link" title="Open ${{tier.label}} hangs in Sentry">
        ${{impacted.toLocaleString()}}
      </td>
      <td class="num">
        <span style="background:#f3f4f6;border-radius:4px;padding:2px 8px;font-weight:600">
          ${{rate}}
        </span>
      </td>`;
    tr.querySelectorAll("td")[2].addEventListener("click", () => openSentry(tier));
    tbody.appendChild(tr);
  }});

  const grandTotal    = DEVICE_CLASS.reduce((s, t) => s + t.total_riders,   0);
  const grandImpacted = DEVICE_CLASS.reduce((s, t) => s + t.impacted_users, 0);
  const grandRate     = grandTotal > 0 ? (grandImpacted / grandTotal * 100).toFixed(2) + "%" : "—";
  const trTotal = document.createElement("tr");
  trTotal.style.cssText = "font-weight:700; border-top:2px solid #e5e7eb;";
  trTotal.innerHTML = `
    <td>Total</td>
    <td class="num">${{grandTotal > 0 ? grandTotal.toLocaleString() : "—"}}</td>
    <td class="num">${{grandImpacted.toLocaleString()}}</td>
    <td class="num">
      <span style="background:#f3f4f6;border-radius:4px;padding:2px 8px;font-weight:600">
        ${{grandRate}}
      </span>
    </td>`;
  tbody.appendChild(trTotal);
}})();

// ── By Country tab ───────────────────────────────────────────────────────────
(function() {{
  if (!COUNTRY_DATA.countries || !COUNTRY_DATA.countries.length) {{
    document.getElementById("panel-country").innerHTML =
      '<p style="color:#9ca3af;padding:24px;">No country data available — re-run the report generator.</p>';
    return;
  }}

  const countries    = [...COUNTRY_DATA.countries].sort((a, b) => {{
    const rateA = a.total_riders > 0 ? a.users / a.total_riders : 0;
    const rateB = b.total_riders > 0 ? b.users / b.total_riders : 0;
    return rateB - rateA;
  }});
  const totalImpact  = countries.reduce((s, c) => s + c.users, 0);
  const totalRiders  = countries.reduce((s, c) => s + c.total_riders, 0);
  const top          = countries.slice(0, 30);

  const PALETTE = [
    "#6366f1","#f59e0b","#10b981","#ef4444","#8b5cf6",
    "#0ea5e9","#f97316","#14b8a6","#ec4899","#64748b",
    "#84cc16","#a78bfa","#fb923c","#34d399","#f472b6",
    "#38bdf8","#fbbf24","#4ade80","#f87171","#a3e635",
    "#818cf8","#fb7185","#2dd4bf","#facc15","#c084fc",
    "#60a5fa","#fca5a1","#6ee7b7","#fde68a","#93c5fd",
  ];

  function impactRate(impacted, total) {{
    return total > 0 ? (impacted / total * 100).toFixed(2) + "%" : "—";
  }}

  new Chart(document.getElementById("country-chart").getContext("2d"), {{
    type: "bar",
    data: {{
      labels: top.map(c => c.country_code),
      datasets: [{{
        label: "Hang impact rate (%)",
        data:  top.map(c => c.total_riders > 0 ? parseFloat((c.users / c.total_riders * 100).toFixed(2)) : 0),
        backgroundColor: top.map((_, i) => PALETTE[i % PALETTE.length]),
        borderRadius: 3,
        borderWidth: 0,
      }}],
    }},
    options: {{
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{
          callbacks: {{
            label: ctx => {{
              const c = top[ctx.dataIndex];
              const rate = impactRate(c.users, c.total_riders);
              return ` ${{rate}} (${{c.users.toLocaleString()}} of ${{c.total_riders.toLocaleString()}} riders)`;
            }}
          }}
        }}
      }},
      scales: {{
        x: {{ beginAtZero: true,
               title: {{ display: true, text: "Hang impact rate (%)", font: {{ size: 11 }} }},
               ticks: {{ font: {{ size: 11 }} }} }},
        y: {{ ticks: {{ font: {{ size: 11 }} }}, grid: {{ display: false }} }}
      }},
      onClick: (evt, elements) => {{
        if (!elements.length) return;
        const c = top[elements[0].index];
        const p = new URLSearchParams({{
          environment: "production",
          project: PROJECT_ID,
          query: c.query,
          start: COUNTRY_DATA.start,
          end:   COUNTRY_DATA.end,
        }});
        window.open(BASE_URL + "?" + p.toString(), "_blank");
      }}
    }}
  }});

  const tbody = document.getElementById("country-tbody");
  countries.forEach(c => {{
    const rate = impactRate(c.users, c.total_riders);
    const tr   = document.createElement("tr");
    tr.style.cursor = "pointer";
    tr.title = `Open ${{c.country_code}} hangs in Sentry`;
    tr.innerHTML = `
      <td style="font-weight:600">${{c.country_code}}</td>
      <td class="num">${{c.total_riders > 0 ? c.total_riders.toLocaleString() : "—"}}</td>
      <td class="num cell-link">${{c.users.toLocaleString()}}</td>
      <td class="num">
        <span style="background:#f3f4f6;border-radius:4px;padding:2px 8px;font-weight:600">
          ${{rate}}
        </span>
      </td>`;
    tr.addEventListener("click", () => {{
      const p = new URLSearchParams({{
        environment: "production",
        project: PROJECT_ID,
        query: c.query,
        start: COUNTRY_DATA.start,
        end:   COUNTRY_DATA.end,
      }});
      window.open(BASE_URL + "?" + p.toString(), "_blank");
    }});
    tbody.appendChild(tr);
  }});

  const grandRate = impactRate(totalImpact, totalRiders);
  const trTotal   = document.createElement("tr");
  trTotal.style.cssText = "font-weight:700; border-top:2px solid #e5e7eb;";
  trTotal.innerHTML = `
    <td>Total</td>
    <td class="num">${{totalRiders > 0 ? totalRiders.toLocaleString() : "—"}}</td>
    <td class="num">${{totalImpact.toLocaleString()}}</td>
    <td class="num">
      <span style="background:#f3f4f6;border-radius:4px;padding:2px 8px;font-weight:600">
        ${{grandRate}}
      </span>
    </td>`;
  tbody.appendChild(trTotal);
}})();
</script>
</body>
</html>"""

os.makedirs(OUT_DIR, exist_ok=True)
out_path = os.path.join(OUT_DIR, "index.html")
with open(out_path, "w") as f:
    f.write(html)

print(f"\nReport written to: {out_path}")
print(f"Deploy: drop the '{OUT_DIR}/' folder on https://netlify.com/drop")
