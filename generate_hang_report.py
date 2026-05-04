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

# ── Config ────────────────────────────────────────────────────────────────────

TOKEN   = os.environ.get("SENTRY_AUTH_TOKEN", "")
ORG     = "delivery-hero-pm"
PROJECT = "4506937839976448"
OUT_DIR = "hang_report"

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
    """
    Dynamic month window:
    - After the 20th: current month + previous 2 months
    - On/before the 20th: current month + previous 3 months
    """
    month_count = 3 if TODAY.day > 20 else 4
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

BASE_QUERY = 'is:unresolved error.mechanism:AppHang'

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
        "filters":  ["stack.function:*NavigationMapView*"],
        "excl":     excl("mapbox"),
        "link_filter": "stack.function:*NavigationMapView*",
        "culprits": [
            "NavigationMapView.subscribeToNavigatonUpdates",
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
        "filters":  ["stack.package:*CoreLocation*", "stack.function:*CLLocation*"],
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
            "stack.function:*CameraProvider*",
        ],
        "excl":     excl("mapbox", "naver", "webkit", "firebase", "sentry", "keyboard"),
        "link_filter": "stack.package:*ARKit*",
        "culprits": [
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
                        next_url = seg[1:-1]
                        next_cursor = urllib.parse.parse_qs(
                            urllib.parse.urlparse(next_url).query
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


# ── Run weekly analysis for weekly trend tab ──────────────────────────────────

all_week_results = []

for week in WEEKS:
    print(f"\n── {week['label']} ({week['start'][:10]} → {week['end'][:10]}) ──")

    assigned = set()
    results = []

    for cat in CATEGORIES:
        cat_issues = {}
        for filt in cat["filters"]:
            q = f"{BASE_QUERY} {cat['excl']} {filt}".strip()
            try:
                fetched = fetch_all(q, week["start"], week["end"])
                for iid, u in fetched.items():
                    if iid not in cat_issues or cat_issues[iid] < u:
                        cat_issues[iid] = u
                time.sleep(0.5)
            except Exception as e:
                print(f"  ERROR {cat['label']} / {filt}: {e}")
                time.sleep(3)

        new_issues = {iid: u for iid, u in cat_issues.items() if iid not in assigned}
        cat_users = sum(new_issues.values())
        assigned.update(new_issues.keys())

        results.append({**cat, "users": cat_users})
        print(f"  {cat['label']:30s} {cat_users:6} users ({len(new_issues)} issues)")

    total_unique = sum(r["users"] for r in results)
    print(f"  {'TOTAL':30s} {total_unique:6} users")

    cumulative_excl = ""
    for r in results:
        r["link_query"] = f'{BASE_QUERY} {cumulative_excl} {r["link_filter"]}'.strip()
        nx_key = r["key"] if r["key"] in NX else None
        if nx_key:
            cumulative_excl = (cumulative_excl + " " + NX[nx_key]).strip()

    all_week_results.append({
        "week": week,
        "results": results,
        "total": total_unique,
    })

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
        "label":  w["week"]["label"],
        "short":  w["week"]["short"],
        "total":  w["total"],
        "start":  w["week"]["start"],
        "end":    w["week"]["end"],
        "categories": [
            {
                "key":      r["key"],
                "label":    r["label"],
                "color":    r["color"],
                "users":    r["users"],
                "query":    r["link_query"],
            }
            for r in w["results"]
        ],
    }
    for w in all_week_results
], indent=2)



month_tabs_html = "\n".join(
    f'  <div class="tab" data-panel="month-{i}">{m["short"]}</div>'
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
  Filter: <code>is:unresolved error.mechanism:AppHang</code> &nbsp;|&nbsp;
  Data fetched: {fetch_date} &nbsp;|&nbsp;
  Months: {month_summary_text}
</p>

<div class="tabs" id="tabs">
  <div class="tab active" data-panel="overview">Overview</div>
{month_tabs_html}
  <div class="tab" data-panel="weekly">Weekly Trend</div>
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

<div class="panel" id="panel-weekly">
  <p style="font-size:13px;color:#666;margin-bottom:20px;">
    Last 4 Monday-Sunday weekly windows, including current partial week when applicable &nbsp;|&nbsp;
    Click any table cell to open the matching Sentry query for that week
  </p>
  <div class="weekly-summary" id="weekly-summary"></div>
  <div class="weekly-chart-wrap">
    <canvas id="weekly-chart"></canvas>
  </div>
  <table class="compare-table" id="weekly-table">
    <thead>
      <tr id="weekly-header-row">
        <th>Category</th>
      </tr>
    </thead>
    <tbody id="weekly-tbody"></tbody>
  </table>
</div>

<div class="footer">
  <strong>Notes:</strong>
  <br>• <strong>Counts are deduplicated</strong> — each issue counted in exactly one category (highest priority wins).
     Priority: Mapbox › Naver › WebKit › Firebase › Sentry SDK › Keyboard › Location › Audio › Camera › Storage.
  <br>• <strong>Filter:</strong> <code>error.mechanism:AppHang</code> — matches all AppHang events reported by the Sentry SDK.
  <br>• <strong>Naver</strong> matched via <code>stack.function:*NavigationMapView*</code> (app-level wrapper).
  <br>• <strong>Firebase</strong> is statically linked — matched via <code>message:</code> contains for known culprits (<code>FIRCLS*</code>, <code>FireApp*</code>).
  <br>• <strong>{current_month_note}</strong> Earlier months in this report are complete months.
  <br>• <strong>Delta arrows:</strong> ▲ = more riders impacted (worse), ▼ = fewer (better).
</div>

<script>
const ORG        = "{ORG}";
const PROJECT_ID = "{PROJECT}";
const BASE_URL   = `https://${{ORG}}.sentry.io/issues/`;

const MONTHS = {months_js};
const WEEKS = {weeks_js};

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

// ── Weekly trend tab ────────────────────────────────────────────────────────
(function() {{
  if (!WEEKS.length) return;

  const summary = document.getElementById("weekly-summary");
  WEEKS.forEach(w => {{
    const card = document.createElement("div");
    card.className = "weekly-card";
    card.innerHTML = `
      <div class="week-name">${{w.label}}</div>
      <div class="week-total">${{w.total.toLocaleString()}}</div>
      <div class="month-sub">unique riders impacted</div>`;
    summary.appendChild(card);
  }});

  const headerRow = document.getElementById("weekly-header-row");
  WEEKS.forEach(w => {{
    const th = document.createElement("th");
    th.className = "num";
    th.textContent = w.short;
    headerRow.appendChild(th);
  }});
  for (let i = 1; i < WEEKS.length; i++) {{
    const th = document.createElement("th");
    th.className = "num";
    th.textContent = `${{WEEKS[i].short}} vs ${{WEEKS[i-1].short}}`;
    headerRow.appendChild(th);
  }}

  function deltaHTML(a, b) {{
    if (a === 0 && b === 0) return '<span class="delta-neu">—</span>';
    const diff = b - a;
    const pct = a === 0 ? "∞" : Math.abs(Math.round(diff / a * 100)) + "%";
    if (diff > 0) return `<span class="delta-pos">▲ +${{diff.toLocaleString()}} (${{pct}})</span>`;
    if (diff < 0) return `<span class="delta-neg">▼ ${{diff.toLocaleString()}} (${{pct}})</span>`;
    return '<span class="delta-neu">± 0</span>';
  }}

  const tbody = document.getElementById("weekly-tbody");
  const cats = WEEKS[0].categories;

  cats.forEach((cat, ci) => {{
    const vals = WEEKS.map(w => w.categories[ci].users);
    const tr = document.createElement("tr");

    const tdLabel = document.createElement("td");
    tdLabel.innerHTML = `<div class="td-group">
      <span class="td-swatch" style="background:${{cat.color}}"></span>${{cat.label}}
    </div>`;
    tr.appendChild(tdLabel);

    WEEKS.forEach((w, wi) => {{
      const td = document.createElement("td");
      td.className = "num cell-link";
      td.title = `Open ${{cat.label}} in Sentry — ${{w.label}}`;
      td.textContent = vals[wi].toLocaleString();
      td.addEventListener("click", () =>
        window.open(sentryURL(w.categories[ci], w), "_blank"));
      tr.appendChild(td);
    }});

    for (let i = 1; i < WEEKS.length; i++) {{
      const td = document.createElement("td");
      td.className = "num";
      td.innerHTML = deltaHTML(vals[i-1], vals[i]);
      tr.appendChild(td);
    }}

    tbody.appendChild(tr);
  }});

  const totals = WEEKS.map(w => w.total);
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

  const ctx = document.getElementById("weekly-chart").getContext("2d");
  new Chart(ctx, {{
    type: "line",
    data: {{
      labels: WEEKS.map(w => w.label),
      datasets: WEEKS[0].categories.map((cat, ci) => ({{
        label: cat.label,
        data: WEEKS.map(w => w.categories[ci].users),
        borderColor: cat.color,
        backgroundColor: cat.color,
        tension: 0.3,
        fill: false,
      }}))
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{ position: "right", labels: {{ font: {{ size: 12 }} }} }},
        tooltip: {{
          callbacks: {{
            label: ctx => ` ${{ctx.dataset.label}}: ${{ctx.parsed.y.toLocaleString()}} riders`
          }}
        }}
      }},
      scales: {{
        y: {{ beginAtZero: true, title: {{ display: true, text: "Riders impacted", font: {{ size: 11 }} }} }}
      }}
    }}
  }});
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
</script>
</body>
</html>"""

os.makedirs(OUT_DIR, exist_ok=True)
out_path = os.path.join(OUT_DIR, "index.html")
with open(out_path, "w") as f:
    f.write(html)

print(f"\nReport written to: {out_path}")
print(f"Deploy: drop the '{OUT_DIR}/' folder on https://netlify.com/drop")
