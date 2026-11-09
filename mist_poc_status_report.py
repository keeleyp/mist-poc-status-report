#!/usr/bin/env python3
"""Mist POC status report - is the POC being used?

For every API key in mist_poc_status_report.ini (one or more per Mist cloud),
finds every org the key can see and collects:

  * org name and number of sites
  * APs, switches and gateways - online / offline
  * unique wireless and wired clients seen in the last 24 hours
  * audit log entries (and distinct admins) in the last 24 hours

The results are printed as a table and saved to an Excel workbook. The
workbook keeps a History sheet with one row per org per day (re-running on
the same day replaces that day's figures), a Latest sheet, and one Trend
sheet per figure (orgs down the side, dates across the top, with a chart).

If any cloud or org fails, the script stops, lists the failures and writes
nothing, so the workbook never holds a partial day.
"""
import configparser
import os
import sys
import time
from datetime import date, datetime
from urllib.parse import urlparse

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INI_NAME = "mist_poc_status_report.ini"
config = configparser.ConfigParser()
config.optionxform = str  # keep cloud names exactly as written
if not config.read(os.path.join(SCRIPT_DIR, INI_NAME)):
    sys.exit(f"{INI_NAME} not found next to the script - copy {INI_NAME}.example and fill it in.")

OUTPUT_FILE = os.path.join(SCRIPT_DIR, os.path.expanduser(
    config.get("settings", "output_file", fallback="mist_poc_status.xlsx")))
EXCLUDE_ORGS = {x.strip().lower() for x in
                config.get("settings", "exclude_orgs", fallback="").split(",") if x.strip()}

REQUEST_TIMEOUT = 60
PAGE_LIMIT = 1000
WINDOW_SECONDS = 24 * 3600
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
ACTIVE_FILL = PatternFill("solid", fgColor="C6EFCE")
IDLE_FILL = PatternFill("solid", fgColor="FFC7CE")

# (header, key) for every column in the History / Latest sheets.
COLUMNS = [
    ("Date", "date"),
    ("Cloud", "cloud"),
    ("Org Name", "org_name"),
    ("Org ID", "org_id"),
    ("Usage", "usage"),
    ("Sites", "sites"),
    ("APs Total", "ap_total"),
    ("APs Online", "ap_online"),
    ("APs Offline", "ap_offline"),
    ("Switches Total", "switch_total"),
    ("Switches Online", "switch_online"),
    ("Switches Offline", "switch_offline"),
    ("Gateways Total", "gateway_total"),
    ("Gateways Online", "gateway_online"),
    ("Gateways Offline", "gateway_offline"),
    ("Wireless Clients 24h", "wireless_clients"),
    ("Wired Clients 24h", "wired_clients"),
    ("Audit Log Entries 24h", "audit_entries"),
    ("Admins Active 24h", "audit_admins"),
]
HEADERS = [h for h, _ in COLUMNS]
KEYS = [k for _, k in COLUMNS]

# (Trend sheet name, key) - one Trend sheet per figure.
TRENDS = [
    ("Trend Sites", "sites"),
    ("Trend APs Online", "ap_online"),
    ("Trend APs Offline", "ap_offline"),
    ("Trend Switches Online", "switch_online"),
    ("Trend Switches Offline", "switch_offline"),
    ("Trend Gateways Online", "gateway_online"),
    ("Trend Gateways Offline", "gateway_offline"),
    ("Trend Wireless Clients", "wireless_clients"),
    ("Trend Wired Clients", "wired_clients"),
    ("Trend Audit Entries", "audit_entries"),
]


class ApiError(Exception):
    pass


# --------------------------------------------------------------------------
# HTTP
# --------------------------------------------------------------------------
def build_session():
    """Shared HTTP session, wired up for a TLS-inspecting proxy (e.g. Zscaler)
    when the optional [network] section is present in the .ini. With no
    [network] section, behaviour is unchanged (certifi CA store, and the
    HTTP_PROXY / HTTPS_PROXY / NO_PROXY environment variables if set)."""
    session = requests.Session()
    if "network" not in config:
        return session

    section = config["network"]
    ca_bundle = section.get("ca_bundle", "").strip()
    if ca_bundle:
        ca_path = os.path.expanduser(ca_bundle)
        if not os.path.isfile(ca_path):
            sys.exit(f"[network] ca_bundle does not exist: {ca_path}\n"
                     f"This should be a PEM file containing your proxy's root CA "
                     f"certificate (e.g. exported from Zscaler).")
        session.verify = ca_path
    elif not section.getboolean("verify_ssl", fallback=True):
        session.verify = False
        print("WARNING: TLS certificate verification is DISABLED ([network] verify_ssl = false). "
              "Only use this as a last resort.")
        requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)

    for scheme in ("http", "https"):
        proxy = section.get(f"{scheme}_proxy", "").strip()
        if proxy:
            session.proxies[scheme] = proxy
    return session


SESSION = build_session()


def cloud_to_host(cloud):
    """Normalise an .ini cloud value to an API host name like api.eu.mist.com."""
    value = cloud.lower().strip()
    if "://" in value:
        value = urlparse(value).netloc
    value = value.strip("/")
    if value.endswith("mist.com"):
        if value.startswith("manage."):
            value = "api." + value[len("manage."):]
        return value
    if value in ("us", "global", "global01"):
        return "api.mist.com"
    return f"api.{value}.mist.com"


def api_get(host, token, path, params=None):
    """GET https://{host}/api/v1{path}. Returns the Response, or raises ApiError."""
    url = f"https://{host}/api/v1{path}"
    try:
        resp = SESSION.get(url, headers={"Authorization": f"Token {token}"},
                           params=params, timeout=REQUEST_TIMEOUT)
    except requests.exceptions.SSLError as e:
        raise ApiError(f"TLS error ({e}) - if you're behind Zscaler or a similar proxy, "
                       f"set [network] ca_bundle in {INI_NAME}")
    except requests.exceptions.ProxyError as e:
        raise ApiError(f"proxy error ({e}) - check [network] http_proxy / https_proxy in {INI_NAME}")
    except requests.exceptions.RequestException as e:
        raise ApiError(f"could not connect ({e.__class__.__name__})")

    if resp.status_code == 401:
        raise ApiError("HTTP 401 - token not valid on this cloud")
    if resp.status_code == 403:
        raise ApiError(f"HTTP 403 - token has no access to {path}")
    if resp.status_code == 429:
        raise ApiError("HTTP 429 - API rate limit reached (5000 calls/hour), try again later")
    if not resp.ok:
        raise ApiError(f"HTTP {resp.status_code} on {path}: {resp.text[:200]}")
    return resp


def api_get_all(host, token, path, params=None):
    """Page through a list endpoint (limit/page) and return every item."""
    items, page = [], 1
    while True:
        p = dict(params or {}, limit=PAGE_LIMIT, page=page)
        resp = api_get(host, token, path, p)
        batch = resp.json()
        if not isinstance(batch, list):
            raise ApiError(f"unexpected response from {path} (not a list)")
        items.extend(batch)
        total = resp.headers.get("X-Page-Total")
        if len(batch) < PAGE_LIMIT or (total is not None and len(items) >= int(total)):
            return items
        page += 1


def count_distinct(host, token, path, distinct, start, end):
    """Call a /count endpoint. Returns the list of {distinct_value, count} rows."""
    resp = api_get(host, token, path, {"distinct": distinct, "start": start, "end": end,
                                       "limit": PAGE_LIMIT})
    data = resp.json()
    results = data.get("results") or []
    if len(results) >= PAGE_LIMIT and data.get("total", 0) > len(results):
        # More distinct values than one response holds - fall back to 'total'.
        return results, data["total"]
    return results, len(results)


# --------------------------------------------------------------------------
# Config + org discovery
# --------------------------------------------------------------------------
def load_keys():
    """Return [(api_host, token)] from [clouds]. A cloud can hold several
    comma-separated tokens (e.g. one per POC customer on the same cloud)."""
    if "clouds" not in config or not config["clouds"]:
        sys.exit(f"No clouds listed - add 'cloud = api_token' lines under [clouds] in {INI_NAME}.")
    keys = []
    for cloud, value in config["clouds"].items():
        tokens = [t.strip() for t in value.split(",") if t.strip()]
        if not tokens or any(t.startswith("YOUR_") for t in tokens):
            sys.exit(f"No API token set for '{cloud}' in {INI_NAME}.")
        host = cloud_to_host(cloud)
        for token in tokens:
            if (host, token) in keys:
                sys.exit(f"The same token is listed twice for {host} in {INI_NAME}.")
            keys.append((host, token))
    return keys


def key_label(host, token, self_name=None):
    who = f" ({self_name})" if self_name else ""
    return f"{host} key ...{token[-4:]}{who}"


def discover_orgs(host, token, self_data):
    """Return [(org_id, org_name)] the key can see, from /self privileges.
    MSP-level privileges are expanded to every org in that MSP. Site-level
    privileges are skipped (org-wide figures need org access)."""
    orgs, skipped_sites = {}, set()
    for priv in self_data.get("privileges", []):
        scope = priv.get("scope")
        if scope == "org" and priv.get("org_id"):
            orgs[priv["org_id"]] = priv.get("name") or priv.get("org_name") or ""
        elif scope == "msp" and priv.get("msp_id"):
            for org in api_get_all(host, token, f"/msps/{priv['msp_id']}/orgs"):
                orgs.setdefault(org["id"], org.get("name", ""))
        elif scope == "site" and priv.get("org_id"):
            skipped_sites.add(priv["org_id"])
    for org_id in skipped_sites - set(orgs):
        print(f"  note: {key_label(host, token)} only has site-level access to org {org_id} - skipped")
    for org_id, name in orgs.items():
        if not name:
            orgs[org_id] = api_get(host, token, f"/orgs/{org_id}").json().get("name", org_id)
    return sorted(orgs.items(), key=lambda o: o[1].lower())


# --------------------------------------------------------------------------
# Per-org figures
# --------------------------------------------------------------------------
def collect_org(host, token, org_id, org_name, start, end, today):
    row = {"date": today, "cloud": host, "org_name": org_name, "org_id": org_id}

    row["sites"] = len(api_get_all(host, token, f"/orgs/{org_id}/sites"))

    for dev_type in ("ap", "switch", "gateway"):
        row[f"{dev_type}_total"] = row[f"{dev_type}_online"] = row[f"{dev_type}_offline"] = 0
    for dev in api_get_all(host, token, f"/orgs/{org_id}/stats/devices", {"type": "all"}):
        dev_type = dev.get("type")
        if dev_type not in ("ap", "switch", "gateway"):
            continue
        row[f"{dev_type}_total"] += 1
        state = "online" if dev.get("status") == "connected" else "offline"
        row[f"{dev_type}_{state}"] += 1

    _, row["wireless_clients"] = count_distinct(host, token, f"/orgs/{org_id}/clients/count",
                                                "mac", start, end)
    _, row["wired_clients"] = count_distinct(host, token, f"/orgs/{org_id}/wired_clients/count",
                                             "mac", start, end)
    admins, n_admins = count_distinct(host, token, f"/orgs/{org_id}/logs/count",
                                      "admin_name", start, end)
    row["audit_entries"] = sum(a.get("count", 0) for a in admins)
    row["audit_admins"] = n_admins

    used = row["wireless_clients"] or row["wired_clients"] or row["audit_entries"]
    row["usage"] = "Active" if used else "Idle"
    return row


# --------------------------------------------------------------------------
# Output
# --------------------------------------------------------------------------
def print_table(rows):
    cols = [("Cloud", lambda r: r["cloud"]),
            ("Org", lambda r: r["org_name"]),
            ("Usage", lambda r: r["usage"]),
            ("Sites", lambda r: r["sites"]),
            ("APs on/off", lambda r: f"{r['ap_online']}/{r['ap_offline']}"),
            ("Switches on/off", lambda r: f"{r['switch_online']}/{r['switch_offline']}"),
            ("GWs on/off", lambda r: f"{r['gateway_online']}/{r['gateway_offline']}"),
            ("WiFi clients", lambda r: r["wireless_clients"]),
            ("Wired clients", lambda r: r["wired_clients"]),
            ("Audit entries", lambda r: r["audit_entries"])]
    cells = [[str(f(r)) for _, f in cols] for r in rows]
    widths = [max(len(h), *(len(c[i]) for c in cells)) if cells else len(h)
              for i, (h, _) in enumerate(cols)]
    left = {0, 1, 2}

    def fmt(values):
        return "  ".join(v.ljust(w) if i in left else v.rjust(w)
                         for i, (v, w) in enumerate(zip(values, widths)))

    print()
    print(fmt([h for h, _ in cols]))
    print("  ".join("-" * w for w in widths))
    for c in cells:
        print(fmt(c))
    active = sum(r["usage"] == "Active" for r in rows)
    print(f"\n{len(rows)} org(s): {active} active, {len(rows) - active} idle "
          f"(active = any client or audit log entry in the last 24 hours)")


def style_header_row(ws, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(1, c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.freeze_panes = "A2"


def write_rows_sheet(wb, title, rows, index):
    if title in wb.sheetnames:
        del wb[title]
    ws = wb.create_sheet(title, index)
    ws.append(HEADERS)
    for r in rows:
        ws.append([r.get(k) for k in KEYS])
    style_header_row(ws, len(HEADERS))
    widths = {"Date": 12, "Cloud": 18, "Org Name": 32, "Org ID": 38, "Usage": 9}
    for i, h in enumerate(HEADERS, 1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(h, 12)
    usage_col = KEYS.index("usage") + 1
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(r, usage_col)
        cell.fill = ACTIVE_FILL if cell.value == "Active" else IDLE_FILL
    ws.auto_filter.ref = ws.dimensions
    return ws


def write_trend_sheet(wb, title, key, history):
    if title in wb.sheetnames:
        del wb[title]
    ws = wb.create_sheet(title)
    dates = sorted({r["date"] for r in history})
    orgs = {}  # (cloud, org_id) -> latest name
    for r in sorted(history, key=lambda r: r["date"]):
        orgs[(r["cloud"], r["org_id"])] = r["org_name"]
    org_keys = sorted(orgs, key=lambda k: (k[0], orgs[k].lower()))
    values = {(r["cloud"], r["org_id"], r["date"]): r.get(key) for r in history}

    ws.append(["Cloud", "Org Name"] + dates)
    for ok in org_keys:
        ws.append([ok[0], orgs[ok]] + [values.get((ok[0], ok[1], d)) for d in dates])
    style_header_row(ws, len(dates) + 2)
    ws.freeze_panes = "C2"
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 32
    for i in range(3, len(dates) + 3):
        ws.column_dimensions[get_column_letter(i)].width = 12

    if dates and org_keys:
        chart = LineChart()
        chart.title = title.replace("Trend ", "")
        chart.y_axis.title = chart.title
        chart.x_axis.title = "Date"
        chart.height, chart.width = 10, 24
        # One series per org: each row is a series, dates (header) are categories.
        data = Reference(ws, min_col=2, max_col=len(dates) + 2, min_row=2, max_row=len(org_keys) + 1)
        chart.add_data(data, from_rows=True, titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=3, max_col=len(dates) + 2, min_row=1))
        ws.add_chart(chart, f"A{len(org_keys) + 4}")


def read_history(wb):
    if "History" not in wb.sheetnames:
        return []
    ws = wb["History"]
    header = [c.value for c in ws[1]]
    rows = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        if not any(values):
            continue
        rec = dict(zip(header, values))
        rows.append({k: rec.get(h) for h, k in COLUMNS})
    return rows


def save_workbook(new_rows, today):
    if os.path.exists(OUTPUT_FILE):
        try:
            wb = load_workbook(OUTPUT_FILE)
        except Exception as e:
            sys.exit(f"Could not open {OUTPUT_FILE} ({e}) - is it open in Excel?")
        print(f"\nUpdating {OUTPUT_FILE}")
    else:
        wb = Workbook()
        wb.remove(wb.active)
        print(f"\nCreating {OUTPUT_FILE}")

    # One row per org per day: drop today's earlier figures for these orgs.
    replaced = {(r["cloud"], r["org_id"]) for r in new_rows}
    history = [r for r in read_history(wb)
               if not (r["date"] == today and (r["cloud"], r["org_id"]) in replaced)]
    history.extend(new_rows)
    history.sort(key=lambda r: (r["date"], r["cloud"], (r["org_name"] or "").lower()))

    latest = sorted(new_rows, key=lambda r: (r["cloud"], r["org_name"].lower()))
    write_rows_sheet(wb, "Latest", latest, 0)
    write_rows_sheet(wb, "History", history, 1)
    for title, key in TRENDS:
        write_trend_sheet(wb, title, key, history)

    try:
        wb.save(OUTPUT_FILE)
    except PermissionError:
        sys.exit(f"Could not save {OUTPUT_FILE} - close it in Excel and run again.")
    days = len({r["date"] for r in history})
    print(f"Saved: {len(latest)} org(s) today, {days} day(s) of history.")


def fail(title, failures):
    print(f"\n{title} - stopping, nothing written.\n")
    width = max(len(w) for w, _ in failures)
    for where, reason in failures:
        print(f"  {where.ljust(width)}  {reason}")
    sys.exit(1)


# --------------------------------------------------------------------------
def main():
    keys = load_keys()
    now = int(time.time())
    start, end = now - WINDOW_SECONDS, now
    today = date.today().isoformat()
    print(f"Mist POC status report - {datetime.now():%Y-%m-%d %H:%M} "
          f"(last 24h window), {len(keys)} API key(s)")

    # 1. Check every key first, so a bad key stops the run before any work.
    failures, checked = [], []
    for host, token in keys:
        try:
            self_data = api_get(host, token, "/self").json()
            checked.append((host, token, self_data))
        except ApiError as e:
            failures.append((key_label(host, token), str(e)))
    if failures:
        fail("Some API keys failed", failures)

    # 2. Find orgs for every key.
    work, seen = [], set()
    for host, token, self_data in checked:
        label = key_label(host, token, self_data.get("name") or self_data.get("email"))
        try:
            orgs = discover_orgs(host, token, self_data)
        except ApiError as e:
            failures.append((label, f"listing orgs: {e}"))
            continue
        print(f"  {label}: {len(orgs)} org(s)")
        for org_id, org_name in orgs:
            if org_id.lower() in EXCLUDE_ORGS or org_name.lower() in EXCLUDE_ORGS:
                print(f"    excluded: {org_name}")
                continue
            if (host, org_id) in seen:
                continue  # same org visible to two keys on one cloud
            seen.add((host, org_id))
            work.append((host, token, org_id, org_name))
    if failures:
        fail("Org discovery failed", failures)
    if not work:
        sys.exit("No orgs found for any API key.")

    # 3. Collect figures per org.
    rows = []
    for i, (host, token, org_id, org_name) in enumerate(work, 1):
        print(f"  [{i}/{len(work)}] {host}  {org_name}")
        try:
            rows.append(collect_org(host, token, org_id, org_name, start, end, today))
        except ApiError as e:
            failures.append((f"{host} / {org_name}", str(e)))
    if failures:
        fail("Some orgs failed", failures)

    print_table(sorted(rows, key=lambda r: (r["cloud"], r["org_name"].lower())))
    save_workbook(rows, today)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nStopped - nothing written.")
