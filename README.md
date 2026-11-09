# Mist POC Status Report

A quick check of whether Juniper Mist POC orgs are actually being used, across several clouds and API keys.

For every API key in the ini, the script finds every org the key can see and reports:

- Org name and number of sites
- APs, switches and gateways: total, online and offline
- Unique wireless and wired clients (distinct MACs) in the last 24 hours. Wired clients include any Mist devices that the switches see as clients.
- Audit log entries, and distinct admins, in the last 24 hours
- **Usage**: *Active* if there were any clients or audit log entries in the last 24 hours, otherwise *Idle*

The results are printed as a table and saved to an Excel workbook. Every run is time-stamped, so the workbook builds up a trend over time.

## Setup

```bash
pip install -r requirements.txt
cp mist_poc_status_report.ini.example mist_poc_status_report.ini
```

Then edit `mist_poc_status_report.ini`:

```ini
[clouds]
api.mist.com    = TOKEN
eu              = TOKEN
gc3             = TOKEN_POC_A, TOKEN_POC_B   # several keys on one cloud
```

- A cloud can be written as `api.eu.mist.com`, in short form (`eu`, `gc3`, `us`), or as the portal URL.
- A key reports every org in its `/self` privileges. MSP-level access covers every org in that MSP. Orgs where the key only has site-level access are skipped, because the figures need org-wide access.
- An org that two keys on the same cloud can both see is reported once.
- `[settings] exclude_orgs` takes a comma-separated list of org names or IDs to leave out.

## Run

```bash
python3 mist_poc_status_report.py
```

Each org costs about 6 API calls, which is well within the 5,000 calls/hour limit.

## Excel workbook (`mist_poc_status.xlsx`)

| Sheet | Contents |
|---|---|
| **Latest** | This run's figures, one row per org |
| **History** | Every run's figures, one row per org per run, with a **Date Time** column (`YYYY-MM-DD HH:MM`). A re-run within the same minute replaces that run's figures. |
| **Trend …** | One sheet per figure (sites, online/offline devices, clients and audit entries). Orgs run down the side and run times across the top, with a line chart below. |

Close the workbook in Excel before running the script. Otherwise the script can't save it.

## Failures

The script checks every API key first. If any key, org discovery or org collection fails (for example 401, 403 or 429), the script lists the failures and exits with code 1, **writing nothing**. The workbook never holds a partial run.

## Zscaler / proxies

If you're behind a TLS-inspecting proxy, set `[network] ca_bundle` to a PEM file that holds the proxy's root CA. `http_proxy` and `https_proxy` can also be set there. If you leave the section out, the script uses the normal defaults (certifi, plus the `HTTP(S)_PROXY` environment variables). TLS errors point you to `ca_bundle`.
