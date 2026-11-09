# Mist POC Status Report

A quick check of whether Juniper Mist POC orgs are actually being used, across several clouds and API keys.

For every API key in the ini, the script finds every org the key can see and reports:

- Org name and number of sites
- APs, switches and gateways: total, online and offline
- Unique wireless and wired clients (distinct MACs) in the last 24 hours. Wired clients include any Mist devices that the switches see as clients.
- Audit log entries, and distinct admins, in the last 24 hours
- **Usage**: *Active* if there were any clients or audit log entries in the last 24 hours, otherwise *Idle*

The results are printed as a table and saved to an Excel workbook. Every run is time-stamped, so the workbook builds up a trend over time.

After the report you can record each POC's details: Salesforce record, Sales email, SE email, and as many time-stamped notes as you like.

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

### POC details and notes

When the table has been shown, the script lists the orgs with the status of their details:

```
POC details - enter an org # to view or edit, or press Enter to finish.
    1  Dragon School                             complete; 3 note(s)
    2  Highgate School (UK)                      missing: SE Email; 0 note(s)
```

Pick an org number to see its details and all of its notes.

- If any fields are missing, the script asks for them (press Enter to leave one blank). It then asks whether you want to edit the fields that are already filled in.
- If the details are complete, you can edit them. Press Enter to keep a value, or type `-` to clear it.
- Email fields must look like an email address.
- Then add any number of notes, one per line, and press Enter on an empty line to finish. Each note is time-stamped and kept. Earlier notes are never overwritten.

Everything is saved as soon as you finish with an org. Press Enter at the `Org #` prompt to exit.

For scheduled or unattended runs, use `--no-edit` to skip the editor. The script also skips it when it isn't run from a terminal.

```bash
python3 mist_poc_status_report.py --no-edit
```

## Excel workbook (`mist_poc_status.xlsx`)

| Sheet | Contents |
|---|---|
| **Latest** | This run's figures, one row per org, followed by the Salesforce record, Sales email, SE email and all notes |
| **History** | Every run's figures, one row per org per run, with a **Date Time** column (`YYYY-MM-DD HH:MM`). A re-run within the same minute replaces that run's figures. |
| **Trend …** | One sheet per figure (sites, online/offline devices, clients and audit entries). Orgs run down the side and run times across the top. |
| **POC Details** | Salesforce record, Sales email and SE email per org (keyed by cloud and org ID) |
| **POC Notes** | Every note: org, date and time, and note text |

Close the workbook in Excel before running the script. If it's still open when the script saves, the editor asks you to close it and press Enter to retry. A `--no-edit` run stops instead.

## Failures

The script checks every API key first. If any key, org discovery or org collection fails (for example 401, 403 or 429), the script lists the failures and exits with code 1, **writing nothing**. The workbook never holds a partial run.

## Zscaler / proxies

If you're behind a TLS-inspecting proxy, set `[network] ca_bundle` to a PEM file that holds the proxy's root CA. `http_proxy` and `https_proxy` can also be set there. If you leave the section out, the script uses the normal defaults (certifi, plus the `HTTP(S)_PROXY` environment variables). TLS errors point you to `ca_bundle`.
