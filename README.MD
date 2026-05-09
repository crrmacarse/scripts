# Scripts, Apps, and Tools

Making my life easier, one script at a time.

Available as a web hub at **[tools.crrm.dev](https://tools.crrm.dev)**.

## tools.crrm.dev (web app)

A FastAPI app that wraps these scripts for browser use, deployed on Vercel.

### Local development

```sh
pip3 install -r requirements.txt
cp .env.example .env  # fill in values
uvicorn app.main:app --reload
```

`app/main.py` calls `load_dotenv()` at import time, so values in `.env` are picked up automatically in dev. In production, Vercel injects env vars itself.

Then open http://127.0.0.1:8000.

### Environment variables

| Variable | Required | Notes |
| --- | --- | --- |
| `SESSION_SECRET` | yes | Long random string. Signs session cookies. |
| `TOOLS_PASSWORD` | for gated tools | Shared password protecting the CC analyzer. Other tools stay public. |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | one of these | Raw JSON of the Google service account. |
| `GOOGLE_SERVICE_ACCOUNT_JSON_B64` | one of these | Base64-encoded JSON. Best for Vercel — paste as a single line. |
| `GOOGLE_SERVICE_ACCOUNT_FILE` | one of these | Absolute path to the JSON file (handy for dev). |

If none of those are set, the app falls back to `./google-service-account.json` at the repo root (this is the easiest local-dev setup — the file is gitignored). Resolution order is: raw JSON → base64 → explicit file path → repo-root fallback.

To produce the base64 form for Vercel: `base64 -i google-service-account.json | pbcopy`.

### Vercel deployment

1. `vercel link` and `vercel deploy` from the project root.
2. Add the env vars above in the Vercel project's Settings → Environment Variables.
3. Map the `tools.crrm.dev` domain to the project (Settings → Domains).

`vercel.json` rewrites everything to `api/index.py`, which mounts the FastAPI app from `app/main.py`.

### Adding new tools

1. Add a service module under `app/services/`.
2. Add a router under `app/routes/` and include it in `app/main.py`.
3. Register the tool in `app/tools_registry.py` so it shows on the home page.
4. Add `dependencies=[Depends(require_auth)]` to the router if it should sit behind the password gate.

### CC Analyzer Plus — workflow reference

The result page also exposes this as a "How this works" modal in the top-right.

**The flow.** Upload PDF + cutoff + MM source → server parses PDF, fetches MM for the billing cycle, reconciles → review flagged rows → fix MM on phone → Re-analyze → Confirm & Push when reconciled.

**Status labels.** A CC transaction can land in one of three states:
- `matched` — MM entry within ±3 days at the same amount.
- `inaccurate` — MM entry within ±5 days at ≤5% amount drift.
- `missing` — no MM entry close enough.

**Resolve options** (per non-matched row, in fall-through order):
1. Algorithm-suggested candidate checkboxes (inaccurate only).
2. Manual-pick dropdown — every MM entry in the cycle.
3. "To be added manually" — fall-through; you'll create a new MM entry.

Within a row, picking any option clears the others. Picking an MM entry in row A locks it from being picked in row B and strikes it through in the bottom MM table.

**MM lookup window.** Anchored to the billing cycle: `cutoff − 1 month + 1 day … cutoff`, padded ±3 days. The `Period` column is shown verbatim; matching internally compares against `Period − 1 day` to absorb a known one-day drift.

**Punch list actions** (live in "What to do in Money Manager"):
- `Add` — create a new MM entry from the CC tx.
- `Update` — change an existing MM entry's amount to match CC.
- `Match` — MM entry already equals CC; no change needed.
- `Remove` — appears once every CC row is resolved; lists MM entries that didn't link to anything.

**Verification panel** projects MM count + total *after* applying the punch list. Diff card turns green when projected MM count = CC count and projected MM total = CC total (within 1 cent).

**Re-analyze vs Confirm & Push.** Re-analyze (Google Sheet source only) re-fetches MM with the cached PDF — no re-upload. Confirm & Push writes a new worksheet to the destination Google Sheet; don't push until the Verification panel is green.

**Due date.** Cutoff + 21 days, bumped past weekends. Display only — does not affect the lookup window. Holidays are not modelled.

## Setting up

### Utilizing venv

This allows a python project to run separately from root

```sh
# initialize venv
python3 -m venv venv 

# activate venv
source venv/bin/activate
```

### Install Dependencies

```sh
pip3 install -r requirements.txt
```

## Running the scripts via GUI

Start as an app instead of typing in the cli!

```sh
python3 start.py
```

## Analyze Money Manager

A set of reports to further visualize your Money Manager data. 

> Money Manager is a financial app available
on App store and Play store.

```sh
# pass an exported xlsx file from money manager app
python3 analyze_money_manager.py
```

## CC Analyzer

Analyzes Credit Card bill sent by your banks in your email. Currently supports Security Bank.

### Prerequisites

- Enable Google Sheets and Google Drive API in Google Cloud Console
- Share targeted Google sheet file to Service account email as "Editor"
- Add Service Account's JSON file as `google-service-account.json` in the root
- Download PDF file of your CC bill

### Running the app

```sh
python3 cc-analyzer.py --sheet-name "NAME_OF_GOOGLE_SHEET" --billing-period "MMMM YYYYY" --pdf-path "pdf/path" --pdf-password "pdf_password" 
```

> Supported Credit Card: Security Bank

## Set Slack status

Updates Slack status remotely. This could be use in conjunction with CRON to automate "Lunch Break"

### Prerequisites

- Secure Slack OATH token in https://api.slack.com/apps
- Add User Token Scopes of `users.profile:write`
- Install the App
- Copy the OAuth token to `SLACK_OATH_TOKEN` inside `.env` file

### Running the App

```sh
# or leave template blank to input status manually
python3 slack_status.py --template "TEMPLATE"     
```

### Adding to CRON

```sh
# Run every 12nn at Monday to Friday the slack status for Lunch
0 12 * * 1-5 /usr/bin/python3 /path/to/slack_status.py --template "Lunch"
```

## Get Coordinates

Get lat and lng of given location

```sh
python3 get-coordinates.py
```

## Is Weather Good Here?

Checks weather in a given location

```sh
python3 is-weather-good-here.py
```

## Web Scraper

Scrapes a website

```sh
python3 scraper.py
```

##### Get in touch with me via: [Twitter](https://twitter.com/crrmacarse) / [LinkedIn](https://www.linkedin.com/in/christian-ryan-r-macarse-692974166/) / crrmacarse@gmail.com