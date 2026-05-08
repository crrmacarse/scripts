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