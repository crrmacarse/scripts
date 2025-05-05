# Scripts, Apps, and Tools

Making my life easier, one script at a time.

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

```sh
python3 start.py
```

## Analyze Money Manager

```sh
# pass an exported xlsx file from money manager app
python3 analyze_money_manager.py
```

## CC Analyzer

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

```sh
python3 get-coordinates.py
```

## Is Weather Good Here?

```sh
python3 is-weather-good-here.py
```

## Web Scraper

```sh
python3 scraper.py
```

##### Get in touch with me via: [Twitter](https://twitter.com/crrmacarse) / [LinkedIn](https://www.linkedin.com/in/christian-ryan-r-macarse-692974166/) / crrmacarse@gmail.com