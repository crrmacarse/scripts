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