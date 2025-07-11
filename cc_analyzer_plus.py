import gspread
from oauth2client.service_account import ServiceAccountCredentials
import argparse
from PyPDF2 import PdfReader
import re
from gspread_formatting import format_cell_range, CellFormat, TextFormat
from datetime import datetime

def is_match_within_days_and_amount(tran_date, amount, mm_data, days=3):
    try:
        pdf_date = datetime.strptime(tran_date, "%m/%d/%y")
    except ValueError:
        pdf_date = datetime.strptime(tran_date, "%m/%d/%Y")
    for mm_tuple in mm_data:
        mm_date_str, mm_amount, *rest = mm_tuple if isinstance(mm_tuple, tuple) else (mm_tuple, None, None)
        try:
            mm_date = datetime.strptime(mm_date_str, "%m/%d/%y")
        except ValueError:
            mm_date = datetime.strptime(mm_date_str, "%m/%d/%Y")
        if abs((pdf_date - mm_date).days) <= days and amount == mm_amount:
            return mm_tuple
    return None

def extract_mentions(description):
    return re.findall(r'@\w+', description or "")

def remove_mentions(description):
    return re.sub(r'@\w+', '', description or "")

def extract_money_manager_data():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    client = authorize_google_sheets(scope)
    sheet = client.open("Money Manager Data")
    worksheet = sheet.worksheet("Sheet1")
    data = worksheet.get_all_values()

    header = data[0]
    try:
        period_idx = header.index("Period")
        accounts_idx = header.index("Accounts")
        amount_idx = header.index("Amount")
        type_idx = header.index("Income/Expense")
        description_idx = header.index("Description")
    except ValueError as e:
        raise ValueError(f"Missing column in Money Manager data: {e}")

    # TODO: Should be dependent from the period of processed pdf
    # Retrieve 2 months ago data
    today = datetime.today()
    if today.month > 2:
        cutoff = today.replace(month=today.month-2)
    else:
        # If it's Jan or Feb, subtract from previous year
        cutoff = today.replace(year=today.year-1, month=12 + today.month-2)

    mm_data = set(
        (
            # TODO: Bug. MM Data seems to add +1 date on extracted data
            datetime.strptime(row[period_idx], "%m/%d/%Y").strftime("%m/%d/%y"), 
            f"{float(row[amount_idx].replace(',', '')):,.2f}",
            row[description_idx]
        )
        for row in data[1:]
        if len(row) > max(period_idx, amount_idx, accounts_idx)
        and row[accounts_idx] == "Security Bank World"
        and row[type_idx] == "Exp."
        and datetime.strptime(row[period_idx], "%m/%d/%Y") >= cutoff
    )
    return mm_data

def extract_pdf_data(pdf_path, pdf_password):
    # read PDF file
    reader = PdfReader(pdf_path)

    # decrypt PDF password
    if reader.is_encrypted:
        reader.decrypt(pdf_password)

    pdf_data = []
    total_amount = 0
    total_cr = 0

    mm_data = extract_money_manager_data()
    matched_mm_set = set()

    # extract row data from PDF
    for page in reader.pages:
        text = page.extract_text()

        table_pattern = re.compile(r"(\d{2}/\d{2}/\d{2})\s+(\d{2}/\d{2}/\d{2})\s+(.+?)\s+([\d,]+\.\d{2})(?:\s+(CR))?")
        matches = table_pattern.findall(text)

        for match in matches:
            tran_date, post_date, description, amount, is_cr = match

            # skips payment transactions
            if not is_cr:
                total_amount += float(amount.replace(",", ""))
                note = ""
                mm_description = ""

                matched_mm = is_match_within_days_and_amount(tran_date, amount, mm_data, days=3)

                if not matched_mm:
                    # Add missing note if data is not found in Money Manager
                    note = "[!] "

                if matched_mm:
                    matched_mm_set.add(matched_mm)
                    note += remove_mentions(matched_mm[2])
                    mm_description = ", ".join(extract_mentions(matched_mm[2]))

                pdf_data.append([tran_date, post_date, description, amount, note, mm_description, "", ""])
            elif not description.startswith("PAYMENT"):
                total_cr += float(amount.replace(",", ""))
    
    # unused_mm_data = mm_data - matched_mm_set
    # print(unused_mm_data)

    return pdf_data, total_amount, total_cr

def authorize_google_sheets(scope):
    creds = ServiceAccountCredentials.from_json_keyfile_name("google-service-account.json", scope)
    client = gspread.authorize(creds)
    return client

def update_google_sheet(sheet_name, billing_period, pdf_data, total_amount, total_cr):
    # Google Sheets API setup
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    client = authorize_google_sheets(scope)

    sheet = client.open(sheet_name)

    # check if a worksheet with the same name already exists
    existing_titles = [ws.title for ws in sheet.worksheets()]
    if billing_period in existing_titles:
        raise ValueError(f"The following billing period already exists('{billing_period}')!")

    # length of data + 10 extra row
    row_length = str(len(pdf_data) + 10)

    # create new worksheet
    new_worksheet = sheet.add_worksheet(title=billing_period, rows=row_length, cols="20")

    # TODO: doesn't work
    # reorder worksheets to put the new worksheet at the left
    worksheets = sheet.worksheets()
    sheet.reorder_worksheets([new_worksheet] + [ws for ws in worksheets if ws != new_worksheet])

    # header columns
    header_format = CellFormat(textFormat=TextFormat(bold=True))
    new_worksheet.append_row(["Transaction", "Post date", "Merchant", "Amount", "Notes", "Shoulder", "C", "S"])

    # apply bold text format to first row
    format_cell_range(new_worksheet, "1:1", header_format)

    # TODO: Add summary

    # load pulled pdf data to worksheet
    new_worksheet.append_rows(pdf_data, value_input_option="USER_ENTERED")

    grand_total = total_amount - total_cr
    
    # add total row
    new_worksheet.append_row(["", "", "TOTAL:", f"{total_amount:,.2f}"], value_input_option="USER_ENTERED")

    new_worksheet.append_row(["", "", "REIMBURSED TOTAL:", f"{total_cr:,.2f}"], value_input_option="USER_ENTERED")

    new_worksheet.append_row(["", "", "Grand Total:", f"{grand_total:,.2f}"], value_input_option="USER_ENTERED")

    return sheet


if __name__ == "__main__":
    # parse passed params
    parser = argparse.ArgumentParser(description="CC Analyzer")
    parser.add_argument("--bank", required=True, help="Name of the bank")

    args = parser.parse_args()

    # load passed params to variables
    bank = args.bank

    if not bank == "Security Bank":
        raise ValueError("Invalid bank name. Only 'Security Bank' is supported.")
    
    # input details
    sheet_name = input("Enter the name of the Google Sheet (Default: Security Bank World CC): ") or "Security Bank World CC"
    billing_period = input("Enter billing period(MM YYYY): ") # TODO: Validate format
    pdf_path = input("Enter the path to the PDF file(Default: dump/test.pdf): ") or "dump/test.pdf"
    pdf_password = input("Enter the password for the PDF file: ")
    
    try:
        pdf_data, total_amount, total_cr = extract_pdf_data(pdf_path, pdf_password)

        # sort from oldest
        pdf_data.sort(key=lambda x: datetime.strptime(x[1], "%m/%d/%y"))


        sheet = update_google_sheet(sheet_name, billing_period, pdf_data, total_amount, total_cr)

        sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet.id}"

        print("Success", sheet_url)
    except Exception as e:
        print(f"Error: {e}")