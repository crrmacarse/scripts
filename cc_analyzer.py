import gspread
from oauth2client.service_account import ServiceAccountCredentials
import argparse
from PyPDF2 import PdfReader
import re
from gspread_formatting import format_cell_range, CellFormat, TextFormat
from datetime import datetime

#TODO:
# - Integrate moneymanager which is a backup

def extract_pdf_data(pdf_path, pdf_password):
    # read PDF file
    reader = PdfReader(pdf_path)

    # decrypt PDF password
    if reader.is_encrypted:
        reader.decrypt(pdf_password)

    pdf_data = []
    total_amount = 0

    # extract row data from PDF
    for page in reader.pages:
        text = page.extract_text()

        table_pattern = re.compile(r"(\d{2}/\d{2}/\d{2})\s+(\d{2}/\d{2}/\d{2})\s+(.+?)\s+([\d,]+\.\d{2})")
        matches = table_pattern.findall(text)

        for match in matches:
            tran_date, post_date, description, amount = match
            # skips payment transactions
            if not description.startswith("PAYMENT"):
                total_amount += float(amount.replace(",", ""))
                pdf_data.append([tran_date, post_date, description, amount, "", "", "", ""])

    return pdf_data, total_amount

def update_google_sheet(sheet_name, billing_period, pdf_data, total_amount):
    # Google Sheets API setup
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("google-service-account.json", scope)

    client = gspread.authorize(creds)
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

    # add total row
    new_worksheet.append_row(["", "", "", f"{total_amount:,.2f}"], value_input_option="USER_ENTERED")

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
        pdf_data, total_amount = extract_pdf_data(pdf_path, pdf_password)

        # sort from oldest
        pdf_data.sort(key=lambda x: datetime.strptime(x[1], "%m/%d/%y"))


        sheet = update_google_sheet(sheet_name, billing_period, pdf_data, total_amount)

        sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet.id}"

        print("Success", sheet_url)
    except Exception as e:
        print(f"Error: {e}")