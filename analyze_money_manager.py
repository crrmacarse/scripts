import re
from openpyxl import load_workbook
from pprint import pprint
from datetime import datetime 

def analyze_money_manager(file_path):
    workbook = load_workbook(file_path)
    sheet = workbook.active

    header_row  = next(sheet.iter_rows(values_only=True))

    period_index = header_row.index("Period") + 1
    account_index = header_row.index("Accounts") + 1
    amount_index = header_row.index("Amount") + 1
    income_expense_column_index = header_row.index("Income/Expense") + 1
    note_index = header_row.index("Note") + 1

    total_expense = 0
    total_expense_count = 0
    total_income = 0
    total_income_count = 0
    total_shopee_count = 0
    expense_account_counts = {}

    income_from_counts = {}
    purchase_from_counts = {}

    for row in sheet.iter_rows(min_row=2):
        period_cell = row[period_index - 1]
        account_cell = row[account_index - 1]
        amount_cell = row[amount_index - 1]
        income_expense_cell = row[income_expense_column_index - 1]
        note_cell = row[note_index - 1]

        # get the values from the cells
        period_value = period_cell.value
        amount_value = amount_cell.value
        income_expense_value = income_expense_cell.value
        account_value = account_cell.value
        note_value = note_cell.value

        if isinstance(amount_value, (int, float)):
            if income_expense_value == "Exp.":
                total_expense_count += 1
                total_expense += amount_value

                # count expense accounts
                if account_value:
                    expense_account_counts[account_value] = expense_account_counts.get(account_value, 0) + 1

                if note_value:
                    # count purchase_from entry, total amount, and first instance
                    if note_value not in purchase_from_counts:
                        purchase_from_counts[note_value] = (0, 0, period_value.strftime("%B %d, %Y"))
                    current_count, current_amount_total, current_period_value = purchase_from_counts[note_value]

                    # check if current_period_value is older than period_value then update it
                    if datetime.strptime(current_period_value, "%B %d, %Y") > period_value:
                        current_period_value = period_value.strftime("%B %d, %Y")
                    purchase_from_counts[note_value] = (current_count + 1, round(current_amount_total + amount_value, 2), current_period_value)

                    # catch shopee orders
                    if re.search(r'Shopee$', note_value, re.IGNORECASE):
                        total_shopee_count += 1
                
            elif income_expense_value == "Income":
                total_income_count += 1
                total_income += amount_value

                if note_value:
                    # count income_from entry, total amount, and first instance
                    if note_value not in income_from_counts:
                        income_from_counts[note_value] = (0, 0, period_value.strftime("%B %d, %Y"))
                    current_count, current_amount_total, current_period_value = income_from_counts[note_value]

                    # check if current_period_value is older than period_value then update it
                    if datetime.strptime(current_period_value, "%B %d, %Y") > period_value:
                        current_period_value = period_value.strftime("%B %d, %Y")
                    income_from_counts[note_value] = (current_count + 1, round(current_amount_total + amount_value, 2), current_period_value)

    print(f"Total Income: {total_income:,.2f}")
    print(f"Total Expense: {total_expense:,.2f}")
    balance = total_income - total_expense
    print(f"Balance: {balance:,.2f}")
    
    print(f"\nTotal Income Entry: {total_income_count}")
    print(f"Total Expense Entry: {total_expense_count}")
    
    print("\nExpense Account Counts:")
    pprint(sorted(expense_account_counts.items(), key=lambda item: item[1], reverse=True))

    print("\nTop 10 Income From:")
    top_30_income_from_counts = sorted(income_from_counts.items(), key=lambda item: item[1], reverse=True)[:10]
    pprint(top_30_income_from_counts)

    print("\nTop 30 Purchase From:")
    top_30_purchase_from_counts = sorted(purchase_from_counts.items(), key=lambda item: item[1], reverse=True)[:30]
    pprint(top_30_purchase_from_counts)

    print(f"\nTotal Shopee Order Count: {total_shopee_count}")

if __name__ == "__main__":
    try:
        # file_path = input("Enter the path to the .xlsx file: ")
        file_path = "./dump/mm.xlsx"
        analyze_money_manager(file_path)
    except Exception as e:
        print(f"Error: {e}")