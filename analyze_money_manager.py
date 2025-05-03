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
    category_index = header_row.index("Category") + 1
    subcategory_index = header_row.index("Subcategory") + 1
    amount_index = header_row.index("Amount") + 1
    income_expense_column_index = header_row.index("Income/Expense") + 1
    note_index = header_row.index("Note") + 1

    total_expense = 0
    total_expense_count = 0
    total_income = 0
    total_income_count = 0
    
    # data
    expense_account_data = {}
    income_from_data = {}
    purchase_from_data = {}
    food_data = {}

    # special cases
    total_shopee_count = 0
    total_lazada_count = 0
    total_grab_count = 0
    total_711 = 0

    for row in sheet.iter_rows(min_row=2):
        period_cell = row[period_index - 1]
        account_cell = row[account_index - 1]
        category_cell = row[category_index - 1]
        subcategory_cell = row[subcategory_index - 1]
        amount_cell = row[amount_index - 1]
        income_expense_cell = row[income_expense_column_index - 1]
        note_cell = row[note_index - 1]

        # get the values from the cells
        period_value = period_cell.value
        category_value = category_cell.value
        subcategory_value = subcategory_cell.value
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
                    expense_account_data[account_value] = expense_account_data.get(account_value, 0) + 1

                if note_value:
                    # count purchase_from entry, total amount, and first instance
                    if note_value not in purchase_from_data:
                        purchase_from_data[note_value] = (0, 0, period_value.strftime("%B %d, %Y"))
                    current_count, current_amount_total, current_period_value = purchase_from_data[note_value]

                    # check if current_period_value is older than period_value then update it
                    if datetime.strptime(current_period_value, "%B %d, %Y") > period_value:
                        current_period_value = period_value.strftime("%B %d, %Y")
                    purchase_from_data[note_value] = (current_count + 1, round(current_amount_total + amount_value, 2), current_period_value)

                    if category_value == "Food":
                        if note_value not in food_data:
                            food_data[note_value] = (0, 0, period_value.strftime("%B %d, %Y"))
                        current_count, current_amount_total, current_period_value = food_data[note_value]

                        # check if current_period_value is older than period_value then update it
                        if datetime.strptime(current_period_value, "%B %d, %Y") > period_value:
                            current_period_value = period_value.strftime("%B %d, %Y")
                        food_data[note_value] = (current_count + 1, round(current_amount_total + amount_value, 2), current_period_value)

                    # catch shopee orders
                    if re.search(r'Shopee$', note_value, re.IGNORECASE):
                        total_shopee_count += 1

                    # catch lazada orders
                    if re.search(r'Lazada$', note_value, re.IGNORECASE):
                        total_lazada_count += 1

                    # catch 711 orders
                    if re.search(r'^711', note_value, re.IGNORECASE):
                        total_711 += 1

                    # catch grabfood and grabcar
                    if subcategory_value:
                        if re.search(r'Grab$', subcategory_value, re.IGNORECASE):
                            total_grab_count += 1
                            
                
            elif income_expense_value == "Income":
                total_income_count += 1
                total_income += amount_value

                if note_value:
                    # count income_from entry, total amount, and first instance
                    if note_value not in income_from_data:
                        income_from_data[note_value] = (0, 0, period_value.strftime("%B %d, %Y"))
                    current_count, current_amount_total, current_period_value = income_from_data[note_value]

                    # check if current_period_value is older than period_value then update it
                    if datetime.strptime(current_period_value, "%B %d, %Y") > period_value:
                        current_period_value = period_value.strftime("%B %d, %Y")
                    income_from_data[note_value] = (current_count + 1, round(current_amount_total + amount_value, 2), current_period_value)

    print(f"Total Income: {total_income:,.2f}")
    print(f"Total Expense: {total_expense:,.2f}")
    balance = total_income - total_expense
    print(f"Balance: {balance:,.2f}")
    
    print(f"\nTotal Income Entry: {total_income_count}")
    print(f"Total Expense Entry: {total_expense_count}")
    
    print("\nExpense Accounts:")
    pprint(sorted(expense_account_data.items(), key=lambda item: item[1], reverse=True))

    print("\nTop 10 Income From:")
    top_30_income_from_data = sorted(income_from_data.items(), key=lambda item: item[1], reverse=True)[:10]
    pprint(top_30_income_from_data)

    print("\nTop 30 Purchase From:")
    top_30_purchase_from_data = sorted(purchase_from_data.items(), key=lambda item: item[1], reverse=True)[:30]
    pprint(top_30_purchase_from_data)

    print("\nTop 30 Food:")
    top_30_food_data = sorted(food_data.items(), key=lambda item: item[1], reverse=True)[:30]
    pprint(top_30_food_data)

    print("\nSPECIAL CASES")
    print(f"Total Shopee Order Count: {total_shopee_count}")
    print(f"Total Lazada Order Count: {total_lazada_count}")
    print(f"Total Grab(GrabFood and GrabCar) Count: {total_grab_count}")
    print(f"Total 711 Count: {total_711}")

    # Top 10 Food

if __name__ == "__main__":
    try:
        # file_path = input("Enter the path to the .xlsx file: ")
        file_path = "./dump/mm.xlsx"
        analyze_money_manager(file_path)
    except Exception as e:
        print(f"Error: {e}")