import gspread
from google.oauth2.service_account import Credentials

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('financial_statements')

profit_and_loss_sheet = SHEET.worksheet('profit_and_loss')
balance_sheet = SHEET.worksheet('balance_sheet')


def update_profit_and_loss():
    """Update the profit and loss account numbers for particular lines in Google sheets profit_and_loss tab"""
    print("Step 1. Update Profit and Loss account numbers:")

    sales_revenue = int(input("Please enter the number for sales revenue. The number should be between 50,000 and 500,000. "))
    profit_and_loss_sheet.update_acell('B5',sales_revenue)

    purchased_inventory = int(input("Please enter the number for inventory. The number should be between 10,000 and 200,000. "))
    profit_and_loss_sheet.update_acell('B9',purchased_inventory)

    rent = int(input("Please enter the number for rent. The number should be between 5,000 and 20,000. "))
    profit_and_loss_sheet.update_acell('B18',rent)

    interest_expenses = int(input("Please enter the number for interest expenses.The number should be between 1,000 and 10,000. "))
    profit_and_loss_sheet.update_acell('B25',interest_expenses)

def update_balance_sheet():
    """Update the Balance Sheet numbers for particular lines in Google sheets balance_sheet tab"""  
    print("The following should be updated in BS:\n\tCash and Cash Equivalents\n\tShort-Term Loans")

    cash_and_equivalents = int(input("Please enter the number for cash. The number should be between 5,000 and 100,000. "))
    balance_sheet.update_acell('B8',cash_and_equivalents)

    short_term_loans = int(input("Please enter the number for short-term loans. The number should be between 1,000 and 50,000. "))
    balance_sheet.update_acell('B20',short_term_loans)    

def main():
    update_profit_and_loss()
    print("The Profit and loss account has been updated")
    update_balance_sheet()
    print("The Balance sheet has been updated")

if __name__ == "__main__":
    main()