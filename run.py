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
    

    accounts_to_update = {
        'Sales Revenue': ('B5',50_000,500_000),
        'Purchased Inventory': ('B9',10_000,200_000),
        'Rent': ('B18',5_000,20_000),
        'Interest Expense': ('B25',1_000,10_000)
    }
    
    print("\nStep 1. Update Profit and Loss account numbers:")
    for account_name, (cell, min_value, max_value) in accounts_to_update.items():
        message = f"\nPlease enter the number for {account_name}. The number should be between {min_value} and {max_value}: "
        value = input(message)
        profit_and_loss_sheet.update_acell(cell,value)
        print(f"\n\t{account_name} updated successfully")

   

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