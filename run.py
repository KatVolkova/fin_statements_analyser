import gspread
from google.oauth2.service_account import Credentials

# Authenticate and authorise the google sheets API

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('financial_statements')


# Get the worksheets objects
profit_and_loss_sheet = SHEET.worksheet('profit_and_loss')
balance_sheet = SHEET.worksheet('balance_sheet')


# Financial statements update

# Function to validate user's input

def validate_input(message, min_value, max_value, account_name):
    """Validate the user's input to ensure that only a number within defined range is entered"""
    while True:
        try:
            value = input(message).replace(',', '')
            if value.lower() == 'exit':
                print("Exiting the program.")
                exit()
            if not value.isdigit():
                raise ValueError("Input must contain only digits.")
            if value.startswith('0') and value != '0':
                raise ValueError("Leading zeros are not allowed.")
            value = float(value)
            if min_value <= value <= max_value:
                return value
            else:
                message = f"Please enter the number for {account_name}. The number should not contain any decimal places. The number should be between {min_value:,.0f} and {max_value:,.0f}:"
                print(f"The number should be between {min_value:,.0f} and {max_value:,.0f}. The number should not contain any decimal places. Please try again.")
        except ValueError as e:
            print(f"Invalid input: {e}")

# Function to update the profit and loss account

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
        message = f"\nPlease enter the number for {account_name}. The number should not contain any decimal places. The number should be between {min_value:,.0f} and {max_value:,.0f}: "
        value = validate_input(message, min_value, max_value, account_name)
        profit_and_loss_sheet.update_acell(cell,value)
        print(f"\n\t{account_name} updated successfully")

   
# Funciton to update the balance sheet 

def update_balance_sheet():
    """Update the Balance Sheet numbers for particular lines in Google sheets balance_sheet tab"""  
    accounts_to_update = {
        'Cash and Cash Equivalents': ('B8', 5_000, 100_000),
        'Short-Term Loans': ('B20', 1_000, 50_000)
    }
    
    print("\nStep 2. Update the Balance Sheet numbers:")
    for account_name, (cell, min_value, max_value) in accounts_to_update.items():
        message = f"\nPlease enter the number for {account_name}. The number should not contain any decimal places. The number should be between {min_value:,.0f} and {max_value:,.0f}: "
        value = validate_input(message, min_value, max_value, account_name)
        balance_sheet.update_acell(cell, value)
        print(f"\n\t{account_name} updated successfully") 

def main():
    update_profit_and_loss()
    print("\nThe Profit and loss account has been updated")
    update_balance_sheet()
    print("\nThe Balance sheet has been updated")

if __name__ == "__main__":
    main()