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


# Generate Financial Statememts


# Geneerate the Profit and Loss account

def generate_profit_and_loss():
    """Generate and display the Profit and Loss Account"""
    # Extract data from Google Sheets
    sales_revenue = float(profit_and_loss_sheet.acell('B5').value.replace(',', ''))
    # Calculate cost of goods sold
    beginning_inventory = float(profit_and_loss_sheet.acell('B8').value.replace(',', ''))
    purchased_inventory = float(profit_and_loss_sheet.acell('B9').value.replace(',', ''))
    ending_inventory = float(profit_and_loss_sheet.acell('B10').value.replace(',', ''))
    cost_of_goods_sold = beginning_inventory + purchased_inventory - ending_inventory

    # Calculate gross profit
    gross_profit = sales_revenue - cost_of_goods_sold
    # Extract Operating expenses from google sheets
    payroll = float(profit_and_loss_sheet.acell('B16').value.replace(',', ''))
    utilities = float(profit_and_loss_sheet.acell('B17').value.replace(',', ''))
    rent_expense = float(profit_and_loss_sheet.acell('B18').value.replace(',', ''))
    advertising_marketing_expenses = float(profit_and_loss_sheet.acell('B19').value.replace(',', ''))
    depreciation_expense = float(profit_and_loss_sheet.acell('B20').value.replace(',', ''))
    # Calculate  total operating expenses
    total_operating_expenses = payroll + utilities + rent_expense + advertising_marketing_expenses + depreciation_expense
    # Calculate operating income
    operating_income = gross_profit - total_operating_expenses
    # Extract additional data from google sheets
    interest_expenses = float(profit_and_loss_sheet.acell('B25').value.replace(',', ''))
    # Calculate net income
    net_income = operating_income - interest_expenses

    # Display profit and loss account
    print("\nStep 3. Generate Financial Statements")
    print("\nProfit and Loss Account:")
    print("-------------------------------")
    print(f"Sales Revenue: ${sales_revenue:,.2f}")
    print(f"Cost of Goods Sold: ${cost_of_goods_sold:,.2f}")
    print("-------------------------------")
    print(f"Gross profit: ${gross_profit:,.2f}")
    print("-------------------------------")
    print(f"Operating Expenses:")
    print(f"Payroll: ${payroll:,.2f}")
    print(f"Utilities: ${utilities:,.2f}")
    print(f"Rent Expense: ${rent_expense:,.2f}")
    print(f"Advertising and Marketing Expenses: ${advertising_marketing_expenses:,.2f}")
    print(f"Depreciation Expense: ${depreciation_expense:,.2f}")
    print("-------------------------------")
    print(f"Total Operating Expenses: ${total_operating_expenses:,.2f}")
    print("-------------------------------")
    print(f"Operating Income: ${operating_income:,.2f}")
    print("-------------------------------")
    print(f"Interest Expenses: ${interest_expenses:,.2f}")
    print("-------------------------------")
    print(f"Net Income: ${net_income:,.2f}")
    print("-------------------------------")


def generate_balance_sheet():
    """Generate and display the Balance Sheet"""
    # Extract data from Google Sheets
    property_plant_equipment = float(balance_sheet.acell('B5').value.replace(',', ''))
    cash_and_equivalents = float(balance_sheet.acell('B8').value.replace(',', ''))
    accounts_receivable = float(balance_sheet.acell('B9').value.replace(',', ''))
    inventory = float(balance_sheet.acell('B10').value.replace(',', ''))
    # Calculate total assets
    total_assets = property_plant_equipment + cash_and_equivalents + accounts_receivable + inventory
    
    # Extract data from Google Sheets
    long_term_debt = float(balance_sheet.acell('B16').value.replace(',', ''))
    accounts_payable = float(balance_sheet.acell('B19').value.replace(',', ''))
    short_term_loans = float(balance_sheet.acell('B20').value.replace(',', ''))

    # Calculate total liabilities
    total_liabilities = long_term_debt + accounts_payable + short_term_loans

    # Extract data from Google Sheets
    common_stock = float(balance_sheet.acell('B25').value.replace(',', ''))
    retained_earnings = float(balance_sheet.acell('B26').value.replace(',', ''))
    
    # Calculate total liabilities and equity
    total_liabilities_and_equity = total_liabilities + common_stock + retained_earnings

    # Display the Balance Sheet
    print("\nBalance Sheet:")
    print("-------------------------------")
    print(f"Non-Current Assets:")
    print(f"Property, Plant and Equipment: ${property_plant_equipment:,.2f}")
    print(f"Current Assets:")
    print(f"Cash and Cash Equivalents: ${cash_and_equivalents:,.2f}")
    print(f"Accounts Receivable: ${accounts_receivable:,.2f}")
    print(f"Inventory: ${inventory:,.2f}")
    print("-------------------------------")
    print(f"Total Assets: ${total_assets:,.2f}")
    print("-------------------------------")
    print(f"Non-Current Liabilities:")
    print(f"Long-term Debt: ${long_term_debt:,.2f}")
    print(f"Current Liabilities:")
    print(f"Accounts Payable: ${accounts_payable:,.2f}")
    print(f"Short-Term Loans: ${short_term_loans:,.2f}")
    print("-------------------------------")
    print(f"Total Liabilities: ${total_liabilities:,.2f}")
    print("-------------------------------")
    print(f"Equity:")
    print(f"Common Stock: ${common_stock:,.2f}")
    print(f"Retained Earnings: ${retained_earnings:,.2f}")
    print("-------------------------------")
    print(f"Total Liabilities and Equity: ${total_liabilities_and_equity:,.2f}")
    print("-------------------------------")





def main():
    update_profit_and_loss()
    print("\nThe Profit and loss account has been updated")
    update_balance_sheet()
    print("\nThe Balance sheet has been updated")
    generate_profit_and_loss()
    generate_balance_sheet()

if __name__ == "__main__":
    main()