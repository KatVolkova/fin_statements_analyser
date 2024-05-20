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
fin_ratios_sheet = SHEET.worksheet('ratios_historical_data')
benchmarks_sheet = SHEET.worksheet('industry_benchmarks')

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

# Function to extract data from the google sheets

def get_value(worksheet, cell):
    """Extract and convert cell value from the google sheets to float."""
    return float(worksheet.acell(cell).value.replace(',', ''))

# Geneerate the Profit and Loss account

def generate_profit_and_loss():
    """Generate and display the Profit and Loss Account"""
    # Extract data from Google Sheets
    sales_revenue = get_value(profit_and_loss_sheet, 'B5')
    beginning_inventory = get_value(profit_and_loss_sheet, 'B8')
    purchased_inventory = get_value(profit_and_loss_sheet, 'B9')
    ending_inventory = get_value(profit_and_loss_sheet, 'B10')
    payroll = get_value(profit_and_loss_sheet, 'B16')
    utilities = get_value(profit_and_loss_sheet, 'B17')
    rent_expense = get_value(profit_and_loss_sheet, 'B18')
    advertising_marketing_expenses = get_value(profit_and_loss_sheet, 'B19')
    depreciation_expense = get_value(profit_and_loss_sheet, 'B20')
    interest_expenses = get_value(profit_and_loss_sheet, 'B25')
    
    # Calculations
   
    cost_of_goods_sold = beginning_inventory + purchased_inventory - ending_inventory
    gross_profit = sales_revenue - cost_of_goods_sold
    total_operating_expenses = payroll + utilities + rent_expense + advertising_marketing_expenses + depreciation_expense
    operating_income = gross_profit - total_operating_expenses
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
    property_plant_equipment = get_value(balance_sheet,'B5')
    cash_and_equivalents = get_value(balance_sheet,'B8')
    accounts_receivable = get_value(balance_sheet,'B9')
    inventory = get_value(balance_sheet,'B10')
    long_term_debt = get_value(balance_sheet,'B16')
    accounts_payable = get_value(balance_sheet,'B19')
    short_term_loans = get_value(balance_sheet,'B20')
    common_stock = get_value(balance_sheet,'B25')
    retained_earnings = get_value(balance_sheet,'B26')

    # Calculations
    total_assets = property_plant_equipment + cash_and_equivalents + accounts_receivable + inventory
    total_liabilities = long_term_debt + accounts_payable + short_term_loans
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


# Calculate financial ratios

def calculate_liquidity_ratios():
    """Calculate liquidity ratios based on the updated financial statements: current and quick ratios"""
    # Extract values
    current_assets = get_value(balance_sheet, 'B11')
    current_liabilities = get_value(balance_sheet, 'B21')
    # Calculations and validation to avoid division by zero
    if current_liabilities == 0:
        raise ValueError("Current Liabilities should not be zero to avoid division by zero.")
    current_ratio = current_assets / current_liabilities
    quick_ratio = (current_assets - get_value(balance_sheet, 'B10')) / current_liabilities

    # Liquidity ratios results:
    print("\nLiquidity ratios:")
    print(f"\n\tCurrent ratio: {current_ratio:.2f} times")
    print(f"\n\tQuick ratio: {quick_ratio:.2f} times")
    return current_ratio, quick_ratio

def calculate_profitability_ratios():
    """Calculate profitability ratios based on the updated financial statements: net profit margin and return on assets"""
    # Extract values
    net_profit = get_value(profit_and_loss_sheet,'B27')
    sales_revenue = get_value(profit_and_loss_sheet,'B5')
    total_assets = get_value(balance_sheet,'B13')

    # Calculations and validation to avoid division by zero
    if sales_revenue == 0:
        raise ValueError("Sales Revenue should not be zero to avoid division by zero.")
    if total_assets == 0:
        raise ValueError("Total Assets should not be zero to avoid division by zero.")
    net_profit_margin = (net_profit / sales_revenue) * 100
    return_on_assets = (net_profit / total_assets) * 100

    # Profitability ratios results:
    print("\nProfitability ratios:")
    print(f"\n\tNet Profit Margin: {net_profit_margin:.2f}%")
    print(f"\n\tReturn on Assets: {return_on_assets:.2f}%")
    return net_profit_margin, return_on_assets
    
def calculate_solvency_ratios():
    """Calculate solvency ratios based on the updated financial statements: debt-to-equity and interest cover ratios"""
    # Extract values
    total_liabilities = get_value(balance_sheet,'B21')
    total_equity = get_value(balance_sheet,'B27')
    interest_expenses = get_value(profit_and_loss_sheet,'B25')

     # Calculations and validation to avoid division by zero
    if total_equity == 0:
        raise ValueError("Total Equity should not be zero to avoid division by zero.")
    if interest_expenses == 0:
        raise ValueError("Interest Expenses should not be zero to avoid division by zero.")
    debt_to_equity = (total_liabilities / total_equity) * 100
    interest_cover = (get_value(profit_and_loss_sheet,'B23') / interest_expenses)

    # Solvency ratios results:
    print("\nSolvency ratios:")
    print(f"\n\tDebt-to-Equity ratio: {debt_to_equity:.2f}%")
    print(f"\n\tInterest cover ratio: {interest_cover:.2f} times")
    return debt_to_equity, interest_cover
   
def update_ratios_googlews(current_ratio, quick_ratio, net_profit_margin, return_on_assets, debt_to_equity, interest_cover):
    """Update the Google Sheet with the calculated financial ratios"""
    ratios_to_update = {
        'Current Ratio': ('E5', f"{current_ratio:.2f}"),
        'Quick Ratio': ('E6', f"{quick_ratio:.2f}"),
        'Net Profit Margin': ('E8', f"{net_profit_margin:.2f}"),
        'Return on Assets': ('E9', f"{return_on_assets:.2f}"),
        'Debt-to-Equity Ratio': ('E11', f"{debt_to_equity:.2f}"),
        'Interest Cover Ratio': ('E12', f"{interest_cover:.2f}")
    }
    
    for ratio_name, (cell, value) in ratios_to_update.items():
        fin_ratios_sheet.update_acell(cell, value)
        print(f"\n\t{ratio_name} updated successfully")

# Analyse Financial Ratios results

def analyse_ratios(current_ratio, quick_ratio, net_profit_margin, return_on_assets, debt_to_equity, interest_cover):
    print("\nLiquidity Ratios Analysis:")
    if current_ratio >= 1.5:
        print("\n\tCurrent ratio indicates good liquidity.")
    elif 1 <= current_ratio < 1.5:
        print("\n\tCurrent ratio suggests adequate liquidity, but caution may be needed.")
    else:
        print("\n\tCurrent ratio indicates poor liquidity, and immediate attention may be required.")

    if quick_ratio >= 1:
        print("\n\tQuick ratio indicates strong ability to meet short-term obligations.")
    else:
        print("\n\tQuick ratio suggests potential difficulties in meeting short-term obligations.")

    print("\nProfitability Ratios Analysis:")
    if net_profit_margin > 10:
        print("\n\tNet profit margin indicates healthy profitability.")
    elif 5 <= net_profit_margin <= 10:
        print("\n\tNet profit margin suggests satisfactory profitability.")
    else:
        print("\n\tNet profit margin indicates low profitability, and improvement may be needed.")

    if return_on_assets > 10:
        print("\n\tReturn on assets suggests efficient utilization of assets.")
    elif 5 <= return_on_assets <= 10:
        print("\n\tReturn on assets indicates satisfactory performance in asset utilization.")
    else:
        print("\n\tReturn on assets suggests inefficiency in asset utilization, and optimization may be required.")

    print("\nSolvency Ratios Analysis:")
    if debt_to_equity < 50:
        print("\n\tDebt-to-equity ratio indicates low financial risk.")
    elif 50 <= debt_to_equity <= 100:
        print("\n\tDebt-to-equity ratio suggests moderate financial risk.")
    else:
        print("\n\tDebt-to-equity ratio indicates high financial risk, and debt management strategies may be needed.")

    if interest_cover >= 2:
        print("\n\tInterest cover ratio indicates comfortable ability to cover interest expenses.")
    else:
        print("\n\tInterest cover ratio suggests potential challenges in covering interest expenses.")

# Actual vs Benchmark ratios comparison

# Extract Benchmark Values

def get_benchmarks():
    """Extract benchmark values from Google sheets"""
    return {
        'current_ratio': get_value(benchmarks_sheet,'B4'),
        'quick_ratio': get_value(benchmarks_sheet,'B5'),
        'net_profit_margin': get_value(benchmarks_sheet,'B6'),
        'return_on_assets': get_value(benchmarks_sheet,'B7'),
        'debt_to_equity': get_value(benchmarks_sheet,'B8'),
        'interest_cover': get_value(benchmarks_sheet,'B9')
    }

# Compare Actual results to Benchmarks

def compare_with_benchmarks(benchmarks, current_ratio, quick_ratio, net_profit_margin, return_on_assets,
                            debt_to_equity, interest_cover):
    """Compare actual results with benchmark values"""
    ratios = {
        'Current Ratio': current_ratio,
        'Quick Ratio': quick_ratio,
        'Net Profit Margin': net_profit_margin,
        'Return on Assets': return_on_assets,
        'Debt to Equity': debt_to_equity,
        'Interest Cover': interest_cover
    }

    print("\nBenchmark Comparison Analysis:")
    for ratio, value in ratios.items():
        benchmark_value = benchmarks[ratio.lower().replace(' ', '_')]
        unit = 'times' if ratio in ['Current Ratio', 'Quick Ratio', 'Interest Cover'] else '%'
        print(f"\n\t{ratio}: {value:.2f} {unit} (Benchmark: {benchmark_value:.2f} {unit})")
        if value > benchmark_value:
            print(f"\n\tThe {ratio} is above the industry benchmark, indicating better performance.")
        elif value < benchmark_value:
            print(f"\n\tThe {ratio} is below the industry benchmark, indicating underperformance.")
        else:
            print(f"\n\tThe {ratio} is equal to the industry benchmark, indicating average performance.")

# Carry out the trend analysis based on historical data

# Extract historical data

def get_historical_data():
    """Extract historical data for four quarters from Google Sheets."""
    quarters = ['Q1', 'Q2', 'Q3', 'Q4']
    columns = 'BCDE'

    historical_data = {
        'current_ratio': {quarter: get_value(fin_ratios_sheet, f'{column}5') for quarter, column in zip(quarters, columns)},
        'quick_ratio': {quarter: get_value(fin_ratios_sheet, f'{column}6') for quarter, column in zip(quarters, columns)},
        'net_profit_margin': {quarter: get_value(fin_ratios_sheet, f'{column}8') for quarter, column in zip(quarters, columns)},
        'return_on_assets': {quarter: get_value(fin_ratios_sheet, f'{column}9') for quarter, column in zip(quarters, columns)},
        'debt_to_equity': {quarter: get_value(fin_ratios_sheet, f'{column}11') for quarter, column in zip(quarters, columns)},
        'interest_cover': {quarter: get_value(fin_ratios_sheet, f'{column}12') for quarter, column in zip(quarters, columns)}
    }

    return historical_data


def main():
    update_profit_and_loss()
    print("\nThe Profit and loss account has been updated")
    update_balance_sheet()
    print("\nThe Balance sheet has been updated")
    generate_profit_and_loss()
    generate_balance_sheet()
    print("\nStep 4. Calculate Financial Ratios")
    current_ratio, quick_ratio = calculate_liquidity_ratios()
    net_profit_margin, return_on_assets = calculate_profitability_ratios()
    debt_to_equity, interest_cover = calculate_solvency_ratios()
    print("\nStep 5. Update Google sheets with calculated ratios numbers:")
    update_ratios_googlews(current_ratio, quick_ratio, net_profit_margin, return_on_assets, debt_to_equity, interest_cover)
    print("\nStep 6. Analyse Financial Results:")
    analyse_ratios(current_ratio, quick_ratio, net_profit_margin, return_on_assets, debt_to_equity, interest_cover)
    print("\nStep 7. Carry out benchmark comparison:")
    benchmarks = get_benchmarks()
    compare_with_benchmarks(benchmarks, current_ratio, quick_ratio, net_profit_margin, return_on_assets,
                                        debt_to_equity, interest_cover)

if __name__ == "__main__":
    main()