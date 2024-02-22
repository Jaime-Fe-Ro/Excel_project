import yfinance as yf
import pandas as pd
import xlwings as xw
import datetime
import mibian


def update_data():
    # MANUAL - Define the expiry dates and strike upper bounds
    expiry_strike_bounds = {
        '2024-01-19': 825,
        '2024-03-15': 600,
        '2024-06-21': 758.33
    }

    # MANUAL - Set the names for the 'load_risk_free_interest_rate_and_TSLA_price' function
    file_path = r'C:\Users\jaime\OneDrive\Trading V2\TSLA.Data\Data.xlsx'
    sheet_name = r'Risk Free Interest Rate & $TSLA'
    TSLA_Price_table_name = r'TSLA_Price'
    RFIR_table_name = r'RFIR'

    # MANUAL - set ticker to download call option chain
    Ticker = r'TSLA'

    # Initialize an empty DataFrame to store the new option data
    new_data = pd.DataFrame()
    # Download options data for Tesla from yahoo finance
    # Get all expiration dates
    TSLA = yf.Ticker(Ticker)
    expirations = TSLA.options
    for expiration in expirations:
        # Get the option chain for this expiration
        opt_chain = TSLA.option_chain(expiration)
        # Get only call data and
        call_data = opt_chain.calls
        # Add the expiration date as a new column
        call_data['expiration_date'] = expiration
        # Append the call data to the new_data DataFrame
        new_data = pd.concat([new_data, call_data])

    # Initialize the array of rows to delete in the 'new_data' DataFrame
    rows_to_drop = []
    # Initialize a dictionary to store deletion reasons grouped by expiration date
    deletions_by_expiry = {}
    # Initialize a dictionary to store deletion reasons with counts
    deletion_reasons_count = {
        'low implied volatility': 0,
        'expires today': 0,
    }

    # Calculate Greeks for each row in the DataFrame
    # Get risk-free interest rate (RFIR) and Tesla's price (TSLA_Price) for the BS model
    # Call the function with the parameters that gets the values for RFIR and TSLA_Price from an Excel document
    result = load_risk_free_interest_rate_and_TSLA_price(file_path, sheet_name, TSLA_Price_table_name, RFIR_table_name)
    # Checks if the function returned error.
    if result is None:
        # Handle the error in running the function
        print("Could not calculate the Greeks")  # Exit the program with an error status code.
    else:
        # Unpack the result tuple into risk_free_rate and tesla_price
        risk_free_rate, tesla_price = result
        print(f"Risk-free interest rate: {risk_free_rate}%")
        print(f"Tesla share price: ${tesla_price}\n")

        for index, row in new_data.iterrows():
            expiration = row['expiration_date']
            # Check if the expiration date is in the expiry_strike_bounds dictionary
            if expiration in expiry_strike_bounds:
                # Check if the strike price is greater than the bound
                if row['strike'] > expiry_strike_bounds[expiration]:
                    rows_to_drop.append(index)
                    continue
            # Calculate days until expiration
            expiration_date = datetime.datetime.strptime(row['expiration_date'], '%Y-%m-%d').date()
            today = datetime.date.today()
            days_to_expiration = (expiration_date - today).days
            # Check if the option expires today
            if days_to_expiration < 1:
                deletion_reasons_count['expires today'] += 1
                reason = f"Strike {row['strike']} because it expires today."
                if expiration in deletions_by_expiry:
                    deletions_by_expiry[expiration].append(reason)
                else:
                    deletions_by_expiry[expiration] = [reason]
                rows_to_drop.append(index)
                continue

            # Calculate implied volatility as a percentage
            implied_volatility = row['impliedVolatility'] * 100
            # Check if the implied volatility is zero
            if implied_volatility < 0.01:
                deletion_reasons_count['low implied volatility'] += 1
                reason = f"Strike {row['strike']} because implied volatility = ({row['impliedVolatility']})"
                if expiration in deletions_by_expiry:
                    deletions_by_expiry[expiration].append(reason)
                else:
                    deletions_by_expiry[expiration] = [reason]
                    rows_to_drop.append(index)
                continue

            # Create a BS model using Mibian for each option
            bs = mibian.BS([tesla_price, row['strike'], risk_free_rate, days_to_expiration],
                           volatility=implied_volatility)
            # Add Greeks to the DataFrame
            new_data.at[index, 'Delta'] = bs.callDelta
            new_data.at[index, 'Gamma'] = bs.gamma
            new_data.at[index, 'Vega'] = bs.vega
            new_data.at[index, 'Theta'] = bs.callTheta
            new_data.at[index, 'Rho'] = bs.callRho

    # Drop the rows
    new_data.drop(rows_to_drop, inplace=True)

    # Print the deletion reasons grouped by expiration date
    print("Deleted:")
    for expiration, reasons in deletions_by_expiry.items():
        print(f"\nExpiration Date: {expiration}")
        for reason in reasons:
            print(f"- {reason}")

    # Connect to the target Excel workbook and get the active sheet
    wb = xw.Book(file_path)
    sheet = wb.sheets['Data']
    table = sheet.tables['Call_Data']

    # Check the number of rows in the 'Call_Data' table
    rows_in_table = table.range.rows.count

    # Check if there are data rows in the table (excluding the header)
    if rows_in_table > 1:
        # Determine the number of columns in the table
        cols_in_table = table.range.columns.count
        # Clear contents of all rows except the header
        table.range[1, 0].resize(rows_in_table - 1, cols_in_table).clear_contents()

    # Write the new data into the table, Adjust the starting cell if necessary to avoid overwriting the header row
    if len(new_data) > 0:
        table.range[1, 0].options(index=False, header=False).value = new_data

    # Print rows copied and rows deleted
    rows_deleted = len(rows_to_drop)
    expired_today = deletion_reasons_count['expires today']
    low_iv = deletion_reasons_count['low implied volatility']

    if expired_today > 0:
        print(f"\n{rows_deleted} rows deleted, {expired_today} expired today and {low_iv} had low implied volatility")
    else:
        print(f"\n{rows_deleted} rows deleted because of low implied volatility")

    # Save and close
    wb.save()
    wb.close()

    # Wait for confirmation before finishing
    wait_for_user()


# Function stops program to prevent closing window before user is done reading output
def wait_for_user():
    input("\nPress Enter to exit.")


def load_risk_free_interest_rate_and_TSLA_price(File_path_, Sheet_name_, TSLA_Price_table_name_, RFIR_table_name_):
    """
    Gets the values for the risk-free interest rate (RFIR) and Tesla's share price (TSLA_Price)
    It looks for the values in an Excel document saved locally to the computer - Data.xlsx
    Uses the xw library to get the values for the variables from tables in Excel.
    Returns the values for RFIR and TSLA_Price if valid ones are found,
    Returns None otherwise.
    """
    try:  # Try to open the Excel document using the xw library.
        with xw.Book(File_path_) as wb:
            sheet = wb.sheets[Sheet_name_]
            TSLA_Price = sheet.tables[TSLA_Price_table_name_].range[1, 1].value
            RFIR = sheet.tables[RFIR_table_name_].range[1, 0].value

            # If either RFIR or share price is invalid ( x <= 0 ), print error and return None
            if not isinstance(RFIR, (float, int)) or not isinstance(TSLA_Price, (float, int)):
                print("Invalid data type for RFIR or TSLA Price. Please check Data.xlsx")
                return None

            if RFIR <= 0 or RFIR > 100:
                print(f"{RFIR} is an invalid risk-free interest rate. Please check Data.xlsx")
                return None

            if TSLA_Price <= 0:
                print(f"${TSLA_Price} is an invalid share price for TSLA. Please check Data.xlsx")
                return None

            # Valid values found. Returns them in the order BRIF, TSLA_Price
            return RFIR, TSLA_Price

    # If the document cannot be reached, print error and exit:
    except Exception as e:
        print(f"Error accessing Excel file: {e}")
        return None


update_data()
