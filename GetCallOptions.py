import datetime

import pandas as pd
import Constants
import yfinance as yf
import xlwings as xw


def copy_call_option_chain_to_excel():
    trimmed_dataframe, deletion_summary = trim_call_option_chain(download_call_option_chain())

    # Clear contents of table
    workbook = xw.Book(Constants.EXCEL_FILE_PATH)
    sheet = workbook.sheets[Constants.SHEET_NAME]
    table = sheet.tables[Constants.TABLE_CALL_DATA]
    rows_in_table = table.range.rows.count
    if rows_in_table > 1:
        cols_in_table = table.range.columns.count
        table.range[1, 0].resize(rows_in_table - 1, cols_in_table).clear_contents()

    table.range[0, 0].options(index=False, header=True).value = trimmed_dataframe
    print_info(deletion_summary)
    input("Press enter to finish")


def download_call_option_chain():
    ticker = Constants.TICKER
    call_data_dataframe = pd.DataFrame()
    ticker_yf = yf.Ticker(ticker)
    expirations = ticker_yf.options
    for expiration in expirations:
        new_call_data = ticker_yf.option_chain(expiration).calls
        new_call_data[Constants.EXPIRATION_COLUMN] = expiration
        call_data_dataframe = pd.concat([call_data_dataframe, new_call_data]).reset_index(drop=True)
    return call_data_dataframe


def trim_call_option_chain(new_dataframe):
    initial_row_count = len(new_dataframe)
    deletions_summary = {}
    low_implied_volatility = 0
    expires_today = 0
    unavailable_strike = 0
    detailed_deletion_summary = {}
    rows_to_drop = []
    today = datetime.datetime.today().date()

    for index, row in new_dataframe.iterrows():
        expiration = datetime.datetime.strptime(row[Constants.EXPIRATION_COLUMN], '%Y-%m-%d').date()
        days_to_expiration = (expiration - today).days
        strike = row['strike']

        # Check for unavailable strikes
        if expiration in Constants.EXPIRY_AND_STRIKE_UPPER_LIMITS and \
                strike > Constants.EXPIRY_AND_STRIKE_UPPER_LIMITS[expiration]:
            rows_to_drop.append(index)
            unavailable_strike += 1
            # Add to logs
            if expiration not in deletions_summary:
                deletions_summary[expiration] = \
                    {"low_implied_volatility": 0, "expires_today": 0, "unavailable_strike": 0}
            deletions_summary[expiration]["unavailable_strike"] += 1
            if expiration not in detailed_deletion_summary:
                detailed_deletion_summary[expiration] = []
            reason = "it's an invalid strike price"
            detailed_deletion_summary[expiration].append({
                "strike": strike,
                "reason": reason
            })
            continue

        # Check for options expiring today
        if days_to_expiration < 1:
            rows_to_drop.append(index)
            expires_today += 1
            # Add to logs
            if expiration not in deletions_summary:
                deletions_summary[expiration] = \
                    {"low_implied_volatility": 0, "expires_today": 0, "unavailable_strike": 0}
            deletions_summary[expiration]["expires_today"] += 1
            if expiration not in detailed_deletion_summary:
                detailed_deletion_summary[expiration] = []
            reason = " expires today"
            detailed_deletion_summary[expiration].append({
                "strike": strike,
                "reason": reason
            })
            continue

        # Check for low implied volatility
        if row['impliedVolatility'] < 0.0001:
            rows_to_drop.append(index)
            low_implied_volatility += 1
            # Add to logs
            if expiration not in deletions_summary:
                deletions_summary[expiration] = \
                    {"low_implied_volatility": 0, "expires_today": 0, "unavailable_strike": 0}
            deletions_summary[expiration]["low_implied_volatility"] += 1
            if expiration not in detailed_deletion_summary:
                detailed_deletion_summary[expiration] = []
            reason = f"it has an implied volatility of {row['impliedVolatility']}"
            detailed_deletion_summary[expiration].append({
                "strike": strike,
                "reason": reason
            })
            continue

    new_dataframe.drop(rows_to_drop, inplace=True, axis=0)

    return new_dataframe, {
        "initial_row_count": initial_row_count,
        "total_deletions": len(rows_to_drop),
        "deleted_low_implied_volatility": low_implied_volatility,
        "deleted_expires_today": expires_today,
        "deleted_unavailable_strike": unavailable_strike,
        "trimmed_row_count": initial_row_count - len(rows_to_drop),
        "deletions_summary": deletions_summary,
        "detailed_deletion_summary": detailed_deletion_summary
    }


def print_info(deletion_summary):
    print("---------------------------------------------------------------------------------------\n"
          "\nDetailed deletion breakdown:")
    for expiry, details in deletion_summary["detailed_deletion_summary"].items():
        formatted_expiry = expiry.strftime('%d/%m/%Y')
        print(f"\n{formatted_expiry}:")
        for detail in details:
            print(f"Strike: {detail['strike']} because {detail['reason']}")

    print("\n\n---------------------------------------------------------------------------------------\n"
          "\nExpiry summary deletion breakdown:")
    for expiry, reasons in deletion_summary["deletions_summary"].items():
        if isinstance(expiry, datetime.date):
            formatted_expiry = expiry.strftime('%d/%m/%Y')
        else:
            formatted_expiry = datetime.datetime.strptime(expiry, '%Y-%m-%d').strftime('%d/%m/%Y')
        print(f"\n{formatted_expiry}:")
        if reasons["low_implied_volatility"] > 0:
            print(f"Implied volatility too low: {reasons['low_implied_volatility']}")
        if reasons["expires_today"] > 0:
            print(f"Expires today: {reasons['expires_today']}")
        if reasons["unavailable_strike"] > 0:
            print(f"Nonexistent strike: {reasons['unavailable_strike']}")

    print("\n\n---------------------------------------------------------------------------------------\n"
          "\nSummary deletion breakdown:")
    print(f"\nInitial Row Count: {deletion_summary['initial_row_count']}")
    print(f"Total Rows Deleted: {deletion_summary['total_deletions']}")
    if deletion_summary['deleted_low_implied_volatility'] > 0:
        print(f"    Low implied volatility: {deletion_summary['deleted_low_implied_volatility']}")
    if deletion_summary['deleted_expires_today'] > 0:
        print(f"    Expires today: {deletion_summary['deleted_expires_today']}")
    if deletion_summary['deleted_unavailable_strike'] > 0:
        print(f"    Unavailable strike: {deletion_summary['deleted_unavailable_strike']}")
    print(f"Trimmed Row Count: {deletion_summary['trimmed_row_count']}\n")


def get_risk_free_interest_rate_and_TSLA_price(path, sheet, table_price, table_rfir):
    try:
        with xw.Book(path) as wb:
            sheet = wb.sheets[sheet]
            TSLA_Price = sheet.tables[table_price].range[1, 1].value
            RFIR = sheet.tables[table_rfir].range[1, 0].value

            if not isinstance(RFIR, (float, int)) or not isinstance(TSLA_Price, (float, int)):
                print("Invalid data type for RFIR or TSLA Price. Please check Data.xlsx")
                return None

            if RFIR <= 0 or RFIR > 100:
                print(f"{RFIR} is an invalid risk-free interest rate. Please check Data.xlsx")
                return None

            if TSLA_Price <= 0:
                print(f"${TSLA_Price} is an invalid share price for TSLA. Please check Data.xlsx")
                return None

            print(TSLA_Price, RFIR)
            return RFIR, TSLA_Price

    except Exception as e:
        print(f"Error accessing Excel file: {e}")
        return None


def calculate_greeks():
    pass


def attach_greeks():
    pass


# file_path = Constants.EXCEL_FILE_PATH
# sheet_name = Constants.SHEET_NAME
# table_tsla_price = Constants.TABLE_TSLA_PRICE
# table_risk_free_ir = Constants.TABLE_RFIR
# get_risk_free_interest_rate_and_TSLA_price(file_path, sheet_name, table_tsla_price, table_risk_free_ir)


copy_call_option_chain_to_excel()
