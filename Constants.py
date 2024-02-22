from datetime import datetime

TICKER = r'TSLA'
EXPIRY_AND_STRIKE_UPPER_LIMITS = {
    datetime(2024, 1, 19).date(): 825,
    datetime(2024, 3, 15).date(): 600,
    datetime(2024, 6, 21).date(): 758.33
}
EXCEL_FILE_PATH = r'C:\Users\jaime\OneDrive\Trading V2\TSLA.Data\Data.xlsx'
SHEET_NAME = r'Data'
TABLE_CALL_DATA = r'Call_data'
TABLE_TSLA_PRICE = r'TSLA_price'
TABLE_RFIR = r'RFIR'
EXPIRATION_COLUMN = r'expiration_date'
