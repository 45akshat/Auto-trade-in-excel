import threading
import xlwings as xw
import time
from datetime import datetime
import os
import pandas as pd
import json
import requests
from SmartApi import SmartConnect
import pyotp
from pywintypes import com_error

import warnings
warnings.filterwarnings("ignore") 

# Connect to WebSocket for live ticks
from SmartApi.smartWebSocketV2 import SmartWebSocketV2
from logzero import logger



# Global variables
sws = None
connected = False
previous_data_token = None
AUTH_TOKEN = None
FEED_TOKEN = None
obj = SmartConnect(api_key="api key here")

def setup_smartapi_session():
    global AUTH_TOKEN, FEED_TOKEN, obj
    # data = obj.generateSession("user id", "pass", pyotp.TOTP("totp here").now())

    # print(data)
    AUTH_TOKEN = data['data']['jwtToken']
    refreshToken = data['data']['refreshToken']
    FEED_TOKEN = obj.getfeedToken()
    res = obj.getProfile(refreshToken)
    print("*****************")
    print(obj.rmsLimit())
    print("*****************")

    print(res)

# Initialize SmartWebSocketV2 object
def initialize_sws():
    global sws
    sws = SmartWebSocketV2(AUTH_TOKEN, obj.api_key, "p37872", FEED_TOKEN)
    sws.on_open = on_open
    sws.on_data = on_data
    sws.on_error = on_error
    sws.on_close = on_close
    sws.on_control_message = on_control_message

def printNo(message):
    print(type(message))


# Modify the existing function to update ticks in Excel
# Define a dictionary to store tick data for symbols
tick_data_dict = {}
pending_order_dict = {} 
exit_data_dict = {}


def clear_sheet():
    # Load the Excel workbook
    wb = xw.Book("excelbast.xlsx")
    
    # Select the 'OrderBook' sheet
    sheet = wb.sheets['OrderBook']
    sheet.clear_contents()

    sheet = wb.sheets['Position']
    sheet.clear_contents()

    sheet = wb.sheets['Holdings']
    sheet.clear_contents()

    sheet = wb.sheets['Funds']
    sheet.clear_contents()

    # Select the 'Live' sheet
    sheet = wb.sheets['Live']

    # Clear the cells from J2 to K500
    # sheet.range('J2:K500').clear_contents()

    
    # Clear the contents of all cells in the sheet

# Call the function to clear the sheet
clear_sheet()

def track_pending_orders(symbol, ticks_data):
    # print("trackng start")
    # Check if there are pending orders for this symbol
    for row_idx, row_data in list(pending_order_dict.items()):
        if row_data[0] == symbol:  # Check if symbol matches
            last_traded_price = ticks_data['last_traded_price']
            limit_price = row_data[7]
            signal = row_data[8]
            if signal == "limit_above" and last_traded_price > limit_price:
                print("Place order for symbol:", symbol, row_idx)
                tokenForSymbol = find_token(row_data[0], json_data)
                orderAngel(row_data[0], tokenForSymbol, row_data[9], row_data[10], row_idx)
                # Call your placeOrder function here if the condition is met
                # Remove the pending order once the order is placed
                pending_order_dict.pop(row_idx)
            elif signal == "limit_below" and last_traded_price < limit_price:
                print("Place order for symbol:", symbol, row_idx)
                tokenForSymbol = find_token(row_data[0], json_data)
                orderAngel(row_data[0], tokenForSymbol, row_data[9], row_data[10], row_idx)
                # Call your placeOrder function here if the condition is met
                # Remove the pending order once the order is placed
                pending_order_dict.pop(row_idx)
    # print("trackng end")


def track_exit_orders():
    print("Tracking exit orders start", exit_data_dict)
    
    while True:
        try:
            for row_idx, row_data in list(exit_data_dict.items()):  
                if len(row_data) < 14:
                    continue  # Skip incomplete data rows

                symbol = row_data[0]
                ticks_data = tick_data_dict.get(symbol)
                # print(ticks_data)
                if ticks_data is None:
                    continue  # Skip if symbol not found in tick_data_dict

                buy_sell = row_data[10]
                order_price = row_data[13]
                qty = row_data[10]

                last_traded_price = ticks_data.get('last_traded_price')
                if last_traded_price is None:
                    continue

                if buy_sell == 1:  # Buy order
                    target = row_data[14]
                    stoploss = row_data[15]
                    if target is not None:
                        if last_traded_price >= target:
                            print(f"Exit order for symbol {symbol}, row {row_idx}: Hit target. {row_data}")
                            # Implement exit order logic here for hitting target
                            exit_data_dict.pop(row_idx)  # Remove the exited order from exit_data_dict
                    if stoploss is not None:
                        if last_traded_price <= stoploss:
                            print(f"Exit order for symbol {symbol}, row {row_idx}: Hit stoploss. {row_data}")
                            # Implement exit order logic here for hitting stoploss
                            exit_data_dict.pop(row_idx)  # Remove the exited order from exit_data_dict
                elif buy_sell == 2:  # Sell order
                    target = row_data[14]
                    stoploss = row_data[15]
                    if target is not None:
                        if last_traded_price <= target:
                            print(f"Exit order for symbol {symbol}, row {row_idx}: Hit target. {row_data}")
                            # Implement exit order logic here for hitting target
                            exit_data_dict.pop(row_idx)  # Remove the exited order from exit_data_dict
                    if stoploss is not None:
                        if last_traded_price >= stoploss:
                            print(f"Exit order for symbol {symbol}, row {row_idx}: Hit stoploss. {row_data}")
                            # Implement exit order logic here for hitting stoploss
                            exit_data_dict.pop(row_idx)  # Remove the exited order from exit_data_dict
            # Sleep for some time before checking again
            time.sleep(2)
        except Exception as e:
            # print("An error occurred:", e)
            # print("Retrying in 5 seconds...")
            time.sleep(5)
            continue

    print("Tracking exit orders end") 

def get_symbols_and_row_numbers(sheet):
    symbols_and_row_numbers = []

    for i, value in enumerate(sheet.range('A2:A501').value):
        if value:
            symbols_and_row_numbers.append((value, i + 2))  # Add 2 to convert index to row number
    return symbols_and_row_numbers

def create_main_list():
    while True:
        main_list = []
        try:
            # print("main list start")
            wb = xw.Book("excelbast.xlsx")
            sheet = wb.sheets['Live']  # Assuming the sheet name is 'Live'
            
            # Fetch symbols and row numbers in bulk
            try:
                symbols_range = sheet.range('A2:A501').value
            except AttributeError as e:
                return
            except Exception as e:
                return
            
            try:
                row_numbers = list(range(2, 502))
            except AttributeError as e:
                return
            except Exception as e:
                return
            
            try:
                for symbol, row_number in zip(symbols_range, row_numbers):
                    if symbol and symbol in tick_data_dict:
                        tick_data = tick_data_dict[symbol]
                        row_data = [
                            tick_data.get('last_traded_price'),
                            tick_data.get('open_price_of_the_day'),
                            tick_data.get('high_price_of_the_day'),
                            tick_data.get('low_price_of_the_day'),
                            tick_data.get('closed_price'),
                            tick_data.get('open_interest')
                        ]
                        main_list.append(row_data)
                    else:
                        main_list.append([None] * 6)  # Fill with None if symbol not found or row has no symbol
                # Define the range where the data will be updated
                update_range = sheet.range('B2').expand('table')  # Expand to include the entire table starting from B2
                
                # Update the Excel sheet with the main_list data
                update_range.value = main_list
                # print("main list end")
            
            except AttributeError as e:
                time.sleep(5)
                continue
            except Exception as e:
                time.sleep(5)
                continue

        except com_error as e:
            time.sleep(5)
            continue

        except Exception as e:
            time.sleep(5)
            continue

        time.sleep(0.3)  # Wait for 1 second before the next update

# Example usage:
# Assuming tick_data_dict is your tick data dictionary
# main_list = create_main_list(tick_data_dict)

# Example usage:
# Assuming tick_data_dict is your tick data dictionary


def on_data(wsapp, message):
    ticks_data = message
    # print(ticks_data)
    # print(ticks_data['last_traded_price'])

    # Convert values to two decimal places
    ticks_data['last_traded_price'] = round(ticks_data['last_traded_price'] / 100, 2)
    ticks_data['open_price_of_the_day'] = round(ticks_data['open_price_of_the_day'] / 100, 2)
    ticks_data['high_price_of_the_day'] = round(ticks_data['high_price_of_the_day'] / 100, 2)
    ticks_data['low_price_of_the_day'] = round(ticks_data['low_price_of_the_day'] / 100, 2)
    ticks_data['closed_price'] = round(ticks_data['closed_price'] / 100, 2)
    
    token = ticks_data['token']
    symbol = find_symbol_from_token(token)
    if symbol:
        if symbol in list(tick_data_dict):
            # Update existing tick data for the symbol
            tick_data_dict[symbol] = ticks_data
        else:
            # Store tick data for the new symbol
            tick_data_dict[symbol] = ticks_data

        # create_main_list(tick_data_dict)
        track_pending_orders(symbol, ticks_data)



    else:
        print(f"Symbol not found for token: {token}")

# Define a threading lock
thread_lock = threading.Lock()

# def update_excel_at_interval():
    # if not thread_lock.locked():
    #     with thread_lock:
            # for symbol, ticks_data in list(tick_data_dict.items()):

    # Schedule the function to be called again after 3 seconds
    # threading.Timer(0.3, update_excel_at_interval).start()

# update_excel_at_interval()

def find_symbol_from_token(token):
    for item in json_data:
        if item['token'] == token:
            return item['symbol']
    return None




def on_control_message(wsapp, message):
    logger.info(f"Control Message: {message}")

def on_open(wsapp):
    global connected
    connected = True
    logger.info("on open")

def on_error(wsapp, error):
    logger.error("ERROR RECON ON")
    global connected
    
    global previous_data_token
    previous_data_token = None
    connected = False
    attempt_reconnect()


def on_close(wsapp):
    global connected
    global previous_data_token
    previous_data_token = None
    connected = False
    logger.info("Close")
    attempt_reconnect()

def attempt_reconnect():
    global connected

    while not connected:
        try:
            logger.info("Attempting to reconnect...")
            time.sleep(5)
            setup_smartapi_session()
            initialize_sws()
            sws.connect()
        except Exception as e:
            logger.error(f"Reconnection attempt failed: {e}")


# Assign callback functions

setup_smartapi_session()
initialize_sws()
def start_websocket():
    sws.connect()

def update_ticks_in_excel(symbol, ticks_data):
    # print(symbol, ticks_data.get("last_traded_price"))
    # Open or create Excel file
    file_name = "excelbast.xlsx"
    file_path = os.path.join(os.getcwd(), file_name)
    if os.path.exists(file_path):
        try:
            wb = xw.Book(file_path)
        except Exception as e:
            return
    else:
        wb = xw.Book()
    
    # Check if workbook is properly opened
    if wb is None:
        return
    
    # Access active sheet
    try:
        sheet = wb.sheets['Live']
    except Exception as e:
        return
    
    try:
        row_numbers = [i + 1 for i, value in enumerate(sheet.range('A1:A500').value) if value == symbol]
    except AttributeError as e:
        return
    except Exception as e:
        return

    # If the symbol is found, prepare the data to be updated
    if row_numbers:
        data_to_update = [
            [
                ticks_data.get('last_traded_price'),
                ticks_data.get('open_price_of_the_day'),
                ticks_data.get('high_price_of_the_day'),
                ticks_data.get('low_price_of_the_day'),
                ticks_data.get('closed_price'),
                ticks_data.get('open_interest')
            ]
        ]

        # Attempt to update the Excel sheet with extracted data for each occurrence of the symbol
        for row_number in row_numbers:
            try:
                sheet.range(f'B{row_number}:G{row_number}').value = data_to_update
            except AttributeError as e:
                break
                print("Error updating Excel sheet, retrying after 1 second:", e)
                time.sleep(1)
                try:
                    sheet = wb.sheets['Live']  # Re-access the sheet in case it became invalid
                    sheet.range(f'B{row_number}:G{row_number}').value = data_to_update
                except Exception as e:
                    print("Error updating Excel sheet:", e)
    else:
        print(f"Symbol '{symbol}' not found in Excel sheet.")

# Function to find ticks for a given token
def find_ticks_for_token(token, exc):
    global sws

    correlation_id = "abc123"
    action = 1
    mode = 3
    print(token)
    token_list = [{
        "exchangeType": exc,  # Assuming NSE
        "tokens": [token]  # Token for the desired stock
    }]
    # Subscribe to the desired stock
    sws.subscribe(correlation_id, mode, token_list)



global json_data



# -----------------------------
def create_token_sheet():
    global json_data
    
    url = "https://margincalculator.angelbroking.com/OpenAPI_File/files/OpenAPIScripMaster.json"

    # Fetch the JSON data from the API endpoint
    response = requests.get(url)
    json_data = response.json()
    # Filter and process data
    new_array = [item for item in json_data if item['symbol'].startswith('NIFTY')]
    new_array = [item for item in new_array if 'NFO' in item['exch_seg']]
    nse_json = [item for item in json_data if item['exch_seg'].startswith('N')]

    # print(len(json_data))
    # Construct DataFrame from filtered data
    df = pd.DataFrame(nse_json)

    # Open or create Excel file
    file_name = "excelbast.xlsx"
    wb = xw.Book(file_name)

    # Access or create the "Token" sheet
    try:
        sheet = wb.sheets["Token"]
    except:
        sheet = wb.sheets.add("Token")

    # Write the DataFrame to the worksheet
    sheet.range("A1").value = df

    # Save and close the workbook
    wb.save(file_name)
    return

# -----------------------------
def update_order_book_sheet():
    while True:
        try:
            # Fetch order book details
            holdings_response = obj.orderBook()
            
            # Check if the response is valid
            if 'data' not in holdings_response:
                print("Holdings data not found in the response")
                return

            holdings_data = holdings_response['data']
            
            # Load the Excel workbook and the specific sheet for holdings
            wb = xw.Book("excelbast.xlsx")
            sheet = wb.sheets['OrderBook']  # Assuming the sheet name is 'OrderBook'

            # Convert holdings data to a DataFrame
            df_holdings = pd.DataFrame(holdings_data)

            # Reorder columns to place 'updatetime' at the first position
            columns = ['updatetime'] + [col for col in df_holdings.columns if col != 'updatetime']
            df_holdings = df_holdings[columns]

            # Write the DataFrame to the Excel sheet
            sheet.range("A1").options(index=False).value = df_holdings
            time.sleep(1.5)
            # print("Holdings updated successfully")
        except Exception as e:
            # print(f"An error occurred while fetching or updating holdings: {e}")
            time.sleep(1.5)  # Wait for 2 seconds before continuing


def update_position_sheet():

    while True:
        try:
            # Fetch order book details
            
            holdings_response = obj.position()
            
            # Check if the response is valid
            if 'data' not in holdings_response:
                print("Holdings data not found in the response")
                return

            holdings_data = holdings_response['data']
            
            # Load the Excel workbook and the specific sheet for holdings
            wb = xw.Book("excelbast.xlsx")
            sheet = wb.sheets['Position']  # Assuming the sheet name is 'Holding'

            # Convert holdings data to a DataFrame
            df_holdings = pd.DataFrame(holdings_data)

            # Write the DataFrame to the Excel sheet
            sheet.range("A1").options(index=False).value = df_holdings
            time.sleep(1.5)
            # print("Holdings updated successfully")
        except Exception as e:
            # print(f"An error occurred while fetching or updating holdings: {e}")
            time.sleep(1.5)  # Wait for 2 seconds before continuing


# -----------------------------
def update_holdings_sheet():

    while True:
        try:
            # Fetch order book details
            
            holdings_response = obj.allholding()
            # print(obj.getProfile)
            
            # Check if the response is valid
            if 'data' not in holdings_response:
                print("Holdings data not found in the response")
                return

            holdings_data = holdings_response['data']['holdings']
            
            # Load the Excel workbook and the specific sheet for holdings
            wb = xw.Book("excelbast.xlsx")
            sheet = wb.sheets['Holdings']  # Assuming the sheet name is 'Holding'

            # Convert holdings data to a DataFrame
            df_holdings = pd.DataFrame(holdings_data)

            # Write the DataFrame to the Excel sheet
            sheet.range("A1").options(index=False).value = df_holdings
            time.sleep(1.5)
            # print("Holdings updated successfully")
        except Exception as e:
            # print(f"An error occurred while fetching or updating holdings: {e}")
            time.sleep(1.5)  # Wait for 2 seconds before continuing


def fetch_and_update_funds():
    
    while True:
        try:
            funds_response = obj.rmsLimit()
            
            # Check if the response is valid
            if 'data' not in funds_response:
                print("Funds data not found in the response")
                return

            funds_data = funds_response['data']
            
            # If the response is a dictionary of scalar values, convert it to a list of dictionaries
            if isinstance(funds_data, dict):
                funds_data = [funds_data]
            
            # Load the Excel workbook and the specific sheet for funds
            wb = xw.Book("excelbast.xlsx")
            sheet = wb.sheets['Funds']  # Assuming the sheet name is 'Funds'

            # Convert funds data to a DataFrame
            df_funds = pd.DataFrame(funds_data)

            # Write the DataFrame to the Excel sheet
            sheet.range("A1").options(index=False).value = df_funds
            time.sleep(1.5)
            # print("Funds updated successfully")
        except Exception as e:
            # print(f"An error occurred while fetching or updating funds: {e}")
            time.sleep(2)  # Wait for 2 seconds before continuing



update_thread_holdings = threading.Thread(target=update_holdings_sheet)

# Start the thread
update_thread_holdings.start()

update_thread_funds = threading.Thread(target=fetch_and_update_funds)

# Start the thread
update_thread_funds.start()

update_thread_order_book = threading.Thread(target=update_order_book_sheet)

# Start the thread
update_thread_order_book.start()

update_thread_position = threading.Thread(target=update_position_sheet)

# Start the thread
update_thread_position.start()

# -----------------------------



global wb
global previous_data

# Function to continuously update time in Excel
def update_time_in_excel():
    global wb
    # Open or create Excel file
    file_name = "excelbast.xlsx"
    file_path = os.path.join(os.getcwd(), file_name)
    if os.path.exists(file_path):
        wb = xw.Book(file_path)
    else:
        wb = xw.Book()
        wb.save(file_name)
    
    # Access active sheet
    sheet = wb.sheets['Sheet1']
    
    # Continuously update time in Excel
    while True:
        current_time = datetime.now().strftime("%H:%M:%S")
        sheet.range("B1").value = "Current Time"
        sheet.range("C1").value = current_time
        time.sleep(1)  # Update every second


def find_token(symbol, json_data):

    for item in json_data:
        if item['symbol'] == symbol:
            return item.get('token')
    return None

def find_token_from_tick_data_dict(symbol, ticks_data_dict):
    for symbol_key, tick_data in ticks_data_dict.items():
        if symbol_key == symbol:
            return tick_data.get('token')
    return None

def find_exchange(symbol, json_data):
    for item in json_data:
        if item['symbol'] == symbol:
            if item.get('exch_seg') == "NFO":
                return 2
            else:
                return 1
    return None



def find_column_index(sheet, heading):
    """Find the column index where the specified heading is located."""
    headings_row = sheet.range('1:1').value  # Assuming headings are in the first row
    try:
        return headings_row.index(heading) + 1  # Add 1 because xlwings uses 1-based indexing
    except ValueError:
        return None  # Heading not found

order_counter = 0
last_order_time = time.time()

def place_single_order(symbol, token, qty, buy_or_sell, row_idx):
    global order_lock, order_counter, last_order_time

    wb = xw.Book("excelbast.xlsx")
    sheet = wb.sheets['Live']  # Assuming the sheet name is 'Live'
    
    with order_lock:
        current_time = time.time()
        
        # Reset order counter if a second has passed
        if current_time - last_order_time >= 1:
            order_counter = 0
            last_order_time = current_time
        
        # Wait if order limit exceeds 19 orders per second
        while order_counter >= 19:
            time.sleep(1)  # Sleep for a short time to reduce CPU usage
            current_time = time.time()
            if current_time - last_order_time >= 1:
                order_counter = 0
                last_order_time = current_time
        
        order_counter += 1

        global json_data
        exc_id = ""

        for symbol_key, tick_data in tick_data_dict.items():
            if symbol_key == symbol:
                exc_id = tick_data.get('exchange_type')
        
        exc_id_words = ""
        prod_type = ""
        if exc_id == 1:
            exc_id_words = "NSE"
            prod_type = "DELIVERY"
        elif exc_id == 2:
            exc_id_words = "NFO"
            prod_type = "CARRYFORWARD"

        buy_or_sell_in_words = ""
        if buy_or_sell == 1:
            buy_or_sell_in_words = "BUY"
        elif buy_or_sell == 2:
            buy_or_sell_in_words = "SELL"

        orderparams = {
            "variety": "NORMAL",
            "tradingsymbol": symbol,
            "symboltoken": token,
            "transactiontype": buy_or_sell_in_words,
            "exchange": exc_id_words,
            "ordertype": "MARKET",
            "producttype": prod_type,
            "duration": "DAY",
            "price": 0,
            'squareoff': 0,
            'stoploss': 0,
            "quantity": qty
        }
        print("**********8")
        print(orderparams)
        print("**********8")
        orderId = obj.placeOrder(orderparams)
        print(orderId)

        sheet.range(f"N{row_idx}").value = orderId


def orderAngel(symbol, token, qty, buy_or_sell, row_idx):

    symbol_limits = {
        "BANKNIFTY": 900,
        "NIFTY": 1800,
        "FINNIFTY": 1800,
        "MIDCPNIFTY": 4200
    }

    # Acquire the lock before executing orderAngel
    print("Executing orderAngel for symbol:", symbol)
    print("in orderangel")
    symbol_limits_criteria = False
    # Check if the symbol has a limit and if the quantity is within the limit
    for limit_key in symbol_limits:
        if symbol.startswith(limit_key):
            symbol_limits_criteria = True
            if qty > symbol_limits[limit_key]:
                print(f"Order quantity {qty} exceeds the limit for {symbol}. Limit is {symbol_limits[limit_key]}.")
                split_qty = int(int(qty)/int(symbol_limits[limit_key]))
                print("split_qty", split_qty)
                print("symbol_limits[limit_key]", symbol_limits[limit_key])
                for _ in range(split_qty):
                    place_single_order(symbol, token, symbol_limits[limit_key], buy_or_sell, row_idx)
            else:
                print("Placing a single order DETECTING OPTIMUM QTY")
                place_single_order(symbol, token, qty, buy_or_sell, row_idx)

    if symbol_limits_criteria == False:
        print("Placing a single order NO DETECTED SYMBOL")
        place_single_order(symbol, token, qty, buy_or_sell, row_idx)



def update_order_status(orderDetails, orderId, row_idx):
    wb = xw.Book("excelbast.xlsx")
    sheet = wb.sheets['Live']  # Assuming the sheet name is 'Live'
    for order in orderDetails['data']:
        if order['orderid'] == orderId:
            print(order)
            sheet.range(f"L{row_idx}").value = order['status']
            sheet.range(f"M{row_idx}").value = order['averageprice']
            sheet.range(f"N{row_idx}").value = orderId

order_lock = threading.Semaphore(19)


def placeOrder(row_idx, row_data):
    """Function to place an order."""
    global pending_order_dict
    global json_data
    
    print("Placing order for row:", row_idx)

    # entry_limit_price_column_index = find_column_index(sheet, 'Entry Limit price')
    entry_limit_price_column_index = 7
    tokenForSymbol = ""
    if row_data[8] is None or row_data[9] is None:
        print("Place order directly")
        for symbol_key, tick_data in tick_data_dict.items():
            if symbol_key == row_data[0]:
                tokenForSymbol =  tick_data.get('token')
        
        # Create a thread for orderAngel
        order_thread = threading.Thread(target=orderAngel, args=(row_data[0], tokenForSymbol, row_data[9], row_data[10], row_idx))
        order_thread.start()

        # Wait for the thread to complete
        order_thread.join()
        # orderAngel(row_data[0], tokenForSymbol, row_data[9], row_data[10], row_idx)
    else:
        pending_order_dict[row_idx] = row_data
        wb = xw.Book("excelbast.xlsx")
        sheet = wb.sheets['Live']  # Assuming the sheet name is 'Live'

        sheet.range(f"L{row_idx}").value = 'pending'

        print("Order stored in pending_order_dict", pending_order_dict)



# def watch_qty_and_buysell_changes():
#     # Open the Excel workbook
#     wb = xw.Book("excelbast.xlsx")
#     sheet = wb.sheets['Live']  # Assuming the sheet name is 'Live'

#     # Find the column indices for 'Qty', 'Buy/Sell', and 'Symbol'
#     qty_column_index = 10
#     buysell_column_index = 11
#     symbol_column_index = 1

#     if qty_column_index is None or buysell_column_index is None or symbol_column_index is None:
#         print("Column headings 'Qty', 'Buy/Sell', or 'Symbol' not found.")
#         return

#     # Keep track of previous values to detect updates
#     previous_data = {}

#     while True:
#         try:
#             # print(previous_data)
#             # Read the data from columns A to O
#             data_range = sheet.range('A2:P500').value

#             # Find the last row with data
#             last_row = sheet.range((sheet.cells.last_cell.row, qty_column_index)).end('up').row

#             # Read the Qty, Buy/Sell, and Symbol columns
#             qty_column = sheet.range((2, qty_column_index), (last_row, qty_column_index)).value
#             buysell_column = sheet.range((2, buysell_column_index), (last_row, buysell_column_index)).value
#             symbol_column = sheet.range((2, symbol_column_index), (last_row, symbol_column_index)).value

#             # Check for updates
#             for row_idx, (qty, buy_sell, symbol) in enumerate(zip(qty_column, buysell_column, symbol_column), start=2):
#                 if (qty is not None or buy_sell is not None) and (row_idx, 'Qty') in previous_data and (row_idx, 'Buy/Sell') in previous_data:
#                     # Check if both Qty and Buy/Sell have been updated since the last check
#                     if qty is not None and qty != "Qty" and buy_sell is not None  and buy_sell != "Buy/Sell" and symbol is not None:
#                         # Check if data has changed compared to previous data
#                         if (row_idx, 'Qty') in previous_data and (row_idx, 'Buy/Sell') in previous_data:
#                             prev_qty = previous_data[(row_idx, 'Qty')]
#                             prev_buy_sell = previous_data[(row_idx, 'Buy/Sell')]
#                             if qty != prev_qty or buy_sell != prev_buy_sell:
#                                 # Pass row data from columns A to O to placeOrder() function
#                                 data_range = sheet.range('A2:P500').value

#                                 row_data = data_range[row_idx - 2]  # Adjusting index to match data range
#                                 order_thread = threading.Thread(target=placeOrder, args=(row_idx, row_data))
#                                 # Start the thread
#                                 order_thread.start()

#                 else:
#                     if qty is not None and qty != "Qty" and buy_sell is not None and buy_sell != "Buy/Sell" and symbol is not None:
#                         data_range = sheet.range('A2:P500').value
#                         row_data = data_range[row_idx - 2]  # Adjusting index to match data range
#                         order_thread = threading.Thread(target=placeOrder, args=(row_idx, row_data))

#                         # Start the thread
#                         order_thread.start()

#                 # Update previous_data to reflect the latest values
#                 previous_data[(row_idx, 'Qty')] = qty
#                 previous_data[(row_idx, 'Buy/Sell')] = buy_sell

#         except Exception as e:
#             print("An error occurred:", e)
#             print("Retrying in 5 seconds...")
#             time.sleep(1)
#             continue

#         # Wait for some time before checking again
#         # Adjust the sleep duration based on your requirements
#         time.sleep(0.1)



def watch_qty_and_buysell_changes():
    # Open the Excel workbook
    wb = xw.Book("excelbast.xlsx")
    sheet = wb.sheets['Live']  # Assuming the sheet name is 'Live'

    # Find the column indices for 'Qty', 'Buy/Sell', and 'Symbol'
    qty_column_index = 10
    buysell_column_index = 11
    symbol_column_index = 1

    if qty_column_index is None or buysell_column_index is None or symbol_column_index is None:
        print("Column headings 'Qty', 'Buy/Sell', or 'Symbol' not found.")
        return

    # Keep track of previous values to detect updates
    previous_data = {}

    while True:
        try:
            # Read the data from columns A to O, rows 2 to 500
            data_range = sheet.range('A2:P500').value

            # Read the Qty, Buy/Sell, and Symbol columns
            qty_column = sheet.range((2, qty_column_index), (500, qty_column_index)).value
            buysell_column = sheet.range((2, buysell_column_index), (500, buysell_column_index)).value
            symbol_column = sheet.range((2, symbol_column_index), (500, symbol_column_index)).value

            # Check for updates
            for row_idx, (qty, buy_sell, symbol) in enumerate(zip(qty_column, buysell_column, symbol_column), start=2):
                if (qty is not None or buy_sell is not None) and (row_idx, 'Qty') in previous_data and (row_idx, 'Buy/Sell') in previous_data:
                    # Check if both Qty and Buy/Sell have been updated since the last check
                    if qty is not None and qty != "Qty" and buy_sell is not None and buy_sell != "Buy/Sell" and symbol is not None:
                        # Check if data has changed compared to previous data
                        if (row_idx, 'Qty') in previous_data and (row_idx, 'Buy/Sell') in previous_data:
                            prev_qty = previous_data[(row_idx, 'Qty')]
                            prev_buy_sell = previous_data[(row_idx, 'Buy/Sell')]
                            if qty != prev_qty or buy_sell != prev_buy_sell:
                                # Pass row data from columns A to O to placeOrder() function
                                row_data = data_range[row_idx - 2]  # Adjusting index to match data range
                                order_thread = threading.Thread(target=placeOrder, args=(row_idx, row_data))
                                # Start the thread
                                order_thread.start()

                else:
                    if qty is not None and qty != "Qty" and buy_sell is not None and buy_sell != "Buy/Sell" and symbol is not None:
                        row_data = data_range[row_idx - 2]  # Adjusting index to match data range
                        order_thread = threading.Thread(target=placeOrder, args=(row_idx, row_data))

                        # Start the thread
                        order_thread.start()

                # Update previous_data to reflect the latest values
                previous_data[(row_idx, 'Qty')] = qty
                previous_data[(row_idx, 'Buy/Sell')] = buy_sell
                # print(row_idx, previous_data[(row_idx, 'Qty')])

        except Exception as e:
            # print("An error occurred:", e)
            # print("Retrying in 5 seconds...")
            time.sleep(5)
            continue

        # Wait for some time before checking again
        time.sleep(0.1)



def watch_target_sl_manualexit_changes():
    print("target watching***************")
    # Open the Excel workbook
    wb = xw.Book("excelbast.xlsx")
    sheet = wb.sheets['Live']  # Assuming the sheet name is 'Live'

    # Keep track of previous values to detect updates
    previous_data = {}

    while True:
        try:
            # Read the data from columns O, P, Q
            data_range = sheet.range('O2:Q500').value

            # Find the last row with data
            last_row = sheet.range((sheet.cells.last_cell.row, 15)).end('up').row

            # Read the columns O, P, Q
            target_column = sheet.range((2, 15), (last_row, 15)).value
            stoploss_column = sheet.range((2, 16), (last_row, 16)).value
            manual_exit_column = sheet.range((2, 17), (last_row, 17)).value

            # Check for updates
            for row_idx, (target, stoploss, manual_exit) in enumerate(zip(target_column, stoploss_column, manual_exit_column), start=2):
                if target != previous_data.get((row_idx, 'Target')) and target != "Target" and target is not None:
                    print(f"Change detected in Target at row {row_idx}: {target}")
                    previous_data[(row_idx, 'Target')] = target
                    # Get the row data from columns A to T
                    row_data = sheet.range((row_idx, 1), (row_idx, 20)).value
                    # Save row data in a dictionary
                    if row_data[14] is not None:
                        exit_data_dict[row_idx] = row_data 

                if stoploss != previous_data.get((row_idx, 'Stoploss')) and stoploss != "SL" and stoploss is not None:
                    print(f"Change detected in Stoploss at row {row_idx}: {stoploss}")
                    previous_data[(row_idx, 'Stoploss')] = stoploss
                    # Get the row data from columns A to T
                    row_data = sheet.range((row_idx, 1), (row_idx, 20)).value
                    # Save row data in a dictionary
                    if row_data[14] is not None:
                        exit_data_dict[row_idx] = row_data 

                if manual_exit != previous_data.get((row_idx, 'Manual Exit')) and manual_exit != 'Manual Exit' and manual_exit is not None:
                    print(f"Change detected in Manual Exit at row {row_idx}: {manual_exit}")
                    previous_data[(row_idx, 'Manual Exit')] = manual_exit
                    # Get the row data from columns A to T
                    row_data = sheet.range((row_idx, 1), (row_idx, 20)).value
                    # Save row data in a dictionary
                    if row_data[14] is not None:
                        exit_data_dict[row_idx] = row_data 

            # Wait for some time before checking again
            time.sleep(0.1)
        except Exception as e:
            # print("An error occurred:", e)
            # print("Retrying in 5 seconds...")
            time.sleep(1)
            continue

    
# Function to continuously read and print text cells from Excel
def read_excel_and_print():
    global json_data

    global wb
    global previous_data_token
    previous_data_token = None
    # Continuously read and print text cells from Excel
    while True:
        try:
            print("reading")
            # Open or create Excel file
            file_name = "excelbast.xlsx"
            file_path = os.path.join(os.getcwd(), file_name)
            if os.path.exists(file_path):
                wb = xw.Book(file_path)
            else:
                wb = xw.Book()
                wb.save(file_name)
            
            # Access active sheet
            sheet = wb.sheets['Live']

            # Read the first 500 rows of the first column
            data = sheet.range('A2:A500').value
            
            # Print non-empty values that are different from previous data
            if previous_data_token is not None:
                for value, prev_value in zip(data, previous_data_token):
                    if isinstance(value, str) and value != prev_value:
                        token = find_token(value, json_data)
                        exc_type = find_exchange(value, json_data)
                        print(f"Text in cell: {value} and {token} and {exc_type}")
                        if token:
                            find_ticks_for_token(token, exc_type)
            else:
                for value in data:
                    if isinstance(value, str):
                        token = find_token(value, json_data)
                        exc_type = find_exchange(value, json_data)
                        print(f"Text in cell: {value} and {token} and {exc_type}")
                        if token:
                            find_ticks_for_token(token, exc_type)
            
            previous_data_token = data
            time.sleep(1)

        except Exception as e:
            # print("An error occurred:", e)
            # print("Retrying in 5 seconds...")
            time.sleep(1)
            continue





print("checkpoint 0")

# Create threads for updating time and reading Excel
token_thread = threading.Thread(target=create_token_sheet)
token_thread.start()
token_thread.join()  # Wait for token thread to complete
print("checkpoint 1")

# Start WebSocket connection in a separate thread
websocket_thread = threading.Thread(target=start_websocket)
update_thread = threading.Thread(target=update_time_in_excel)
read_thread = threading.Thread(target=read_excel_and_print)
watch_qty_and_buysell_changes_thread = threading.Thread(target=watch_qty_and_buysell_changes)
watch_target_sl_manualexit_changes_thread = threading.Thread(target=watch_target_sl_manualexit_changes)


# Daemonize the threads so they automatically stop when the main program exits
update_thread.daemon = True
read_thread.daemon = True


update_excel_thread = threading.Thread(target=create_main_list)
update_excel_thread.start()


# Start the threads
websocket_thread.start()
update_thread.start()
read_thread.start()
watch_qty_and_buysell_changes_thread.start()
# watch_target_sl_manualexit_changes_thread.start()


# exit_orders_thread = threading.Thread(target=track_exit_orders)
# exit_orders_thread.daemon = True  # Set as daemon thread so it exits when the main thread exits
# exit_orders_thread.start()

# Main thread continues to do other tasks or just waits



# -------------------------------

# sheets
# speed for orders
# order status earse

