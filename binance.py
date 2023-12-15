from win32com.client.gencache import EnsureDispatch
from win32com.client import constants
import pythoncom

import websockets
import datetime
import asyncio
import json
import os


def format_sheet(sheet):
    """
    Format the sheet headers
    """
    for cell in [("A1", "Symbol"), ("B1", "Price"), ("C1", "Time")]:
        current_cell = sheet.Range(cell[0])
        current_cell.Value = cell[1]
        current_cell.Font.Bold = True
        current_cell.Borders.LineStyle = 1
        current_cell.HorizontalAlignment = constants.xlCenter
        sheet.Columns(cell[0][0]).ColumnWidth = 25


def parse_time(timestamp_ms):
    """
    Convert milliseconds since Unix epoch to a datetime object
    """
    timestamp_sec = timestamp_ms / 1000.0

    # Convert to a datetime object
    trade_time = datetime.datetime.fromtimestamp(timestamp_sec, datetime.UTC).strftime("%d-%m-%Y %H:%M:%S.%f")[:-3]

    return trade_time


async def binance_all_trade_streams(symbols_map, sheet):
    """
    Connects to the Binance WebSocket API to stream real-time trade data for multiple symbols.

    Parameters:
    - symbols_map (dict): A dictionary mapping symbol names to corresponding row indices in the Excel sheet.
    - symbols (list): A list of symbols for which trade data will be streamed.
    - sheet: The Excel sheet where the trade data will be updated.

    This function subscribes to the Binance WebSocket for trade streams and continuously updates an Excel sheet
    with the latest trade price and timestamp for the specified symbols.

    Args:
    - symbols_map (dict): A dictionary where keys are symbol names, and values are corresponding row indices in the sheet.
    - symbols (list): A list of symbol names for which trade data will be streamed.
    - sheet: The Excel sheet where trade data will be updated.

    Returns:
    None

    Note:
    - The function uses asyncio to handle asynchronous operations.
    - The Binance WebSocket connection is established using the URL "wss://stream.binance.com:9443/stream".
    - The function creates tasks to subscribe to the trade streams for up to 1000 symbols per connection.
    - Connection retries are attempted up to 5 times in case of closure.
    - If an error occurs during data processing or connection, an exception is logged.

    Example:
    ```python
    symbols_map = {"BTCUSDT": 1, "ETHUSDT": 2, "BNBUSDT": 3}
    symbols = ["BTCUSDT", "ETHUSDT", "BNBUSDT"]
    sheet = get_excel_sheet()  # Replace with actual function to get the Excel sheet

    await binance_all_trade_streams(symbols_map, symbols, sheet)
    ```
    """
    params = []
    for symbol in symbols_map:
        params.append(f"{symbol.lower()}@trade")

    async def subscribe(symbols_map, sheet, params):
        t = 0
        while True:
            try:
                # Connect to the websocket
                async with websockets.connect(f"wss://stream.binance.com:9443/stream") as websocket:
                    t = 0

                    # Subscribe to the streams
                    subscription_message = {
                        "method": "SUBSCRIBE",
                        "params": params,
                        "id": 1
                    }
                    await websocket.send(json.dumps(subscription_message))

                    # Receive data and update the sheet
                    while True:
                        response = await websocket.recv()

                        try:
                            data = json.loads(response)["data"]
                            row = symbols_map[data["s"]]
                            sheet.Range(f"B{row}").Value = data["p"]
                            sheet.Range(f"C{row}").Value = parse_time(data["T"])
                        except pythoncom.com_error:
                            print("Error updating sheet.", end="\n\n")
                            break
                        except:
                            pass
                break

            except websockets.ConnectionClosed as e:
                print("-"*40, end="\n\n")
                print(f"Connection closed: {e}", end="\n\n")
                t += 1
                if t == 5:
                    print("Maximum retries exceeded.", end="\n\n")
                    print("-"*40, end="\n\n")
                    return
                print("Reconnecting...", end="\n\n")
                print("-"*40, end="\n\n")
                await asyncio.sleep(5)
            except Exception as e:
                print("-"*40, end="\n\n")
                print(f"Error: {e}", end="\n\n")
                print("-"*40, end="\n\n")
                return

    # Create a task for connection (up to 1000 symbols per connection)
    tasks = [subscribe(symbols_map, sheet, params[chunk:chunk+100]) for chunk in range(0, len(params), 100)]

    # Wait for all tasks to complete
    await asyncio.gather(*tasks)

def main():
    global MESSAGE
    MESSAGE = "Success"

    try:
        # Create a new Excel application
        excel_app = EnsureDispatch("Excel.Application")

        # make Excel visible
        excel_app.Visible = True

        # Open the workbook
        file_path = os.path.join(os.getcwd(), "Crypto.xlsx")
        workbook = excel_app.Workbooks.Add()
        excel_app.DisplayAlerts = False
        workbook.SaveAs(file_path)
        excel_app.DisplayAlerts = True
    except:
        print("Error creating Excel application.", end="\n\n")
        return

    # Activate the specified sheet
    try:
        sheet = workbook.Sheets("Sheet1")
        sheet.Activate()
        format_sheet(sheet)
    except:
        print("Sheet Format Error.", end="\n\n")
        MESSAGE = "Error"
        return

    try:
        with open("symbols.txt") as file:
            symbol_list = file.read().splitlines()
    except:
        print("Error reading symbols file.", end="\n\n")
        return

    # Create a dictionary to store the row number for each symbol
    symbols_map = {}
    row = 2
    try:
        for symbol in symbol_list:
            sheet.Range(f"A{row}").Value = symbol.upper()
            # sheet.Range(f"C{row}").NumberFormat = "dd-mm-yyyy hh:mm:ss.000"
            symbols_map[symbol.upper()] = row
            row += 1
        del symbol_list
    except:
        print("Error creating symbols map.", end="\n\n")
        return

    # Run the event loop to connect to the Binance WebSocket for trade streams
    try:
        asyncio.run(binance_all_trade_streams(symbols_map, sheet))
    except KeyboardInterrupt:
        pass

    # Save and close the workbook
    try:
        workbook.Save()
        workbook.Close()
    except:
        pass

    # Quit the Excel application
    try:
        excel_app.Quit()
    except:
        pass


if __name__ == "__main__":
    print(f"--> Started at {datetime.datetime.now().time().strftime("%I:%M:%S %p")} <--", end="\n\n")
    main()
    print(f"--> Finished at {datetime.datetime.now().time().strftime("%I:%M:%S %p")} <--", end="\n\n")
