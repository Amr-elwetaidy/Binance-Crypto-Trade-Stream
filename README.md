# Crypto Trade Stream Script

## Overview

The Crypto Trade Stream script is a Python script designed to connect to the Binance WebSocket API and stream real-time trade data for multiple cryptocurrency symbols. The script updates an Excel sheet with the latest trade price and timestamp for the specified symbols.

## Prerequisites

Before running the script, make sure you have the following installed:

- Python 3.7 or later
- Required Python packages: see `requirements.txt`

## Setup

1. **Install Dependencies:**

    ```bash
    pip install -r requirements.txt
    ```

## Usage

Run the script by executing the following command in the terminal:

```bash
python script_name.py
```

Replace `script_name.py` with the actual name of your Python script.

## Script Structure

### 1. `format_sheet(sheet)`

Formats the headers of the Excel sheet with columns for symbol, price, and time.

### 2. `parse_time(timestamp_ms)`

Converts milliseconds since Unix epoch to a formatted datetime string.

### 3. `binance_all_trade_streams(symbols_map)`

Connects to the Binance WebSocket API, subscribes to trade streams for specified symbols, and continuously updates the Excel sheet.

### 4. `main()`
- Initializes Excel application and creates a new workbook.
- Activates the specified sheet and formats it.
- Reads the list of symbols from "symbols.txt" and creates a symbols map.
- Connects to Binance WebSocket for trade streams.
- Saves and closes the workbook, and quits the Excel application.

## Notes

- The script uses asyncio to handle asynchronous operations.
- The Binance WebSocket connection is established using the URL "wss://stream.binance.com:9443/stream."
- The script creates tasks to subscribe to trade streams for up to 1000 symbols per connection.
- Connection retries are attempted up to 5 times in case of closure.
- If an error occurs during data processing or connection, an exception is logged.

## Example

```python
symbols_map = {"BTCUSDT": 1, "ETHUSDT": 2, "BNBUSDT": 3}

await binance_all_trade_streams(symbols_map)
```

## Troubleshooting

- If you encounter issues, ensure that you have the required dependencies installed as specified in `requirements.txt`.

- If the script is interrupted (e.g., by a KeyboardInterrupt), any in-memory data structure may be lost. Ensure you restart the script to resume trade information display.