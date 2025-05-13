# Intraday Auto Trading Bot with EMA & RSI Strategy Using Dhan API

This Python script is an automated intraday trading bot that places buy orders based on a simple yet effective technical strategy using EMA and RSI indicators. It uses the Dhan API for order execution, TA-Lib for indicator calculations, Pandas for data processing, xlwings for real-time Excel updates, and Telegram Bot API for instant trade alerts.

**Key Features:**
 Strategy:
Buy signal is generated when:
RSI(14) > 45
EMA(5) > EMA(10)

T**argets and stop-loss are placed immediately after entry.**
Automated Trading Flow:
Auto-fetch LTP and 5-minute historical data
Auto place buy orders with SL and TG
Monitor SL/TP and auto-exit when conditions are met

**Option for re-entry after SL or TG hit**
Cancels all open orders after market close

**Live Excel Integration:**
Updates ongoing and completed trades in "Live Trade.xlsx" with xlwings

**Telegram Alerts:**
Sends alerts for buy signals, SL hits, and TG hits

**Margin Check:**
Trades only if sufficient margin is available (calculated dynamically)

**Target and Stop Loss:**
Target: 0.2% above entry price
Stop Loss: 0.2% below entry price


**File Structure for GitHub:**
├── Scan & Managing Multiple stock's Entry SL TG and also Telegram Aleart.py                 # Main trading script (your current code)
├── Dhan_Tradehull.py                                                                        # API helper for Dhan (imported class)
├── Live Trade.xlsx                                                                          # Excel file for real-time tracking
├── README.md                                                                                # Project overview and instructions
├── requirements.txt                                                                         # Required libraries and versions

**Requirements:**
Python 3.8+
TA-Lib
pandas
pandas_ta
xlwings
winsound (Windows only)
Dhan API access
Telegram bot and chat ID
