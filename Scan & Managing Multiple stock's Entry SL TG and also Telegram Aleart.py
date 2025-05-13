import pdb
import time
import datetime
import traceback
from Dhan_Tradehull import Tradehull
import pandas as pd
from pprint import pprint
import talib
import pandas_ta as ta
import xlwings as xw
import winsound

# Client credentials
client_code = "1101067511"
token_id    = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJkaGFuIiwicGFydG5lcklkIjoiIiwiZXhwIjoxNzQ5NTM4NDkwLCJ0b2tlbkNvbnN1bWVyVHlwZSI6IlNFTEYiLCJ3ZWJob29rVXJsIjoiIiwiZGhhbkNsaWVudElkIjoiMTEwMTA2NzUxMSJ9.nDrKCAtxD4fOCnwRE1rdnLAq-QBC1XWU_3T6sy_6rq_rHq4M9nx7b5F9v_VGx-AmVIWbMSQUYMvzr7Nx5xCeEQ"

# Initialize Tradehull object with client credentials
tsl = Tradehull(client_code, token_id)

# Watchlist of stocks to monitor
watchlist = ['BEL', 'TCS', 'TECHM']

# Template for a single order
single_order = {'name': None, 'date': None, 'entry_time': None, 'entry_price': None, 'buy_sell': None, 'qty': None, 'sl': None, 'exit_time': None, 'exit_price': None, 'pnl': None, 'remark': None, 'traded': None, 'alert_sent': False}

# Create an orderbook to store details of trades
orderbook = {}

# Connect to Excel for live trade and completed orders tracking
wb = xw.Book('Live Trade.xlsx')
live_Trading = wb.sheets['Live_Trading']
completed_orders_sheet = wb.sheets['completed_orders']

# Reentry condition to control whether to re-enter a trade after SL/TG hit
reentry = "yes"  # "yes/no"
completed_orders = []

# Telegram bot configuration for alerts
bot_token = "8019906856:AAFaYwtLkYSnHsZqoNpFpixuWo8WyA-Z3dM"
receiver_chat_id = "7740275527"

# Clear previous data from the sheets in Excel
live_Trading.range("A2:Z100").value = None
completed_orders_sheet.range("A2:Z100").value = None

# Initialize orderbook for all watchlist stocks
for name in watchlist:
    orderbook[name] = single_order.copy()

# Main trading loop
while True:
    print("starting while Loop \n\n")

    # Check if market is open (between 9:00 AM and 3:55 PM)
    current_time = datetime.datetime.now().time()
    if current_time < datetime.time(13, 55):
        print(f"Wait for market to start", current_time)
        time.sleep(1)
        continue

    # If market time is over, cancel all orders and break the loop
    if current_time > datetime.time(15, 20):
        order_details = tsl.cancel_all_orders()
        print(f"Market over Closing all trades !! Bye Bye See you Tomorrow", current_time)
        pdb.set_trace()
        break

    # Get live prices for all watchlist stocks
    all_ltp = tsl.get_ltp_data(names=watchlist)

    # Loop through each stock in the watchlist
    for name in watchlist:
        orderbook_df = pd.DataFrame(orderbook).T
        live_Trading.range('A1').value = orderbook_df

        completed_orders_df = pd.DataFrame(completed_orders)
        completed_orders_sheet.range('A1').value = completed_orders_df

        current_time = datetime.datetime.now()
        print(f"Scanning {name} {current_time}")

        try:
            # Fetch historical data for technical analysis
            chart = tsl.get_historical_data(tradingsymbol=name, exchange='NSE', timeframe="5")
            chart['rsi'] = talib.RSI(chart['close'], timeperiod=14)
            chart['ema5'] = talib.EMA(chart['close'], timeperiod=5)
            chart['ema10'] = talib.EMA(chart['close'], timeperiod=10)

            cc = chart.iloc[-2]

            # Conditions for buy entry
            bc1 = cc['rsi'] > 45
            bc2 = cc['ema5'] > cc['ema10']
            bc3 = orderbook[name]['traded'] is None and not orderbook[name]['alert_sent']
        except Exception as e:
            print(e)
            continue

        # If all conditions are met, generate a buy signal
        if bc1 and bc2 and bc3:
            print("buy ", name, "\t")
            alert_message = f"🟢 BUY SIGNAL: {name}\n\nRSI: {round(cc['rsi'], 2)}\nEMA5: {round(cc['ema5'], 2)}\nEMA10: {round(cc['ema10'], 2)}\nLTP: {round(all_ltp[name], 2)}\n\n📍Time: {datetime.datetime.now().strftime('%H:%M:%S')}"
            tsl.send_telegram_alert(message=alert_message, receiver_chat_id=receiver_chat_id, bot_token=bot_token)

            # Mark that an alert has been sent for this stock
            orderbook[name]['alert_sent'] = True

            # Check margin availability for order placement
            margin_avialable = tsl.get_balance()
            margin_required = cc['close'] / 4.5

            if margin_avialable < margin_required:
                print(f"Less margin, not taking order : margin_avialable is {margin_avialable} and margin_required is {margin_required} for {name}")
                continue

            # Update orderbook with trade details
            orderbook[name]['name'] = name
            orderbook[name]['date'] = str(current_time.date())
            orderbook[name]['entry_time'] = str(current_time.time())[:8]
            orderbook[name]['buy_sell'] = "BUY"
            orderbook[name]['qty'] = 1

            try:
                # Place market buy order
                entry_orderid = tsl.order_placement(tradingsymbol=name, exchange='NSE', quantity=orderbook[name]['qty'], price=0, trigger_price=0, order_type='MARKET', transaction_type='BUY', trade_type='MIS')
                orderbook[name]['entry_orderid'] = entry_orderid
                orderbook[name]['entry_price'] = tsl.get_executed_price(orderid=orderbook[name]['entry_orderid'])

                # Set target and stop loss levels based on entry price
                orderbook[name]['tg'] = round(orderbook[name]['entry_price'] * 1.002, 1)   # Target: 1% above entry price
                orderbook[name]['sl'] = round(orderbook[name]['entry_price'] * 0.998, 1)    # Stop loss: 0.2% below entry price

                # Place stop-loss order
                sl_orderid = tsl.order_placement(tradingsymbol=name, exchange='NSE', quantity=orderbook[name]['qty'], price=0, trigger_price=orderbook[name]['sl'], order_type='STOPMARKET', transaction_type='SELL', trade_type='MIS')
                orderbook[name]['sl_orderid'] = sl_orderid
                orderbook[name]['traded'] = "yes"

                # Send Telegram alert with order details
                message = "\n".join(f"'{key}': {repr(value)}" for key, value in orderbook[name].items())
                message = f"Entry_done {name} \n\n {message}"
                tsl.send_telegram_alert(message=message, receiver_chat_id=receiver_chat_id, bot_token=bot_token)

            except Exception as e:
                print(e)
                pdb.set_trace(header="Error in entry order")

        # If the order is already traded, monitor for exit conditions (SL or TG hit)
        if orderbook[name]['traded'] == "yes":
            bought = orderbook[name]['buy_sell'] == "BUY"

            if bought:
                try:
                    # Get live price for monitoring SL/TG conditions
                    ltp = all_ltp[name]
                    sl_hit = tsl.get_order_status(orderid=orderbook[name]['sl_orderid']) == "TRADED"
                    tg_hit = ltp > orderbook[name]['tg']
                except Exception as e:
                    print(e)
                    pdb.set_trace(header="Error in SL order checking")

                # If stop loss is hit, exit the position and record the PnL
                if sl_hit:
                    try:
                        orderbook[name]['exit_time'] = str(current_time.time())[:8]
                        orderbook[name]['exit_price'] = tsl.get_executed_price(orderid=orderbook[name]['sl_orderid'])
                        orderbook[name]['pnl'] = round((orderbook[name]['exit_price'] - orderbook[name]['entry_price']) * orderbook[name]['qty'], 1)
                        orderbook[name]['remark'] = "Bought_SL_hit"

                        # Send alert for SL hit
                        message = "\n".join(f"'{key}': {repr(value)}" for key, value in orderbook[name].items())
                        message = f"SL_HIT {name} \n\n {message}"
                        tsl.send_telegram_alert(message=message, receiver_chat_id=receiver_chat_id, bot_token=bot_token)

                        # Reset alert_sent flag after SL hit
                        orderbook[name]['alert_sent'] = False

                        # Optionally, re-enter the trade if reentry is enabled
                        if reentry == "yes":
                            completed_orders.append(orderbook[name])
                            orderbook[name] = None
                    except Exception as e:
                        print(e)
                        pdb.set_trace(header="Error in SL hit")

                # If target is hit, square off the position and record the PnL
                if tg_hit:
                    try:
                        tsl.cancel_order(OrderID=orderbook[name]['sl_orderid'])
                        time.sleep(2)

                        # Place market sell order to square off the position
                        square_off_buy_order = tsl.order_placement(tradingsymbol=orderbook[name]['name'], exchange='NSE', quantity=orderbook[name]['qty'], price=0, trigger_price=0, order_type='MARKET', transaction_type='SELL', trade_type='MIS')

                        orderbook[name]['exit_time'] = str(current_time.time())[:8]
                        orderbook[name]['exit_price'] = tsl.get_executed_price(orderid=square_off_buy_order)
                        orderbook[name]['pnl'] = (orderbook[name]['exit_price'] - orderbook[name]['entry_price']) * orderbook[name]['qty']
                        orderbook[name]['remark'] = "Bought_TG_hit"

                        # Send alert for TG hit
                        message = "\n".join(f"'{key}': {repr(value)}" for key, value in orderbook[name].items())
                        message = f"TG_HIT {name} \n\n {message}"
                        tsl.send_telegram_alert(message=message, receiver_chat_id=receiver_chat_id, bot_token=bot_token)

                        # Reset alert_sent flag after TG hit
                        orderbook[name]['alert_sent'] = False

                        # Optionally, re-enter the trade if reentry is enabled
                        if reentry == "yes":
                            completed_orders.append(orderbook[name])
                            orderbook[name] = None

                        # Beep sound to indicate target hit
                        winsound.Beep(1500, 10000)

                    except Exception as e:
                        print(e)
                        pdb.set_trace(header="Error in TG hit")
