import time , math
import pandas as pd
import pyRofex
import xlwings as xw


wb = xw.Book('..\\epgb_pyRofex.xlsx')
shtTickers = wb.sheets('pyRofex')
shtData = wb.sheets('HomeBroker')
#shtOperaciones = wb.sheets('Posiciones')

pyRofex._set_environment_parameter(
    "url", "https://api.eco.xoms.com.ar/", pyRofex.Environment.LIVE)
pyRofex._set_environment_parameter(
    "ws", "wss://api.eco.xoms.com.ar/", pyRofex.Environment.LIVE)
pyRofex.initialize(user="20263866623",
                   password="APP883dR#",
                   account="62226",
                   environment=pyRofex.Environment.LIVE)

rng = shtTickers.range('A2:C35').expand() # OPCIONES
tickers = pd.DataFrame(rng.value, columns=['ticker', 'symbol', 'strike'])

rng = shtTickers.range('E2:F5').expand() # ACCIONES
#tickers = tickers.append(pd.DataFrame(rng.value, columns=['ticker', 'symbol'])) # Metodo viejo APPEND
tickers = pd.concat([tickers, pd.DataFrame(rng.value, columns=['ticker', 'symbol'])])

rng = shtTickers.range('H2:I70').expand() # BONOS
tickers = pd.concat([tickers, pd.DataFrame(rng.value, columns=['ticker', 'symbol'])])

rng = shtTickers.range('K2:L15').expand() # LETRAS
tickers = pd.concat([tickers, pd.DataFrame(rng.value, columns=['ticker', 'symbol'])])

rng = shtTickers.range('N2:O30').expand() # ONs
tickers = pd.concat([tickers, pd.DataFrame(rng.value, columns=['ticker', 'symbol'])])

rng = shtTickers.range('Q2:R40').expand() # CEDEARS
tickers = pd.concat([tickers, pd.DataFrame(rng.value, columns=['ticker', 'symbol'])])

rng = shtTickers.range('T2:U20').expand() # CAUCION
tickers = pd.concat([tickers, pd.DataFrame(rng.value, columns=['ticker', 'symbol'])])



instruments_2 = pyRofex.get_detailed_instruments()
data = pd.DataFrame(instruments_2['instruments'])
df = pd.DataFrame.from_dict(dict(data['instrumentId']), orient='index')
df = df['symbol'].to_list()
tickers['remove'] = tickers['ticker'].isin(df).astype(int)
tickers = tickers[tickers['remove'] !=0]
instruments = tickers['ticker'].to_list()


df_datos = pd.DataFrame({'ticker': tickers['ticker'].to_list(),'symbol': tickers['symbol'].to_list()}, columns=[
    'ticker', 'symbol', 'bidsize', 'bid', 'ask', 'asksize', 'last', 'close','open', 'high', 'low', 'volume','lastupdate','nominal','trade'])
df_datos = df_datos.set_index('ticker')

thisData = pd.DataFrame(columns=['ticker','symbol', 'bidsize', 'bid', 'ask', 'asksize', 'last', 'close','open', 'high', 'low', 'volume', 'lastupdate','nominal','trade'])


def addTick(symbol, bidSize, bid, ask, askSize, last, close, open, high, low, volume, lastUpdate, nominal, trade):
    global thisData, bonos
    thisData = pd.DataFrame([{
        'ticker': symbol, 'bidsize': bidSize, 'bid': bid, 'ask': ask, 'asksize': askSize, 'last': last,
        'close':close, 'open': open, 'high': high, 'low': low, 'volume': volume,
        'lastupdate': time.strftime('%m/%d/%Y %H:%M:%S', time.gmtime(lastUpdate / 1000.)), 'nominal': nominal, 'trade': trade}])
    thisData = thisData.set_index('ticker')
    df_datos.update(thisData)

def market_data_handler(message):
    symbol = message['instrumentId']['symbol']
    last = None if not message['marketData']['LA'] else message['marketData']['LA']['price']
    lastUpdate = None if not message['marketData']['LA'] else message['marketData']['LA']['date']
    bid = None if not message['marketData']['BI'] else message['marketData']['BI'][0]['price']
    bidSize = None if not message['marketData']['BI'] else message['marketData']['BI'][0]['size']
    ask = None if not message['marketData']['OF'] else message['marketData']['OF'][0]['price']
    askSize = None if not message['marketData']['OF'] else message['marketData']['OF'][0]['size']
    close = None if not message['marketData']['CL'] else message['marketData']['CL']['price']
    open = None if not message['marketData']['OP'] else message['marketData']['OP']
    high = None if not message['marketData']['HI'] else message['marketData']['HI']
    low = None if not message['marketData']['LO'] else message['marketData']['LO']
    volume = None if not message['marketData']['EV'] else message['marketData']['EV']
    nominal = None if not message['marketData']['NV'] else message['marketData']['NV']
    trade = None if not message['marketData']['TV'] else message['marketData']['TV']

    addTick(symbol, bidSize, bid, ask, askSize, last, close, open, high, low, volume, lastUpdate, nominal, trade)

def error_handler(message):
    print("Error Message Received: {0}".format(message))

def exception_handler(e):
    print("Exception Occurred: {0}".format(e.message))


df_order = pd.DataFrame()

def order_report_handler(message):
    global operaciones
    print(message)
    if message['orderReport']['status'] == "NEW":
        orderId = message['orderReport']['clOrdId']
    else:
        orderId = message['orderReport']['origClOrdId']
    print(orderId)
    symbol = message['orderReport']['instrumentId']['symbol']
    side = message['orderReport']['side']
    price = message['orderReport']['price']
    qty = message['orderReport']['orderQty']
    status = message['orderReport']['status']
    cumQty = message['orderReport']['cumQty']
    leavesQty = message['orderReport']['leavesQty']
    avgPx = message['orderReport']['avgPx']
    print(orderId, symbol, side, price, qty, status, cumQty, leavesQty, avgPx)
    thisOp = pd.DataFrame([{
        'orderId': orderId, 'ticker': symbol, 'Tipo':side, 'Precio':price, 'Cant':qty, 'Status': status, 'Cant Acum':cumQty,
        'Cant Rest':leavesQty, 'Px Prom':avgPx}],
        columns=['orderId', 'ticker', 'Tipo', 'Precio', 'Cant', 'Status', 'Cant Acum', 'Cant Rest', 'Px Prom'])
    thisOp = thisOp.set_index('orderId')
    print('ThisOp: ',thisOp)
    if status != "NEW":
        print('Orden actualizada: ', symbol)
        operaciones.update(thisOp)
    else:
        print('Nueva orden: ', symbol)
        operaciones = operaciones.append(thisOp)
    print(operaciones)

def order_error_handler(message):
    print("Error Message Received: {0}".format(message))

def order_exception_handler(e):
    print("Exception Occurred: {0}".format(e.message))

pyRofex.init_websocket_connection(market_data_handler=market_data_handler,
                                  error_handler=error_handler,
                                  exception_handler=exception_handler,
                                  order_report_handler=order_report_handler)

entries = [pyRofex.MarketDataEntry.BIDS,
           pyRofex.MarketDataEntry.OFFERS,
           pyRofex.MarketDataEntry.LAST,
           pyRofex.MarketDataEntry.OPENING_PRICE,
           pyRofex.MarketDataEntry.CLOSING_PRICE,
           pyRofex.MarketDataEntry.HIGH_PRICE,
           pyRofex.MarketDataEntry.LOW_PRICE,
           pyRofex.MarketDataEntry.TRADE_VOLUME,
           pyRofex.MarketDataEntry.NOMINAL_VOLUME,
           pyRofex.MarketDataEntry.TRADE_EFFECTIVE_VOLUME]

pyRofex.market_data_subscription(tickers=instruments, entries=entries, depth=1)
# pyRofex.order_report_subscription()

# shtOperaciones.range('C4:K500').value = ""
shtData.range('A1:L1000').value = ""


while True:
    try:
        # loop = asyncio.get_event_loop()
        # update = updateSheet()
        # loop.run_until_complete(update)

        shtData.range('A1').options(index=False, headers=True).value = df_datos
        # shtOperaciones.range('C4').options(index=False, headers=False).value = operaciones
        time.sleep(1)
        print(("online"), time.strftime("%H:%M:%S"))
        if time.strftime("%H:%M:%S") > '17:10:00':
            print('Salida por cierre del mercado')
            break

    except:
        print('Hubo un error al actualizar excel')


