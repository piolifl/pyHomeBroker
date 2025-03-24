import pyRofex
import xlwings as xw
import time , math
import pandas as pd
import yfinance as yf
import requests

wb = xw.Book('D:\\pyHomeBroker\\epgb.xlsb')
shtTickers = wb.sheets('pyRofex')
shtData = wb.sheets('MATRIZ OMS')
shtData.range('A1').value = 'symbol'
shtData.range('Q1').value = 'PRECIOS'
shtData.range('S1').value = 'ADR'
shtData.range('T1').value = 'ROLL'
shtData.range('W1').value = 'R'
shtData.range('X1').value = 'STOP'
shtData.range('Y1').value = 'VETA'
shtData.range('Z1').value = 2
rangoDesde = '26'
rangoHasta = '60'
reCompra = False
esFinde = False
    
def diaLaboral():
    global esFinde
    hoyEs = time.strftime("%A")
    if hoyEs == 'Saturday' or hoyEs == 'Sunday':
        print('FIN DE SEMANA, no se actualizan los precios locales y no se envian ordenes al broker.')
        esFinde = True

pyRofex._set_environment_parameter("url", "https://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)
pyRofex._set_environment_parameter("ws", "wss://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)
pyRofex.initialize(user="20263866623", password="Bordame01!", account="47352", environment=pyRofex.Environment.LIVE)
diaLaboral()
print(("online VETA OMS cuenta: 47352"), time.strftime("%H:%M:%S"))

def loguinHB():
    from pyhomebroker import HomeBroker  
    global hb
    try:
        hb = HomeBroker(int('284'))
        hb.auth.login(
            dni='26386662', 
            user='piolifl',  
            password='Bordame01',
            raise_exception=True)
        print(("online VETA HB  cuenta: 47352"), time.strftime("%H:%M:%S"))
    except: 
        print(("    NO se pudo loguear en VETA HOME BROKER 47352    * "), time.strftime("%H:%M:%S"))
        pass
loguinHB()

rng = shtTickers.range('A2:C2').expand() # OPCIONES
opciones = pd.DataFrame(rng.value, columns=['ticker', 'symbol', 'strike'])
rng = shtTickers.range('E2:F2').expand() # ACCIONES
acc = pd.DataFrame(rng.value, columns=['ticker', 'symbol'])
rng = shtTickers.range('H2:I2').expand() # BONOS
bonos = pd.DataFrame(rng.value, columns=['ticker', 'symbol'])
rng = shtTickers.range('K2:L2').expand() # LETRAS
letras = pd.DataFrame(rng.value, columns=['ticker', 'symbol'])
rng = shtTickers.range('N2:O2').expand() # ONS
ons = pd.DataFrame(rng.value, columns=['ticker', 'symbol'])
rng = shtTickers.range('Q2:R2').expand() # CEDARS
cedear = pd.DataFrame(rng.value, columns=['ticker', 'symbol'])
rng = shtTickers.range('T2:U2').expand() # CAUCHO
caucion = pd.DataFrame(rng.value, columns=['ticker', 'symbol'])

tickers = pd.concat([opciones, acc, bonos,cedear,ons,letras, caucion ])

listLength = len(opciones) + len(acc) + 31
allLength = 30 + len(tickers) - len(acc)  - len(caucion)
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

#operaciones = pd.DataFrame(columns=['orderId', 'ticker', 'Tipo', 'Precio', 'Cant', 'Status', 'Cant Acum', 'Cant Rest', 'Px Prom'])
#operaciones = operaciones.set_index('orderId')

def addTick(symbol, bidSize, bid, ask, askSize, last, close, open, high, low, volume, lastUpdate, nominal, trade):
    global thisData
    thisData = pd.DataFrame([{
        'ticker': symbol, 'bidsize': bidSize, 'bid': bid, 'ask': ask, 'asksize': askSize, 'last': last,
        'close':close, 'open': open, 'high': high, 'low': low, 'volume': volume,
        'lastupdate': time.strftime("%H:%M:%S"), 'nominal': nominal, 'trade': trade}])
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
                                  #order_report_handler=order_report_handler
                                  )

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
#pyRofex.order_report_subscription()

def getPortfolioHB(hb, comitente, tipo):
    try:
        shtData.range('U26:'+'U'+str(rangoHasta)).value = ''
        payload = {'comitente': str(comitente),
        'consolida': '0',
        'proceso': '22',
        'fechaDesde': None,
        'fechaHasta': None,
        'tipo': None,
        'especie': None,
        'comitenteMana': None}
        
        portfolio = requests.post("https://cuentas.vetacapital.com.ar/Consultas/GetConsulta", cookies=hb.auth.cookies, json=payload).json()

        try: 
            shtData.range('M1').value = portfolio['Result']['Activos'][0]['Subtotal'][0]['IMPO']
            print('Total:', portfolio['Result']['Activos'][0]['Subtotal'][0]['IMPO'], end='  ')
        except: pass
        try: 
            print('Disponible:',portfolio['Result']['Activos'][0]['Subtotal'][0]['Detalle'][0]['IMPO'], end='  ')
        except: pass

        for i in portfolio['Result']['Activos'][0]['Subtotal'][0]['APERTURA']:
            if i['IMPO'] != None: print(i['DETA'],':',i['IMPO'], end='  ' )
        print('||',time.strftime("%H:%M:%S"))

        subtotal = [ i['Subtotal'] for i in portfolio["Result"]["Activos"][0:] ]

        for i in subtotal[0:]:
            if i[0]['NERE'] != 'Pesos':  
                subtotal = [ ( x['NERE'],x['CAN0'],x['CANT']) for x in i[0:] if x['CANT'] != None]

                for x in subtotal:

                    for valor in shtData.range('A'+str(rangoDesde)+':'+'P'+str(rangoHasta)).value:

                        if not valor[0]: continue
                        ticker = str(valor[0]).split()
                        
                        if x[0] == ticker[0]: 
                            shtData.range('U'+str(int(valor[15]+1))).value = float(x[2])
                            if tipo == 1:
                                if len(ticker) < 2: 
                                    shtData.range('X'+str(int(valor[15]+1))).value = float(x[1])
                                else:
                                    shtData.range('X'+str(int(valor[15]+1))).value = float(x[1]) / 100
        
        #hb.online.disconnect()
    except: pass

def soloContinua():
    pass

def traerADR():
    valorAdr = yf.download(['GGAL','YPF'],period='1d',interval='1d',auto_adjust=False)['Close'].values
    shtData.range('Z61').value = valorAdr[0][0]
    shtData.range('Z62').value = valorAdr[0][1]
    shtData.range('Y62').value = time.strftime("%H:%M:%S")

vuelta = 0
vueltaPortfolio = 0

def haceRulo(celda=int):
    try:
        nombre = str(shtData.range('A'+str(int(celda))).value).split()
        precio = shtData.range('C'+str(int(celda))).value
        nominales = shtData.range('Y'+str(int(celda))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f'//___/ RULO AUT /___// - {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False: 
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass
    try:
        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range('D'+str(int(celda+1))).value
        nominales = shtData.range('Y'+str(int(celda+1))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

    try:
        nombre = str(shtData.range('A'+str(int(celda+2))).value).split()
        precio = shtData.range('C'+str(int(celda+2))).value
        nominales = shtData.range('Y'+str(int(celda+2))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' - {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

    try:
        nombre = str(shtData.range('A'+str(int(celda+3))).value).split()
        precio = shtData.range('D'+str(int(celda+3))).value
        nominales = shtData.range('Y'+str(int(celda+3))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' + {nominales} {nombre[0]} {precio}')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass
           


def buscoOperaciones():
    if not shtData.range('T1').value: # RULO AUTOMATICO activado por columna O
        celda = 2
        for i in shtData.range('O2:O25').value:
            if str(i).upper() == 'R':
                if celda==2 or celda==6 or celda==8 or celda==10 or celda==14 or celda==18 or celda==22:
                    haceRulo(celda)
            celda += 1
        shtData.range('T1').value = 'ROLL'

while True:
    buscoOperaciones()
    if time.strftime("%H:%M:%S") > '16:56:30': 
        if time.strftime("%H:%M:%S") < '16:56:45':
            print(time.strftime("%H:%M:%S"), 'Mercado local cerrado')
            shtData.range('Q1').value = 'PRECIOS'
            shtData.range('S1').value = 'ADR'
            shtData.range('T1').value = 'ROLL'
            shtData.range('W1').value = 'R'
            shtData.range('X1').value = 'STOP'
            shtData.range('Z1').value = 2
            pyRofex.close_websocket_connection()
            hb.online.disconnect()
    else:
        
        if not shtData.range('Q1').value:
            try:
                if esFinde == False: 
                    shtData.range('A30').options(index=False, headers=False).value = df_datos   
            except: print('Hubo un error al actualizar excel')

            if not shtData.range('Q1').value:
                if vueltaPortfolio > 20 : 
                    vueltaPortfolio = 0
                    try: 
                        getPortfolioHB(hb,'47352',2) 
                    except: 
                        print('Hubo un error al traer datos del portafolio')
                else: vueltaPortfolio += 1
            
        if not shtData.range('S1').value:
            try:
                if vuelta > 5: 
                    shtData.range('U1').value = time.strftime("%H:%M:%S")
                    traerADR()
                    vuelta = 0
                else: vuelta += 1
            except:
                shtData.range('Z61').value = shtData.range('F62').value
                shtData.range('Z63').value = shtData.range('F64').value

    #shtOperaciones.range('AI63').options(index=False, headers=False).value = operaciones
    time.sleep(2)

