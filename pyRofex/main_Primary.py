import pyRofex
import xlwings as xw
import time , math
import pandas as pd
import os
import environ
import requests
import random
import yfinance as yf

market_data_recibida = []
reporte_de_ordenes = []

env = environ.Env()
environ.Env.read_env()

wb = xw.Book('epgb.xlsb')
shtTickers = wb.sheets('pyRofex')
shtData = wb.sheets('HOME')

shtData.range('A1').value = 'symbol'
shtData.range('Q1').value = 'PRC'
shtData.range('R1').value = 'ADR'
shtData.range('S1').value = 'D'
#shtData.range('T1').value = 'ROLL'
shtData.range('U1').value = 'veta'
shtData.range('V1').value = 'bcch'
shtData.range('W1').value = 'AUTO'
shtData.range('X1').value = 'SCP'
shtData.range('Y31:Y65').value = ''
shtData.range('Z1').value = 0.5

rangoDesde = '10'
rangoHasta = '34'
reCompra = False
esFinde = False
noMatriz = False
scalpi = False

    
def diaLaboral():
    global esFinde
    hoyEs = time.strftime("%A")
    if hoyEs == 'Saturday' or hoyEs == 'Sunday':
        esFinde = True

def loguinHB():
    from pyhomebroker import HomeBroker  
    global hbVETA
    try:
        hbVETA = HomeBroker(int(os.environ.get('broker')))
        hbVETA.auth.login(dni=str(os.environ.get('dni')), 
        user=str(os.environ.get('user2')),  
        password=str(os.environ.get('password2')),
        raise_exception=True)
        shtData.range('U1').value = 'VETA'
        print(" *** online en HB VETA *** ")
    except: 
        print(" *  NO se pudo loguear en VETA HOME BROKER  *", time.strftime("%H:%M:%S"))

def loguinBCCH():
    from pyhomebroker import HomeBroker  
    global hb
    try:
        hb = HomeBroker(int(os.environ.get('broker2474')))
        hb.auth.login(dni=str(os.environ.get('dni2474')), 
        user=str(os.environ.get('user2474')),  
        password=str(os.environ.get('password2474')),
        raise_exception=True)
        shtData.range('V1').value = 'BCCH'
        print(" *** online en HB BCCH *** || ", time.strftime("%H:%M:%S"))
    except: 
        print(" *  NO se pudo loguear en BCCH HOME BROKER  *", time.strftime("%H:%M:%S"))

diaLaboral()

if esFinde == False:
    try:
        pyRofex._set_environment_parameter("url", "https://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)
        pyRofex._set_environment_parameter("ws", "wss://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)
        pyRofex.initialize(
            user=str(os.environ.get('user')), 
            password=str(os.environ.get('password')), 
            account=str(os.environ.get('account')), 
            environment=pyRofex.Environment.LIVE)
        shtData.range('U1').value = 'VETA'
        print("*** online en MATRIZ OMS VETA *** ", end=' ')
    except: 
        noMatriz = True
        print("No fue posible el logueo con MATRIZ OMS, sigue loguin en HB ... ")
else: print('FIN DE SEMANA, no se actualizan los precios locales.')

try: 
    loguinHB()
    #loguinBCCH()

except: 
    hb = ''
    hbVETA = ''

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
rng = shtTickers.range('H2:I2').expand() # BONOS
bonos = pd.DataFrame(rng.value, columns=['ticker', 'symbol'])
rng = shtTickers.range('T2:U2').expand() # CAUCHO
caucion = pd.DataFrame(rng.value, columns=['ticker', 'symbol'])

tickers = pd.concat([opciones, acc, bonos,cedear,ons,letras, caucion ])
listLength = len(acc) + 36 + len(opciones)
#listLength = 65 # largo sin trabajar con opciones
allLength = len(tickers) - len(acc)  - len(caucion) - 2

'''tickers = pd.concat([bonos, caucion ])
listLength =  31 
allLength = len(tickers)  - len(caucion) - 2'''


if noMatriz == False and esFinde == False:
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
else: df_datos = []

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
#df_order = pd.DataFrame()
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

if noMatriz == False and esFinde == False:
    pyRofex.init_websocket_connection(market_data_handler=market_data_handler)
    '''error_handler=error_handler,
    exception_handler=exception_handler,
    order_report_handler=order_report_handler'''
                                    
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
    #pyRofex.order_report_subscription(snapshot=True)
    #pyRofex.order_report_subscription()

def cancelarOrdenOMS(clientId):
  cancel_order = pyRofex.cancel(clientId)
  print(cancel_order)

def obtenerSaldoMatriz(cuenta=None):
    try:
        resumenCuenta = pyRofex.get_account_report(account=cuenta)
        resumenCuenta = resumenCuenta["accountData"]['detailedAccountReports']['1']['currencyBalance']['detailedCurrencyBalance']['ARS']['available']
        resumenCuenta /= 4
        shtData.range('M1').value = int(resumenCuenta)
        # print('Disponible Matriz para Gtias: ',resumenCuenta["accountData"]['availableToCollateral'],end=' | ')
    except: print('Error obterner disponible Gtias ')

def getPortfolioHB(hb, comitente, tipo):
    try:
        payload = {'comitente': str(comitente),
        'consolida': '0',
        'proceso': '22',
        'fechaDesde': None,
        'fechaHasta': None,
        'tipo': None,
        'especie': None,
        'comitenteMana': None}

        if tipo == 3:
            portfolio = requests.post("https://clientes.bcch.org.ar/Consultas/GetConsulta", cookies=hb.auth.cookies, json=payload).json()
            shtData.range('V31:V'+str(rangoHasta)).value = ''
            try: 
                shtData.range('O1').value = portfolio['Result']['Activos'][0]['Subtotal'][0]['APERTURA'][1]['ACUM']
            except: pass
        else:
            portfolio = requests.post("https://cuentas.vetacapital.com.ar/Consultas/GetConsulta", cookies=hb.auth.cookies, json=payload).json()
            shtData.range('U31:U'+str(rangoHasta)).value = ''
            try: 
                if tipo == 1:
                    shtData.range('M1').value = portfolio['Result']['Activos'][0]['Subtotal'][0]['APERTURA'][1]['ACUM']
                    shtData.range('O1').value = portfolio['Result']['Activos'][0]['Subtotal'][2]['APERTURA'][1]['ACUM']
            except: pass

        subtotal = [ i['Subtotal'] for i in portfolio["Result"]["Activos"][0:] ]

        for i in subtotal[0:]:
            if i[0]['NERE'] != 'Pesos':  
                subtotal = [ ( x['NERE'],x['CAN0'],x['CANT']) for x in i[0:] if x['CANT'] != None]
                for x in subtotal:
                    for valor in shtData.range('A31:P'+str(rangoHasta)).value:
                        if not valor[0]: continue
                        ticker = str(valor[0]).split()
                        if ticker[0][-1:] == 'D' or ticker[0][-1:] == 'C':  
                            if x[0] == ticker[0][:-1]: 
                                
                                if tipo == 3: shtData.range('V'+str(int(valor[15]+1))).value = int(x[2])
                                else: shtData.range('U'+str(int(valor[15]+1))).value = int(x[2])
                        else:
                            if x[0] == ticker[0]: 
                                
                                if tipo == 3: shtData.range('V'+str(int(valor[15]+1))).value = int(x[2])
                                else: shtData.range('U'+str(int(valor[15]+1))).value = int(x[2])

                                hayW = shtData.range('W'+str(int(valor[15]+1))).value

                                if tipo == 1 or tipo == 3:

                                    if len(ticker) < 2: 
                                        if not hayW: shtData.range('W'+str(int(valor[15]+1))).value = valor[5]
                                        shtData.range('X'+str(int(valor[15]+1))).value = float(x[1])
                                    
                                    else:
                                        if not hayW: 
                                            if ticker[0] == 'KO': shtData.range('W'+str(int(valor[15]+1))).value = valor[5]
                                            else: shtData.range('W'+str(int(valor[15]+1))).value = valor[5] / 100
                                        if ticker[0] == 'KO': shtData.range('X'+str(int(valor[15]+1))).value = float(x[1])
                                        else: shtData.range('X'+str(int(valor[15]+1))).value = float(x[1]) / 100

                                else:
                                    if len(ticker) < 2: 
                                        if not hayW: shtData.range('W'+str(int(valor[15]+1))).value = valor[5]
                                    else:
                                        if not hayW: 
                                            if ticker[0] == 'KO': shtData.range('W'+str(int(valor[15]+1))).value = valor[5]
                                            else: 
                                                shtData.range('W'+str(int(valor[15]+1))).value = valor[5] / 100
                                                shtData.range('X'+str(int(valor[15]+1))).value = float(x[1]) / 100
    except: pass

def cancelaCompraHB(celda):
    orderC = shtData.range('AF'+str(int(celda+1))).value
    if orderC == None: orderC = 0

    if esFinde == False: 
        try: 
            hb.orders.cancel_order(int(os.environ.get('account_id2474')),int(orderC))
            print(f"/// Cancelada Compra : {int(orderC)} ",end='\t')
        except: 
            print(f'Error al cancelar COMPRA {orderC} con HB')
    try: shtData.range('V'+str(int(celda+1))).value -= shtData.range('AE'+str(int(celda+1))).value
    except: pass
    shtData.range('AE'+str(int(celda+1))+':'+'AF'+str(int(celda+1))).value = ''
        
def cancelarVentaHB(celda):
    orderV = shtData.range('AH'+str(int(celda+1))).value
    if orderV == None: orderV = 0
    if esFinde == False: 
        try:
            hb.orders.cancel_order(int(os.environ.get('account_id2474')),int(orderV))
            print(f"/// Cancelada Venta  : {int(orderV)} " ,end='\t')
        except: 
            print(f'Error al cancelar VENTA {orderV} con HB')
    try: shtData.range('V'+str(int(celda+1))).value += shtData.range('AG'+str(int(celda+1))).value
    except: pass
    shtData.range('AG'+str(int(celda+1))+':'+'AH'+str(int(celda+1))).value = ''

def cancelarTodo(desde,hasta):
    if esFinde == False:
        try:  
            hb.orders.cancel_all_orders(int(os.environ.get('account_id2474')))
            print("/// Todas las ordenes activas canceladas ")
        except: pass
    shtData.range('AE'+str(desde)+':'+'AH'+str(hasta)).value = ''

def enviarOrdenHB(tipo=str,symbol=str, price=float, size=int, celda=int):
    global orderC, orderV
    orderC, orderV = 'S/D','S/D'
    symbol = str(shtData.range(str(symbol)).value).split()
    precio = shtData.range(str(price)).value

    if tipo.lower() == 'buy': 
        dinero = shtData.range('O1').value
        dinero = 0 if dinero == None else dinero
        try: 
            if len(symbol) < 2:
                if precio * size <= dinero:
                    if esFinde == False: orderC = hb.orders.send_buy_order(symbol[0],'24hs', float(precio), abs(int(size)))
                    print(f'        ______/ BUY HB  opcion + {int(size)} {symbol[0]} // precio: {precio} // {orderC}') 
                else: print('*** SIN ARS disponibles en HB BCCH *** ', time.strftime("%H:%M:%S"))
            else:
                if precio / 100 * size <= dinero:
                    if esFinde == False: orderC = hb.orders.send_buy_order(symbol[0],symbol[2], float(precio), abs(int(size)))
                    print(f'        ______/ BUY HB + {int(size)} {symbol[0]} {symbol[2]} // precio: {round(precio/100,4)} // {orderC}')
                else: print('*** SIN ARS disponibles en HB BCCH *** ', time.strftime("%H:%M:%S"))
            try: shtData.range('V'+str(int(celda+1))).value += abs(int(size))
            except: shtData.range('V'+str(int(celda+1))).value = abs(int(size))
            shtData.range('AE'+str(int(celda+1))).value = abs(int(size))
            shtData.range('AF'+str(int(celda+1))).value = orderC
        except: 
            shtData.range('Q'+str(int(celda+1))+':'+'R'+str(int(celda+1))).value = ''
            print(f'        ______/ ERROR en COMPRA HB. {symbol[0]} // precio: {precio} // + {int(size)}')
        
    else: # VENTA
        stock = shtData.range('V'+str(int(celda+1))).value
        stock = 0 if stock == None else stock
        try:
            if stock >= size:
                if stock < size: size = stock
                if len(symbol) < 2:
                    if esFinde == False: orderV = hb.orders.send_sell_order(symbol[0],'24hs', float(precio), abs(int(size)))
                    print(f'______/ SELL opcion - {int(size)} {symbol[0]} // precio: {precio} // {orderV}')
                else:
                    if esFinde == False: orderV = hb.orders.send_sell_order(symbol[0],symbol[2], float(precio), abs(int(size)))
                    print(f'______/ SELL - {int(size)} {symbol[0]} {symbol[2]} // precio: {round(precio/100,4)} // {orderV}')
                try: shtData.range('V'+str(int(celda+1))).value -= abs(int(size))
                except: shtData.range('V'+str(int(celda+1))).value = int(size) / -1
                shtData.range('AG'+str(int(celda+1))).value = abs(int(size))
                shtData.range('AH'+str(int(celda+1))).value = orderV
            else: print('*** SIN STOCK disponibles en HB BCCH *** ', time.strftime("%H:%M:%S"))
        except:
            shtData.range('S'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            print(f'______/ ERROR en VENTA. {symbol[0]} // precio: {precio} // {int(size)}')

def soloContinua():
    global cantidad
    cantidad = 0
    pass

def namesArs(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA37D'+plazo
    elif nombre[:2] == 'TY': return 'TY30P'+plazo
    elif nombre[:2] == 'KO': return 'KO'+plazo
    elif nombre[:2] == 'GOGL': return 'GOOGL'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') and (nombre[3:4] == 'D' or nombre[3:4] == 'C'):
        if (nombre[1:2] == 'F' or nombre[1:2] == 'Y'): return nombre[:1]+'20'+nombre[1:3]+plazo
        if (nombre[1:2] == 'M'): return nombre[:1]+'31'+nombre[1:3]+plazo
        if (nombre[1:2] == 'N'): return nombre[:1]+'29'+nombre[1:3]+plazo
        if (nombre[1:2] == 'J'): return nombre[:1]+'18'+nombre[1:3]+plazo
        if (nombre[1:2] == 'G'): return nombre[:1]+'29'+nombre[1:3]+plazo
        if (nombre[1:2] == 'O'): return nombre[:1]+'14'+nombre[1:3]+plazo
        if (nombre[1:2] == 'E'): return nombre[:1]+'31'+nombre[1:3]+plazo
        else: return nombre[:1]+'31'+nombre[1:3]+plazo
    elif (nombre[:2] == 'MR' or nombre[:2] == 'CL') and (nombre[4:5] == 'D' or nombre[4:5] == 'C'):
        return nombre[:4]+'O'+plazo 
    else: return nombre[:4]+plazo

def namesCcl(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA7DC'+plazo
    elif nombre[:2] == 'BP': return 'BPA7C'+plazo
    elif nombre[:2] == 'KO': return 'KOC'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') :
        if nombre[3:4] == 'D': return nombre[:3]+'C'+plazo
        else: return nombre[:1]+nombre[3:5]+'C'+plazo
    else: return nombre[:4]+'C'+plazo

def namesMep(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA7DD'+plazo
    elif nombre[:2] == 'BP': return 'BPA7D'+plazo
    elif nombre[:2] == 'KO': return 'KOD'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') :
        if nombre[3:4] == 'C': return nombre[:3]+'D'+plazo
        else: return nombre[:1]+nombre[3:5]+'D'+plazo
    else: return nombre[:4]+'D'+plazo

def cargoXplazo(dicc,moneda):
    mejorMep = dicc['mepCI'][0]
    mejorMep24 = dicc['mep24'][0]
    mejorCcl = dicc['cclCI'][0]
    mejorCcl24 = dicc['ccl24'][0]
    mepArs = namesMep(dicc['arsCImep'][0],' - CI')
    mepArs24 = namesMep(dicc['ars24mep'][0],' - 24hs')
    mepCcl = namesMep(dicc['cclCI'][0],' - CI')
    mepCcl24 = namesMep(dicc['ccl24'][0],' - 24hs')

    if str(moneda).upper() == 'PD':
        shtData.range('A2').value = namesArs(dicc['mepCI'][0],' - CI')
        shtData.range('A3').value = mejorMep
        shtData.range('A4').value = 'AL30D - CI'
        shtData.range('A5').value = 'AL30 - CI'
        shtData.range('A6').value = namesArs(dicc['mep24'][0],' - 24hs')
        shtData.range('A7').value = mejorMep24
        shtData.range('A8').value = 'AL30D - 24hs'
        shtData.range('A9').value = 'AL30 - 24hs'
        shtData.range('A10').value = namesArs(dicc['mepCI'][0],' - CI')
        shtData.range('A11').value = mejorMep
        shtData.range('A12').value = mepArs
        shtData.range('A13').value = dicc['arsCImep'][0]
        shtData.range('A14').value = namesArs(dicc['mep24'][0],' - 24hs')
        shtData.range('A15').value = mejorMep24
        shtData.range('A16').value = mepArs24
        shtData.range('A17').value = dicc['ars24mep'][0]
        # ---------------------------------------------------
        '''shtData.range('A26').value = namesArs(dicc['mep24'][0],' - 24hs')
        shtData.range('A27').value = mejorMep24
        shtData.range('A28').value = 'AL30 - 24hs'
        shtData.range('A29').value = 'AL30D - 24hs'
        '''
        
    elif str(moneda).upper() == 'PC':
        shtData.range('A2').value = namesArs(dicc['cclCI'][0],' - CI')
        shtData.range('A3').value = mejorCcl
        shtData.range('A4').value = 'AL30C - CI'
        shtData.range('A5').value = 'AL30 - CI'
        shtData.range('A6').value = namesArs(dicc['ccl24'][0],' - 24hs')
        shtData.range('A7').value = mejorCcl24
        shtData.range('A8').value = 'AL30C - 24hs'
        shtData.range('A9').value = 'AL30 - 24hs'
        shtData.range('A10').value = namesArs(dicc['cclCI'][0],' - CI')
        shtData.range('A11').value = mejorCcl
        shtData.range('A12').value = namesCcl(dicc['arsCIccl'][0],' - CI')
        shtData.range('A13').value = dicc['arsCIccl'][0]
        shtData.range('A14').value = namesArs(dicc['ccl24'][0],' - 24hs')
        shtData.range('A15').value = mejorCcl24
        shtData.range('A16').value = namesCcl(dicc['ars24ccl'][0],' - 24hs')
        shtData.range('A17').value = dicc['ars24ccl'][0]
        # ---------------------------------------------------
        '''shtData.range('A26').value = namesArs(dicc['ccl24'][0],' - 24hs')
        shtData.range('A27').value = mejorCcl24
        shtData.range('A28').value = 'AL30C - 24hs'
        shtData.range('A29').value = 'AL30 - 24hs'
        '''
    if str(moneda).upper() == 'DP':
        shtData.range('A2').value = mejorMep
        shtData.range('A3').value = namesArs(dicc['mepCI'][0],' - CI')
        shtData.range('A4').value = 'AL30 - CI'
        shtData.range('A5').value = 'AL30D - CI'
        shtData.range('A6').value = mejorMep24
        shtData.range('A7').value = namesArs(dicc['mep24'][0],' - 24hs')
        shtData.range('A8').value = 'AL30 - 24hs'
        shtData.range('A9').value = 'AL30D - 24hs'
        shtData.range('A10').value = mejorMep
        shtData.range('A11').value = namesArs(dicc['mepCI'][0],' - CI')
        shtData.range('A12').value = dicc['arsCImep'][0]
        shtData.range('A13').value = namesMep(dicc['arsCImep'][0],' - CI')
        shtData.range('A14').value = mejorMep24
        shtData.range('A15').value = namesArs(dicc['mep24'][0],' - 24hs')
        shtData.range('A16').value = dicc['ars24mep'][0]
        shtData.range('A17').value = namesMep(dicc['ars24mep'][0],' - 24hs')
        # -------------------------------------------------------
        '''shtData.range('A26').value = mejorMep24
        shtData.range('A27').value = namesArs(dicc['mep24'][0],' - 24hs')
        shtData.range('A28').value = 'AL30D - 24hs'
        shtData.range('A29').value = 'AL30 - 24hs'
        '''

    elif str(moneda).upper() == 'DC':
        shtData.range('A2').value = mejorMep
        shtData.range('A3').value = namesCcl(dicc['mepCI'][0],' - CI')
        shtData.range('A4').value = 'AL30C - CI'
        shtData.range('A5').value = 'AL30D - CI'
        shtData.range('A6').value = mejorMep24
        shtData.range('A7').value = namesCcl(dicc['mep24'][0],' - 24hs')
        shtData.range('A8').value = 'AL30C - 24hs'
        shtData.range('A9').value = 'AL30D - 24hs'
        shtData.range('A10').value = mejorMep
        shtData.range('A11').value = namesCcl(dicc['mepCI'][0],' - CI')
        shtData.range('A12').value = mejorCcl
        shtData.range('A13').value = namesMep(dicc['cclCI'][0],' - CI')
        shtData.range('A14').value = mejorMep24
        shtData.range('A15').value = namesCcl(dicc['mep24'][0],' - 24hs')
        shtData.range('A16').value = mejorCcl24
        shtData.range('A17').value = namesMep(dicc['ccl24'][0],' - 24hs')
    
    if str(moneda).upper() == 'CP':
        shtData.range('A2').value = mejorCcl
        shtData.range('A3').value = namesArs(dicc['cclCI'][0],' - CI')
        shtData.range('A4').value = 'AL30 - CI'
        shtData.range('A5').value = 'AL30C - CI'
        shtData.range('A6').value = mejorCcl24
        shtData.range('A7').value = namesArs(dicc['ccl24'][0],' - 24hs')
        shtData.range('A8').value = 'AL30 - 24hs'
        shtData.range('A9').value = 'AL30C - 24hs'
        # --------------------------------------------------
        '''shtData.range('A26').value = mejorCcl24
        shtData.range('A27').value = namesArs(dicc['ccl24'][0],' - 24hs')
        shtData.range('A28').value = 'AL30 - 24hs'
        shtData.range('A29').value = 'AL30C - 24hs'
        '''
    elif str(moneda).upper() == 'CD':
        shtData.range('A2').value = mejorCcl
        shtData.range('A3').value = namesMep(dicc['cclCI'][0],' - CI')
        shtData.range('A4').value = 'AL30D - CI'
        shtData.range('A5').value = 'AL30C - CI'
        shtData.range('A6').value = mejorCcl24
        shtData.range('A7').value = namesMep(dicc['ccl24'][0],' - 24hs')
        shtData.range('A8').value = 'AL30D - 24hs'
        shtData.range('A9').value = 'AL30C - 24hs'
 

    '''shtData.range('A10').value = dicc['arsCImep'][0]
    shtData.range('A11').value = mepArs
    shtData.range('A12').value = mejorMep
    shtData.range('A13').value = namesArs(dicc['mepCI'][0],' - CI')
    shtData.range('A14').value = dicc['ars24mep'][0]
    shtData.range('A15').value = mepArs24
    shtData.range('A16').value = mejorMep24
    shtData.range('A17').value = namesArs(dicc['mep24'][0],' - 24hs')
    shtData.range('A18').value = dicc['cclCI'][0]
    shtData.range('A19').value = mepCcl
    shtData.range('A20').value = mejorMep
    shtData.range('A21').value = namesCcl(dicc['mepCI'][0],' - CI')
    shtData.range('A22').value = dicc['ccl24'][0]
    shtData.range('A23').value = mepCcl24
    shtData.range('A24').value = mejorMep24
    shtData.range('A25').value = namesCcl(dicc['mep24'][0],' - 24hs')

    {'cclCI': ['AL30C - CI', 1156.071964017991], 'ccl24': ['AL30C - 24hs', 0.9878100531348846], 
         'mepCI': ['AL41D - CI', 1.001461810048996], 'mep24': ['GD38D - 24hs', 1.0023826834418272], 
         'arsCIccl': ['AL30 - CI', 1156.071964017991], 'ars24ccl': ['GD30 - 24hs', 1157.8330893118593], 
         'arsCImep': ['AE38 - CI', 1143.4628975265018], 'ars24mep': ['AL35 - 24hs', 1145.8394160583941]}
    
    '''
    shtData.range('A1').value = 'symbol'

def preparaRulo(monedaInicial):
    celda,pesos,dolar = listLength,1000,0.01
    tikers = {'cclCI':['',dolar],'ccl24':['',dolar],'mepCI':['',dolar],'mep24':['',dolar],'arsCIccl':['',pesos],'ars24ccl':['',pesos],'arsCImep':['',pesos],'ars24mep':['',pesos]}
    for valor in shtData.range('A70:A143').value:
        if not valor: continue
        name = str(valor).split()
        if str(name[2]).lower() == 'ci':
            if str(name[0][-1:]).upper()=='C':
                arsC = shtData.range('AA'+str(celda)).value
                if not arsC: arsC = 1000
                if arsC > tikers['arsCIccl'][1]: tikers['arsCIccl'] = [namesArs(name[0],' - CI'),arsC]
                ccl = shtData.range('Z'+str(celda)).value
                if not ccl: ccl = 0.01
                if ccl > tikers['cclCI'][1]: tikers['cclCI'] = [valor,ccl]
            if str(name[0][-1:]).upper()=='D': 
                arsM = shtData.range('AA'+str(celda)).value
                if not arsM: arsM = 1000
                if arsM > tikers['arsCImep'][1]: tikers['arsCImep'] = [namesArs(name[0],' - CI'),arsM]
                mep = shtData.range('Z'+str(celda)).value
                if not mep: mep = 0.01
                if mep > tikers['mepCI'][1]: tikers['mepCI'] = [valor,mep]

        if str(name[2]) == '24hs':
            if str(name[0][-1:]).upper()=='C':
                arsC = shtData.range('AA'+str(celda)).value
                if not arsC: arsC = 1000
                if arsC > tikers['ars24ccl'][1]: tikers['ars24ccl'] = [namesArs(name[0],' - 24hs'),arsC]
                ccl = shtData.range('Z'+str(celda)).value
                if not ccl: ccl = 0.01
                if ccl > tikers['ccl24'][1]: tikers['ccl24'] = [valor,ccl]
            if str(name[0][-1:]).upper()=='D': 
                arsM = shtData.range('AA'+str(celda)).value
                if not arsM: arsM = 1000
                if arsM > tikers['ars24mep'][1]: tikers['ars24mep'] = [namesArs(name[0],' - 24hs'),arsM]
                mep = shtData.range('Z'+str(celda)).value
                if not mep: mep = 0.01
                if mep > tikers['mep24'][1]: tikers['mep24'] = [valor,mep]
        celda +=1 
    cargoXplazo(tikers,monedaInicial)

def traerADR():
    #valorAdr = yf.download(['GGAL'],period='1d',interval='1d',auto_adjust=False)['Close'].values
    valorAdr = yf.download(['GGAL'],period='1d',interval='1d',auto_adjust=False)['Close'].values
    shtData.range('Z66').value = valorAdr[0][0]
    '''max = yf.download(['GGAL'],period='1d',interval='1d',auto_adjust=False)['High'].values
    min = yf.download(['GGAL'],period='1d',interval='1d',auto_adjust=False)['Low'].values
    shtData.range('AB61').value = max[0][0]
    shtData.range('AB62').value = min[0][0]'''
    shtData.range('Y67').value = time.strftime("%H:%M:%S")

def ruloAutomatico(celda): # Rulo automatico para HOME BROKER
    if celda+1 == 2 or celda+1 == 6 or celda+1 == 8 or celda+1 == 14 or celda+1 == 18 or celda+1 == 22:
        hayStock = cantidadAuto(celda+1)
        if hayStock == 0:
            print('NO tiene stock disponible para iniciar el rulo ')
        else:
            shtData.range('S'+str(int(celda+1))).value = "-"
            shtData.range('R'+str(int(celda+2))).value = "+"
            shtData.range('S'+str(int(celda+3))).value = "-"
            shtData.range('R'+str(int(celda+4))).value = "+"

def cantidadAuto(celda):
    cantidad = shtData.range('Y'+str(int(celda))).value
    cantidad = 0 if cantidad == None else cantidad
    return abs(int(cantidad))

def stockU(celda):
    stok = shtData.range('U'+str(int(celda))).value
    stok = 0 if not stok or stok == None or stok == 'None' else stok        
    return int(stok)

def stockV(celda):
    stok = shtData.range('V'+str(int(celda))).value
    stok = 0 if not stok or stok == None or stok == 'None' else stok     
    return int(stok)


def scalping(celda,lado,tipo,stock=int,nominales=int):
    try: 
        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range(str(lado)+str(int(celda+1))).value
        if len(nombre) < 2: 
            symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
            ganancia = shtData.range('Z1').value * 10
            if str(tipo).upper() == 'BUY': shtData.range('W'+str(int(celda+1))).value = precio
        else: 
            symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
            ganancia = shtData.range('Z1').value * 100
            if str(tipo).upper() == 'BUY': shtData.range('W'+str(int(celda+1))).value = precio / 100
            
        if str(tipo).upper() == 'BUY':
            print(f'//___/ SCALPING BUY  /___ + {nominales} {nombre[0]} // precio: {precio}', end=' | ')
            if esFinde == False and noMatriz == False:
                pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
                if nombre[0][-1:] == 'D' or nombre[0][-1:] == 'C':
                    precio += ganancia / 100
                else: precio += ganancia
                if abs(stock) < nominales: nominales = abs(stock)
                pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
            else: 
                if nombre[0][-1:] == 'D' or nombre[0][-1:] == 'C':
                    precio += ganancia / 100
                else: precio += ganancia
                if abs(stock) < nominales: nominales = abs(stock)
            print(f'___/ SELL /___ - {nominales} {nombre[0]} // precio: {precio}')
        
        else:
            print(f'//___/ SCALPING SELL /___ - {nominales} {nombre[0]} // precio: {precio}',end=' | ')
            if esFinde == False:
                pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
                if nombre[0][-1:] == 'D' or nombre[0][-1:] == 'C':
                    precio -= ganancia / 100
                else: precio -= ganancia
                if len(nombre) < 2 and lado == 'C': print('VENTA descubierto al bid, NO se pide la recompra',end=' ')
                else: pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
            else: 
                if nombre[0][-1:] == 'D' or nombre[0][-1:] == 'C':
                    precio -= ganancia / 100
                else: precio -= ganancia
                if len(nombre) < 2 and lado == 'C': print('ATENCION: Se vende descubierto pero no se pide la recompra',end=' ')
            print(f'___/ BUY /___ + {nominales} {nombre[0]} // precio: {precio}')   
    except: pass

def operacionRapida(celda,lado,tipo, stock=int, nominales=int):
    esCaucho = False
    try: 
        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range(str(lado)+str(int(celda+1))).value

        if len(nombre) == 2: 
            symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[1])
            esCaucho = True
        elif len(nombre) < 2: 
            symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
            if str(tipo).upper() == 'BUY': shtData.range('W'+str(int(celda+1))).value = precio
        else: 
            symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
            if str(tipo).upper() == 'BUY': shtData.range('W'+str(int(celda+1))).value = precio / 100
        
        if str(tipo).upper() == 'BUY':
            print(f'//___/ BUY  /___// + {nominales} {nombre[0]} // precio: {precio}')
            if esFinde == False and noMatriz == False:
                pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
        else:
            if esCaucho == False: 
                if abs(stock) < nominales: nominales = abs(stock)
            else: 
                print(f'Coloca CAUCION ... {nombre[0]} {nombre[1]} ', end='')
                pass
            print(f'//___/ SELL /___// - {nominales} {nombre[0]} // precio: {precio}')
            if esFinde == False and noMatriz == False:
                pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

def buyRollPlus(celda=int):
    try:
        nominales = cantidadAuto(celda)
        nombre = str(shtData.range('A'+str(int(celda))).value).split()
        precio = shtData.range('D'+str(int(celda))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f'//___/ BUY ROLL plus /___ + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False and noMatriz == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
            time.sleep(1)
    except: pass

    try:
        nominales = cantidadAuto(celda+1)
        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range('C'+str(int(celda+1))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' - {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False and noMatriz == False: 
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
            time.sleep(1)
    except: pass

    try:
        nominales = cantidadAuto(celda+2)
        nombre = str(shtData.range('A'+str(int(celda+2))).value).split()
        precio = shtData.range('D'+str(int(celda+2))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False and noMatriz == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
            time.sleep(1)
    except: pass

    try:
        nominales = cantidadAuto(celda+3)
        nombre = str(shtData.range('A'+str(int(celda+3))).value).split()
        precio = shtData.range('D'+str(int(celda+3))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' - {nominales} {nombre[0]} {precio} ')
        if esFinde == False and noMatriz == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
            time.sleep(1)
    except: pass

def buyRoll(celda=int):
    try:
        nominales = cantidadAuto(celda)
        nombre = str(shtData.range('A'+str(int(celda))).value).split()
        precio = shtData.range('D'+str(int(celda))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f'//___/ BUY ROLL /___ + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False and noMatriz == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
            time.sleep(1)
    except: pass

    try:
        nominales = cantidadAuto(celda+1)
        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range('C'+str(int(celda+1))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' - {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False and noMatriz == False: 
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
            time.sleep(1)
    except: pass

    try:
        nominales = cantidadAuto(celda+2)
        nombre = str(shtData.range('A'+str(int(celda+2))).value).split()
        precio = shtData.range('D'+str(int(celda+2))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False and noMatriz == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
            time.sleep(1)
    except: pass

    try:
        nominales = cantidadAuto(celda+3)
        nombre = str(shtData.range('A'+str(int(celda+3))).value).split()
        precio = shtData.range('C'+str(int(celda+3))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' - {nominales} {nombre[0]} {precio} ')
        if esFinde == False and noMatriz == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
            time.sleep(1)
    except: pass

def roll():
    celda = 2
    for i in shtData.range('O2:O18').value:
            if str(i).upper() == 'R':
                if celda==2 or celda==6 or celda==8 or celda==10 or celda==14 or celda==18 or celda==22:
                    buyRoll(celda)
            celda += 1

def posicionRulo(celda):
    if celda==2 or celda==6 or celda==8 or celda==10 or celda==14 or celda==18 or celda==22:
        return 'ok'

def compraUsd(celda,ladoCompra):
    compra = str(shtData.range('A'+str(int(celda))).value).split()
    if len(compra) < 2: pass
    else:
        if compra[2] == '24hs' or compra[2] == 'CI': 
            vende = str(shtData.range('M'+str(int(celda))).value).split()
            if len(vende) < 2: pass
            else:
                if vende[2] == '24hs'or vende[2] == 'CI':
                    nominales = cantidadAuto(celda)
                    gastos = shtData.range('AB1').value

                    ask = shtData.range(str(ladoCompra)+str(int(celda))).value
                    symbol = "MERV - XMEV - " + str(compra[0]) + ' - ' + str(compra[2])
                    print(f'_____ buy mep COMPRO + {int(nominales)} {compra[0]} {compra[2]} {ask} ',end=' | ')
                    if esFinde == False and noMatriz == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(ask),order_type=pyRofex.OrderType.LIMIT)

                    bid = shtData.range('H'+str(int(celda))).value
                    symbol = "MERV - XMEV - " + str(vende[0]) + ' - ' + str(vende[2])
                    print(f'_____ - {int(nominales)} {vende[0]} {vende[2]} {bid}',end=' || MEP: ')
                    compraMep = shtData.range('AB'+str(int(celda))).value
                    print(round(float(compraMep),2))
                    
                    if esFinde == False and noMatriz == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
                    
                    try: shtData.range('M1').value -= (ask / 100) * (1 + gastos) * nominales
                    except: pass
                    try: shtData.range('O1').value += (bid / 100) * (1 - gastos) * nominales
                    except: pass

                    shtData.range('Z'+str(int(celda))).value = round(float(compraMep),2)
                else: pass
        else: pass
        shtData.range('Q'+str(int(celda))).value = ''

def vendeUsd(celda,ladoCompra):
    compra = str(shtData.range('M'+str(int(celda))).value).split()
    if len(compra) < 2: pass
    else:
        if compra[2] == '24hs' or compra[2] == 'CI': 
            vende = str(shtData.range('A'+str(int(celda))).value).split()
            if len(vende) < 2: pass
            else:
                if vende[2] == '24hs'or vende[2] == 'CI':
                    nominales = cantidadAuto(celda)
                    gastos = shtData.range('AB1').value
                    
                    ask = shtData.range(str(ladoCompra)+str(int(celda))).value 
                    symbol = "MERV - XMEV - " + str(compra[0]) + ' - ' + str(compra[2])
                    print(f'_____ sell mep COMPRO + {int(nominales)} {compra[0]} {compra[2]} {ask} ',end=' | ')
                    if esFinde == False and noMatriz == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(ask),order_type=pyRofex.OrderType.LIMIT)
                    
                    bid = shtData.range('C'+str(int(celda))).value
                    symbol = "MERV - XMEV - " + str(vende[0]) + ' - ' + str(vende[2])
                    print(f'_____ - {int(nominales)} {vende[0]} {vende[2]} {bid}',end=' || MEP: ')
                    vendoMep = shtData.range('AC'+str(int(celda))).value
                    print(round(float(vendoMep),2))
                    if esFinde == False and noMatriz == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
                    
                    try: shtData.range('M1').value += (bid / 100) * (1 - gastos) * nominales
                    except: pass
                    try: shtData.range('O1').value -= (ask / 100) * (1 + gastos) * nominales
                    except: pass

                    shtData.range('AA'+str(int(celda))).value = round(float(vendoMep),2)
                else: pass
        else: pass
        shtData.range('Q'+str(int(celda))).value = ''

def compraAuto(celda,esDolar,subeBaja):
    nombre = str(shtData.range('A'+str(int(celda))).value).split()
    if esDolar == False: precio = shtData.range('C'+str(int(celda))).value
    else: precio = shtData.range('C'+str(int(celda))).value
    symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' +str(nombre[2])
    nominales = shtData.range('Y'+str(int(celda))).value
    if subeBaja == 'down': nominales /= 4
    else: nominales /= 2
    nominales = 10 if nominales < 9 else nominales
    nominales = 1000 if nominales > 999 else nominales

    if esDolar == False: dinero = shtData.range('M1').value
    else: dinero = shtData.range('O1').value

    dinero = 0 if dinero == None else dinero
    costo = precio /100 * nominales

    if costo <= dinero:
        print(f'*** ___ COMPRA automatica ___ + {int(nominales)} {nombre[0]} {nombre[2]} // precio: {precio}')
        if esFinde == False and noMatriz == False: 
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
            try:
                if esDolar == False: shtData.range('M1').value -= costo
                else:  shtData.range('O1').value -= costo

                shtData.range('U'+str(int(celda))).value += nominales
            except: pass
    else:
        nominales = dinero / int(precio / 100)
        if nominales > 0:
            costo = precio /100 * nominales
            print(f'*** ___ COMPRA automatica ___ + {int(nominales)} {nombre[0]} {nombre[2]} // precio: {precio}')
            if esFinde == False and noMatriz == False: 
                pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)   
                try:
                    if esDolar == False: shtData.range('M1').value -= costo
                    else:  shtData.range('O1').value -= costo
                    shtData.range('U'+str(int(celda))).value += nominales
                except: pass

def rollNuevo():
    inicio = 18
    arsUSD = {}
    usdARS = {}
    for valor in shtData.range('AB18:AB30').value:
        if not valor: continue
        arsUSD[inicio] = valor
        inicio += 1
    inicio = 18
    for valor in shtData.range('AC18:AC30').value:
        if not valor: continue
        usdARS[inicio] = valor
        inicio += 1
    inicio = 18
    shtData.range('A10').value = shtData.range('A'+str(min(arsUSD,key=arsUSD.get))).value
    shtData.range('A11').value = shtData.range('M'+str(min(arsUSD,key=arsUSD.get))).value
    shtData.range('A12').value = shtData.range('M'+str(max(usdARS,key=usdARS.get))).value
    shtData.range('A13').value = shtData.range('A'+str(max(usdARS,key=usdARS.get))).value
    
    shtData.range('A14').value = shtData.range('M'+str(max(usdARS,key=usdARS.get))).value
    shtData.range('A15').value = shtData.range('A'+str(max(usdARS,key=usdARS.get))).value
    shtData.range('A16').value = shtData.range('A'+str(min(arsUSD,key=arsUSD.get))).value
    shtData.range('A17').value = shtData.range('M'+str(min(arsUSD,key=arsUSD.get))).value

def buscoOperaciones(inicio,fin):
    hora = time.strftime("%H:%M:%S")
    auto = shtData.range('W1').value

    for valor in shtData.range('P'+str(inicio)+':'+'U'+str(fin)).value:

        if not valor[5] or valor[5] == None or valor[5] == 0 or hora <= '11:01:00': pass  
        else:
            cantidad = cantidadAuto(valor[0]+1)
            if cantidad != 0:
                nominalDescubierto = True if valor[5] < 0 else False
                scalpingStop('A'+str((int(valor[0]+1))),cantidad,int(valor[0]),nominalDescubierto,auto)

        if valor[1]: # # Columna Q en el excel /////////////////////////////////////////////////////////////////////////////////
            if str(valor[1]).lower() == 'r' and posicionRulo(valor[0]+1) == 'ok': buyRoll(valor[0]+1)
            elif str(valor[1]).upper() == 'P':  getPortfolioHB(hbVETA, os.environ.get('account_id'),1)
            elif str(valor[1]).lower() == 'rr' and posicionRulo(valor[0]+1) == 'ok': buyRollPlus(valor[0]+1)
            elif str(valor[1]).lower() == 'd': compraUsd(valor[0]+1,'C')
            elif str(valor[1]).lower() == 'm': baseEjercible(valor[0])
            elif str(valor[1]).lower() == 'cm': cerrarMariposa(valor[0])
            elif str(valor[1]).lower() == 'sm': verificaMariposa(valor[0])
            elif str(valor[1]).lower() == 't': compraTasa(valor[0]+1,'C')
            elif valor[1] == '+': 
                cantidad = cantidadAuto(valor[0]+1)
                operacionRapida(valor[0],'C','BUY',valor[5], cantidad)
            elif str(valor[1]).lower() == 's': 
                cantidad = cantidadAuto(valor[0]+1)
                scalping(valor[0],'C','BUY', valor[5], cantidad)
            else: 
                try: enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0]) # Compra Bid
                except: pass

            '''if bcch == None:
                if str(valor[1]).lower() == 'x': cancelaCompraHB(valor[0])
                elif str(valor[1]).lower() == 'xx': cancelarTodo(inicio,fin)
                elif valor[1] == '+': 
                    enviarOrdenHB('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
                elif str(valor[1]).upper() == 'P': getPortfolioHB(hb, os.environ.get('account_id2474'),3)
                else: 
                    try: 
                        if shtData.range('AF'+str(int(valor[0]+1))).value: cancelaCompraHB(valor[0]) # CANCELA oreden compra anterior
                        enviarOrdenHB('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0]) # Compra Bid
                    except: pass'''
            shtData.range('Q'+str(int(valor[0]+1))).value = ''

        if valor[2]: #  Columna R en el excel //////////////////////////////////////////////////////////////////////////////////:
            if str(valor[2]).lower() == 'r' and posicionRulo(valor[0]+1) == 'ok': buyRoll(valor[0]+1)
            elif str(valor[2]).lower() == 'd': compraUsd(valor[0]+1,'D')
            elif str(valor[2]).lower() == 't': compraTasa(valor[0],'D')
            elif str(valor[2]).upper() == 'P':  getPortfolioHB(hbVETA, os.environ.get('account_id'),1)
            elif valor[2] == '+': 
                cantidad = cantidadAuto(valor[0]+1)
                operacionRapida(valor[0],'D','BUY', valor[5], cantidad)
            elif str(valor[2]).lower() == 's': 
                cantidad = cantidadAuto(valor[0]+1)
                scalping(valor[0],'D','BUY', valor[5], cantidad)
            else: 
                try: enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0]) # Compra Ask
                except: pass
            '''if bcch == None:
                if str(valor[2]).lower() == 'x': cancelaCompraHB(valor[0])
                elif str(valor[2]).lower() == 'xx': cancelarTodo(inicio,fin)
                elif valor[2] == '+': 
                    enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
                elif str(valor[2]).upper() == 'P': getPortfolioHB(hb, os.environ.get('account_id2474'), 3)
                else: 
                    try: 
                        if shtData.range('AB'+str(int(valor[0]+1))).value: cancelaCompraHB(valor[0])
                        enviarOrdenHB('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0]) # Compra Ask
                    except: pass'''

            shtData.range('R'+str(int(valor[0]+1))).value = ''
        
        if valor[3]: # Columna S en el excel ///////////////////////////////////////////////////////////////////////////////////
            if str(valor[3]).lower() == 'r' and posicionRulo(valor[0]+1) == 'ok': buyRoll(valor[0]+1)
            elif str(valor[3]).lower() == 'd': vendeUsd(valor[0]+1,'H')
            elif str(valor[3]).upper() == 'P':  getPortfolioHB(hbVETA, os.environ.get('account_id'),1)
            elif valor[3] == '-': 
                cantidad = cantidadAuto(valor[0]+1)
                operacionRapida(valor[0],'C','SELL', valor[5], cantidad)
            elif str(valor[3]).lower() == 's': 
                cantidad = cantidadAuto(valor[0]+1)
                scalping(valor[0],'C','SELL', valor[5], cantidad)
            else: 
                try: enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0]) # Vendo Bid
                except: pass
            '''if bcch == None:
                if str(valor[3]).lower() == 'x': cancelarVentaHB(valor[0])
                elif str(valor[3]).lower() == 'xx': cancelarTodo(inicio,fin)
                elif valor[3] == '-': 
                    enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
                elif str(valor[3]).upper() == 'P': getPortfolioHB(hb, os.environ.get('account_id2474'), 3)
                else: 
                    try: 
                        if shtData.range('AE'+str(int(valor[0]+1))).value: cancelarVentaHB(valor[0])
                        enviarOrdenHB('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0]) # Vendo Bid
                    except: pass'''

            shtData.range('S'+str(int(valor[0]+1))).value = ''

        if valor[4]: # Columna T en el excel //////////////////////////////////////////////////////////////////////////////////
            if str(valor[4]).lower() == 'r' and posicionRulo(valor[0]+1) == 'ok': buyRoll(valor[0]+1)
            elif str(valor[4]).lower() == 'd': vendeUsd(valor[0]+1,'I')
            elif str(valor[4]).upper() == 'P':  getPortfolioHB(hbVETA, os.environ.get('account_id'),1)
            elif valor[4] == '-': 
                cantidad = cantidadAuto(valor[0]+1)
                operacionRapida(valor[0],'D','SELL', valor[5], cantidad)
            elif str(valor[4]).lower() == 's': 
                cantidad = cantidadAuto(valor[0]+1)
                scalping(valor[0],'D','SELL', valor[5], cantidad)
            else: 
                try: enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0]) # Vendo Ask
                except: pass
            '''if bcch == None:
                if str(valor[4]).lower() == 'x': cancelarVentaHB(valor[0])
                elif str(valor[4]).lower() == 'xx': cancelarTodo(inicio,fin)
                elif valor[4] == '-': 
                    enviarOrdenHB('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
                elif str(valor[4]).upper() == 'P': getPortfolioHB(hb, os.environ.get('account_id2474'),3)
                else: 
                    try: 
                        if shtData.range('AH'+str(int(valor[0]+1))).value: cancelarVentaHB(valor[0]) # CANCELA oreden venta anterior
                        enviarOrdenHB('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0]) # Vendo Ask
                    except: pass'''
            shtData.range('T'+str(int(valor[0]+1))).value = ''
        
def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    global reCompra, descubierto
    nombre = str(shtData.range(str(symbol)).value).split()
    precio = shtData.range(str(price)).value
    stock = stockU(int(celda+1))

    if len(nombre) == 2: symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[1]) # Es caucho

    elif len(nombre) > 2: 
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])   # Son bonos / acciones / letras
        if reCompra == True:
            precio -= ganancia * 2
            precio = round(precio, -1)
            print('Re-COMPRA ',end='')
            reCompra = False
    else : 
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs' # Son opciones
        if reCompra == True:
            if descubierto == False : 
                precio -= ganancia * 2
                print('COMPRA el DESCUBIERTO ', end='')
            else: 
                precio += ganancia * 2
                print('VENDE en DESCUBIERTO ', end='')
                descubierto = False
            precio = round(precio, 3)
            reCompra = False

    if tipo.lower() == 'buy': 
        try: 
            if esFinde == False and noMatriz == False:
                pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(size)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
            else: print('ES FINDE ',end=' ')
            print(f'//______/ BUY   + {int(size)} {symbol} // precio: {precio}') 
            if stock == 0: shtData.range('U'+str(int(celda+1))).value = int(size)
            else: shtData.range('U'+str(int(celda+1))).value += int(size)
        except: 
            shtData.range('Q'+str(int(celda+1))+':'+'R'+str(int(celda+1))).value = ''
            print(f'______/ ERROR en COMPRA. {symbol} // precio: {precio} // + {int(size)}')
            
    else: # VENTA
        try:
            if esFinde == False and noMatriz == False: 
                pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(size)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
            else: print('ES FINDE ',end=' ')
            print(f'//______/ SELL  - {int(size)} {symbol} // precio: {precio}')
            if stock == 0: shtData.range('U'+str(int(celda+1))).value = int(size)/-1
            else: shtData.range('U'+str(int(celda+1))).value -= int(size)
        except:
            shtData.range('S'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            print(f'______/ ERROR en VENTA. {symbol} // precio: {precio} // {int(size)}')

    if str(nombre[0]).upper() == 'GGAL' or str(nombre[0]).upper() == 'GGALD' or len(nombre) < 2 :
        shtData.range('W'+str(int(celda+1))).value = precio
    else: shtData.range('W'+str(int(celda+1))).value = (precio / 100)
    
def bullOpciones(nombre=str,cantidad=int,celda=int,nominalDescubierto=bool,stock=int):
    global ganancia, reCompra, descubierto
    hora = time.strftime("%H:%M:%S")
    disponible = stockU(celda+1)
    try:
        last = shtData.range('F'+str(int(celda+1))).value
        if not last or last == None or last == 'None': soloContinua()
        costo = shtData.range('W'+str(int(celda+1))).value 
        if not costo or costo == None or costo == 'None': soloContinua()
        nombre = str(shtData.range(str(nombre)).value).split()
        bid = shtData.range('C'+str(int(celda+1))).value
        ask = shtData.range('D'+str(int(celda+1))).value
        stop = shtData.range('X1').value
        ganancia = shtData.range('Z1').value
        if not ganancia: ganancia = 2
        digitos = len(str(int(bid)))

        if len(nombre) < 2: # Ingresa si son OPCIONES ///////////////////////////////////////////////////////////////////////////
            symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs' 
            if digitos >= 3: ganancia *= 10
            elif digitos > 1: ganancia *= 20
            else: ganancia /= 2
            if nominalDescubierto == False :
                if bid > abs(costo):                         
                    shtData.range('W'+str(int(celda+1))).value = bid
                if not stop and stock > 0:
                    if last <= abs(costo) - ganancia and bid >= last: 
                        print(f'//___/ SELL x STOP /___// - {cantidad} {nombre[0]} {bid} ',end=' ')
                        if esFinde == False and noMatriz == False:
                            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(cantidad)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
                        try: 
                            if disponible - abs(cantidad) == 0:
                                shtData.range('U'+str(int(celda+1))+':X'+str(int(celda+1))).value = ''
                                shtData.range('W'+str(int(celda+1))).value  = ''
                            else:
                                shtData.range('U'+str(int(celda+1))).value -= abs(cantidad)
                                shtData.range('W'+str(int(celda+1))).value = bid
                        except: pass

                        bid -= ganancia * 10
                        bid = round(bid, -1)
                        print(f'____/ BUY el STOP /___  + {cantidad} {nombre[0]} {bid}', hora)
                        if esFinde == False and noMatriz == False:
                            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(cantidad)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
                        
            else: # OPCION VENDIDA EN DESCUBIERTO
                if ask < abs(costo): 
                    shtData.range('W'+str(int(celda+1))).value = ask
                if not stop and stock < 0:
                    if last >= abs(costo) + ganancia and ask <= last: 
                        print(f'//___/ BUY x STOP /___// + {cantidad} {nombre[0]} {ask}',end=' ')
                        if esFinde == False and noMatriz == False:
                            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(cantidad)), price=float(ask),order_type=pyRofex.OrderType.LIMIT)
                        try: 
                            if disponible + abs(cantidad) == 0:
                                shtData.range('U'+str(int(celda+1))+':X'+str(int(celda+1))).value = ''
                                shtData.range('W'+str(int(celda+1))).value  = ''
                            else:
                                shtData.range('U'+str(int(celda+1))).value += abs(cantidad)
                                shtData.range('W'+str(int(celda+1))).value = ask
                        except: pass

                        ask += ganancia * 10
                        ask = round(ask, -1)
                        print(f'____/ SELL el STOP /___  + {cantidad} {nombre[0]} {ask}', hora)
                        if esFinde == False and noMatriz == False:
                            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(cantidad)), price=float(ask),order_type=pyRofex.OrderType.LIMIT)
    except: pass



def scalpingStop(nombre=str,cantidad=int,celda=int,nominalDescubierto=bool,auto=bool):

    nombre = str(shtData.range(str(nombre)).value).split()
    ganancia = shtData.range('Z1').value
    hacerScalping = shtData.range('X1').value
    disponible = stockU(celda+1)
    bid = shtData.range('C'+str(int(celda+1))).value
    ask = shtData.range('D'+str(int(celda+1))).value
    last = shtData.range('F'+str(int(celda+1))).value
    costoStop = shtData.range('V'+str(int(celda+1))).value
    costoStop = bid if costoStop == None else costoStop
    costo = shtData.range('W'+str(int(celda+1))).value
    costoX = shtData.range('X'+str(int(celda+1))).value
    costoX = costo if costoX == None else costoX
    disponibleARS = shtData.range('M1').value

    if len(nombre) < 2:
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
        try:
            digitosB = len(str(int(bid)))
            digitosA = len(str(int(ask)))
            if digitosB == 1 or digitosA == 1: 
                if bid < 1 or ask < 1: ganancia /= 2
            if digitosB == 2 or digitosA == 2: ganancia *= 2
            if digitosB == 3 or digitosA == 3: 
                if bid < 200 or ask < 200: ganancia *= 10
                else: ganancia *= 30
            if digitosB > 3 or digitosA > 3: ganancia *= 60
        except: costo = None

        if costo != None and nominalDescubierto == False:
            if bid > abs(costo) + ganancia:                         
                #ask += ganancia
                if hacerScalping: shtData.range('W'+str(int(celda+1))).value = bid
                else:
                    print(f'____ SCALPING vendo posicion comprada - {cantidad}  {nombre[0]}  {bid} ', end= '|')
                    if esFinde == False and noMatriz == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(cantidad)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
                    try: 
                        if disponible - abs(cantidad) == 0:
                            shtData.range('U'+str(int(celda+1))+':X'+str(int(celda+1))).value = ''
                        else:
                            shtData.range('U'+str(int(celda+1))).value -= cantidad
                            shtData.range('W'+str(int(celda+1))).value = bid
                    except: pass
                    bid -= ganancia
                    bid = round(bid, 2)
                    
                    print(f'____ Recompro posicion + {cantidad} {nombre[0]} {bid}', ' || ', hora)
                    if esFinde == False and noMatriz == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(cantidad)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
                
            elif last <= abs(costo) - ganancia and bid >= last: 
                print(f'//// ____ STOP vendo posicion comprada - {cantidad} {nombre[0]} {bid} ',end=' ')
                if esFinde == False and noMatriz == False:
                    pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(cantidad)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
                try: 
                    if disponible - abs(cantidad) == 0:
                        shtData.range('U'+str(int(celda+1))+':X'+str(int(celda+1))).value = ''
                    else:
                        shtData.range('U'+str(int(celda+1))).value -= cantidad
                        shtData.range('W'+str(int(celda+1))).value = bid
                except: pass
                bid -= ganancia / 2
                bid = round(bid, 2)
                print(f'____ Recompro posicion + {cantidad} {nombre[0]} {bid}',' || ' ,  hora)
                if esFinde == False and noMatriz == False:
                    pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(cantidad)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
                
        elif costo != None: # OPCION VENDIDA EN DESCUBIERTO
            if ask < abs(costo) - ganancia: 
                if hacerScalping: shtData.range('W'+str(int(celda+1))).value = ask
                else:
                    print(f'____ SCALPING compro posicion descubierta + {cantidad} {nombre[0]} {ask} ', end= '| ')
                    if esFinde == False and noMatriz == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(cantidad)), price=float(ask),order_type=pyRofex.OrderType.LIMIT)
                    try: 
                        if disponible - abs(cantidad) == 0:
                            shtData.range('U'+str(int(celda+1))+':X'+str(int(celda+1))).value = ''
                        else:
                            shtData.range('U'+str(int(celda+1))).value += cantidad
                            shtData.range('W'+str(int(celda+1))).value = ask
                    except: pass
                    ask += ganancia
                    ask = round(ask, 2)
                    print(f'____ Revendo posicion - {cantidad} {nombre[0]} {ask}', '|| ' , hora)
                    if esFinde == False and noMatriz == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(cantidad)), price=float(ask),order_type=pyRofex.OrderType.LIMIT)
                    
            elif last >= abs(costo) + ganancia and ask <= last: 
                print(f'//// ____ STOP compro posicion descubierta + {cantidad} {nombre[0]} {ask}',end=' ')
                if esFinde == False and noMatriz == False:
                    pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(cantidad)), price=float(ask),order_type=pyRofex.OrderType.LIMIT)
                try: 
                    if disponible + abs(cantidad) == 0:
                        shtData.range('U'+str(int(celda+1))+':X'+str(int(celda+1))).value = ''
                    else:
                        shtData.range('U'+str(int(celda+1))).value += cantidad
                        shtData.range('W'+str(int(celda+1))).value = ask
                except: pass
                ask += ganancia / 2
                ask = round(ask, 2)
                print(f'____ Revendo posicion + {cantidad} {nombre[0]} {ask}','|| ' ,  hora)
                if esFinde == False and noMatriz == False:
                    pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(cantidad)), price=float(ask),order_type=pyRofex.OrderType.LIMIT)

    else: # OPERACIONES CON BONOS
        cierreHora = False
        
        arsVenta = bid / 100 * cantidad

        if hora > '16:24:00' and str(nombre[2]).lower() == 'CI': cierreHora = True
        if hora > '16:56:00' and str(nombre[2]).lower() == '24hs': cierreHora = True
        if cierreHora == False and costo != None:
            dolar = False
            symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
            if nombre[0][-1:] == 'D' or nombre[0][-1:] == 'C':
                ganancia /= 300
                dolar = True
            try:
                digitosB = len(str(int(bid)))
                if digitosB == 3: 
                    ganancia /= 3
            except: costo = None

            # ENTRA EN STOP
            if last/100 <= abs(costoStop) - ganancia * 1.2 and bid/100 >= last/100:
                cantidad = disponible / 2
                print(f'//// ____ STOP vendo posicion comprada - {int(cantidad)} {nombre[0]} {nombre[2]} {int(bid)} ',end=' ') 
                if esFinde == False and noMatriz == False:
                    pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(cantidad)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
                try:
                    arsVenta = bid/ 100 * cantidad
                    shtData.range('M1').value += int(arsVenta)
                    if disponible - abs(cantidad) < 1:
                        shtData.range('U'+str(int(celda+1))+':X'+str(int(celda+1))).value = ''
                    else:
                        shtData.range('U'+str(int(celda+1))).value -= cantidad
                        shtData.range('V'+str(int(celda+1))).value = (bid / 100)
                except: pass
                shtData.range('V'+str(int(celda+1))).value = last / 100

                if dolar == 'SI':
                    bid -= ganancia * 200
                    bid = round(bid, 2)
                else: 
                    bid -= ganancia * 340
                    bid = round(bid, -1)
                print(f'____ Recompro posicion vendida + {cantidad} {nombre[0]} {nombre[2]} {int(bid)}', hora)
                if esFinde == False and noMatriz == False:
                    pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(cantidad)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
            
            if not auto and disponibleARS > 0: 
                if bid / 100 > abs(costoX) + ganancia / 3 and cantidad > 10: 
                    shtData.range('X'+str(int(celda+1))).value = bid / 100
                    if disponible < cantidad: cantidad = disponible - 1
                    ask += ganancia * 50
                    ask = round(ask, -1)
                    print(f'____ SCALPING rapido por subida. SELL posicion comprada - {cantidad}  {nombre[0]}  {ask} ')
                    if esFinde == False and noMatriz == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(cantidad)), price=float(ask),order_type=pyRofex.OrderType.LIMIT)
                    try: 
                        shtData.range('M1').value += int(arsVenta)
                        shtData.range('U'+str(int(celda+1))).value -= cantidad
                    except: pass

                elif bid / 100 > abs(costo) + ganancia / 2: 
                    shtData.range('V'+str(int(celda+1))).value = bid / 100
                    shtData.range('W'+str(int(celda+1))).value = bid / 100
                    modificaNominalesOperados = shtData.range('Y'+str(int(celda+1))).value + abs(cantidad) / 10
                    if modificaNominalesOperados > 999: pass
                    else: shtData.range('Y'+str(int(celda+1))).value += abs(cantidad) / 10
                    compraAuto(celda+1,dolar,'up')

                elif bid / 100 < abs(costo) - ganancia / 4: 
                    shtData.range('W'+str(int(celda+1))).value = bid / 100
                    modificaNominalesOperados = shtData.range('Y'+str(int(celda+1))).value - abs(cantidad) / 10
                    if modificaNominalesOperados < 10: pass
                    else: shtData.range('Y'+str(int(celda+1))).value -= abs(cantidad) / 10
                    compraAuto(celda+1,dolar,'down')


            elif bid / 100 >= abs(costo) + ganancia:   
                if hacerScalping : 
                    shtData.range('V'+str(int(celda+1))).value = bid / 100
                    shtData.range('W'+str(int(celda+1))).value = (bid / 100)
                else:
                    cantidad /= 2
                    if disponible < cantidad: cantidad = disponible - 1
                    bid += ganancia * 20
                    bid = round(bid, -1)
                    print(f'____ SCALPING vendo posicion comprada - {cantidad}  {nombre[0]}  {bid} ', end= ' | ')
                    if esFinde == False and noMatriz == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(cantidad)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
                    try: 
                        shtData.range('M1').value += int(arsVenta)
                        if disponible - abs(cantidad) < 1:
                            shtData.range('U'+str(int(celda+1))+':X'+str(int(celda+1))).value = ''
                        else:
                            shtData.range('U'+str(int(celda+1))).value -= cantidad
                            shtData.range('W'+str(int(celda+1))).value = (bid / 100)
                    except: pass
                    if auto:
                        if dolar == 'SI':
                            bid -= ganancia * 160
                            bid = round(bid, 2)
                        else: 
                            bid -= ganancia * 300
                            bid = round(bid, -1)
                        print(f'____ Recompro posicion vendida + {cantidad} {nombre[0]} {bid}', ' || ', hora)
                        if esFinde == False and noMatriz == False:
                            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(cantidad)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
                    else: print(' ______ ')

def compraTasa(celda,ladoCompra):
    compra = str(shtData.range('A'+str(int(celda))).value).split()
    if len(compra) < 2: pass
    else:
        if compra[2] == 'CI': 
            vende = str(shtData.range('M'+str(int(celda))).value).split()
            if len(vende) < 2: pass
            else:
                if vende[2] == '24hs':
                    nominales = cantidadAuto(celda)
                    gastos = shtData.range('AB1').value
                    buy = shtData.range(str(ladoCompra)+str(int(celda))).value
                    symbol = "MERV - XMEV - " + str(compra[0]) + ' - ' + str(compra[2])
                    print(f'___/ compra TASA + {int(nominales)} {compra[0]} {compra[2]} {buy} ',end=' | ')
                    if esFinde == False and noMatriz == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(buy),order_type=pyRofex.OrderType.LIMIT)
                    
                    sell = shtData.range('I'+str(int(celda))).value
                    symbol = "MERV - XMEV - " + str(vende[0]) + ' - ' + str(vende[2])
                    print(f'___/ - {int(nominales)} {vende[0]} {vende[2]} {sell}')

                    if esFinde == False and noMatriz == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(sell),order_type=pyRofex.OrderType.LIMIT)
                    
                    try: shtData.range('M1').value -= (buy / 100) * (1 + gastos) * nominales
                    except: pass
                    
                else: pass
        else: pass
        shtData.range('Q'+str(int(celda))).value = ''


# OPCIONES: estrategia tipo MARIPOSA ------------------------
def baseEjercible(celda=int):
    base = shtData.range('F61').value
    ticker = shtData.range('A'+str(int(celda+1))).value
    ticker = str(ticker).split()
    ticker = ticker[0]
    shtData.range('Q'+str(int(celda+1))).value = ''
    if str(ticker[3:4]).upper() == 'C': 
        ticker = ticker[4:9]
        try:
            if int(ticker) - float(base) < 0:
                print(ticker, 'Es un CALL EJERCIBLE. Mariposa NO permitida para operar')
            else: verificaMariposa(celda)
        except:
            ticker = ticker[:4]
            if int(ticker) - float(base) < 0:
                print(ticker, 'Es un CALL EJERCIBLE. Mariposa NO permitida para operar')
            else: verificaMariposa(celda)
    else: 
        ticker = ticker[4:9]
        try:
            if float(base) - int(ticker) < 0:
                print(ticker, 'Es un PUT EJERCIBLE. Mariposa NO permitida para operar')
            else: verificaMariposa(celda)
        except:
            ticker = ticker[:4]
            if float(base) - int(ticker) < 0:
                print(ticker, 'Es un PUT EJERCIBLE. Mariposa NO permitida para operar')
            else: verificaMariposa(celda)

def verificaMariposa(celda=int):
    activo = shtData.range('AB30').value
    if str(activo).upper() == 'B':
        valor = shtData.range('Z'+str(int(celda+1))).value
        try:
            if valor > 10: 
                mariposas(celda)
            else: 
                shtData.range('Q'+str(int(celda+1))).value = ''
                print('Cancela mariposa por vaja rentabilidad o valor negativo: ', valor)
        except: 
            print('Error en valor de mariposa: ', valor)
            shtData.range('Q'+str(int(celda+1))).value = ''

def mariposas(celda=int):
    try: 
        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range('C'+str(int(celda+1))).value
        nominales = 1 # shtData.range('Y'+str(int(celda+1))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
        print(f'//___/ MARIPOSA AUT /___// - {int(nominales)} {nombre[0]} {precio} ', end='|')
        if esFinde == False and noMatriz == False: 
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)

        nombre = str(shtData.range('A'+str(int(celda))).value).split()
        precio = shtData.range('D'+str(int(celda))).value
        #nominales = shtData.range('Y'+str(int(celda))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
        print(f' + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False and noMatriz == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)

        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range('C'+str(int(celda+1))).value
        #nominales = shtData.range('Y'+str(int(celda+1))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
        print(f' - {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False and noMatriz == False: 
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)

        nombre = str(shtData.range('A'+str(int(celda+2))).value).split()
        precio = shtData.range('D'+str(int(celda+2))).value
        #nominales = shtData.range('Y'+str(int(celda+2))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
        print(f' + {nominales} {nombre[0]} {precio} ')
        if esFinde == False and noMatriz == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except:
        print('Falla al intentar una mariposa con: ', shtData.range('A'+str(celda)).value)
        shtData.range('Q'+str(celda)).value = ''

def cerrarMariposa(celda=int):
    nominales = 1
    try: 
        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range('D'+str(int(celda+1))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
        print(f'___/ cierra MARIPOSA AUT /___ + {int(nominales)} {nombre[0]} {precio} ', end='|')
        if esFinde == False and noMatriz == False: 
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)

        nombre = str(shtData.range('A'+str(int(celda))).value).split()
        precio = shtData.range('C'+str(int(celda))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
        print(f' - {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False and noMatriz == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)

        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range('D'+str(int(celda+1))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
        print(f' + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False and noMatriz == False: 
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)

        nombre = str(shtData.range('A'+str(int(celda+2))).value).split()
        precio = shtData.range('C'+str(int(celda+2))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
        print(f' - {nominales} {nombre[0]} {precio} ')
        if esFinde == False and noMatriz == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except:
        print('Falla al intentar una mariposa con: ', shtData.range('A'+str(celda)).value)
        shtData.range('Q'+str(celda)).value = ''

# FIN de MARIPOSAS ---------------------------------------------

vuelta = 0
vueltaPortfolio = 0
obtenerSaldoMatriz(str(os.environ.get('account')))

while True:
    hora = time.strftime("%H:%M:%S")
    traeAdr = shtData.range('R1').value
    #vetaHB = shtData.range('U1').value
    #bcch = shtData.range('V1').value
    
    try:
        preparar =  shtData.range('A1').value
        if not preparar: 
            shtData.range('A1').value = 'symbol'
            rollNuevo()
            #preparaRulo(preparar)
    except: 
        print('error al preparar los Rulos ')
        shtData.range('A1').value = 'symbol'


    buscoOperaciones(rangoDesde,rangoHasta)
    
    if hora >= '17:00:30': 
        if noMatriz == False or esFinde == False:
            if hora >= '17:01:30': pass
            else:
                print(time.strftime("%H:%M:%S"), 'Mercado local cerrado')
                shtData.range('Q1').value = 'PRC'
                shtData.range('R1').value = 'ADR'
                shtData.range('S1').value = 'D'
                #shtData.range('T1').value = 'ROLL'
                shtData.range('U1').value = 'veta'
                shtData.range('V1').value = 'bcch'
                shtData.range('W1').value = 'AUTO'
                shtData.range('X1').value = 'SCP'

                shtData.range('Z1').value = 0.5
                pyRofex.close_websocket_connection()
                hb.online.disconnect()
    else:
        if not shtData.range('Q1').value:
            try:
                if noMatriz == False and esFinde == False: 
                    #shtData.range('A35').options(index=False, headers=False).value = df_datos # CON OPCIONES INCLUIDAS
                    shtData.range('A69').options(index=False, headers=False).value = df_datos  
                    shtData.range('W35').value = hora

            except: print('Hubo un error al actualizar excel', hora)

        if vueltaPortfolio > 30: 
            vueltaPortfolio = 0
            getPortfolioHB(hbVETA, str(os.environ.get('account_id')), 2) 
            #if not bcch: getPortfolioHB(hb, str(os.environ.get('account_id2474')), 3)
            #else: print('BCCH: s/d')
        else: vueltaPortfolio += 1
            
        if traeAdr == None:
            try:
                if vuelta > 4: 
                    traerADR()
                    vuelta = 0
                else: vuelta += 1
            except: pass
    #shtOperaciones.range('AI63').options(index=False, headers=False).value = operaciones
    time.sleep(2)



