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
shtData.range('Z1').value = 0.5
rangoDesde = '2'
rangoHasta = '60'
reCompra = False
esFinde = False
    
def diaLaboral():
    global esFinde
    hoyEs = time.strftime("%A")
    if hoyEs == 'Saturday' or hoyEs == 'Sunday':
        esFinde = True
def loguinHB():
    from pyhomebroker import HomeBroker  
    global hb
    try:
        hb = HomeBroker(int('284'))
        hb.auth.login(
            dni='26386662', 
            user='piolifl',  
            password='Bordame02',
            raise_exception=True)
        print("online VETA HB  cuenta: 47352", time.strftime("%H:%M:%S"))
    except: 
        print("    NO se pudo loguear en VETA HOME BROKER 47352    * ", time.strftime("%H:%M:%S"))
        pass

diaLaboral()

if esFinde == False:
    pyRofex._set_environment_parameter("url", "https://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)
    pyRofex._set_environment_parameter("ws", "wss://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)
    pyRofex.initialize(user="20263866623", password="Bordame01!", account="47352", environment=pyRofex.Environment.LIVE)
    print("online VETA OMS cuenta: 47352", end=' || ')
    loguinHB()
else: print('FIN DE SEMANA, no se actualizan los precios locales y no se envian ordenes al broker.')

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

# tickers = pd.concat([opciones, acc, bonos,cedear,ons,letras, caucion ])

listLength = len(acc) + 31 + len(opciones)
allLength = 28 + len(tickers) - len(acc)  - len(caucion)

if esFinde == False:
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

if esFinde == False:
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
        shtData.range('U2:'+'U'+str(rangoHasta)).value = ''
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

def cancelaCompraHB(celda):
    orderC = shtData.range('AC'+str(int(celda+1))).value
    if not orderC or orderC == None or orderC == 'None' or orderC == '': orderC = 0

    if esFinde == False: 
        try: 
            hb.orders.cancel_order('47352',int(orderC))
            print(f"/// Cancelada Compra : {int(orderC)} ",end='\t')
        except: 
            print(f'Error al cancelar COMPRA {orderC} con HB')
    try: shtData.range('V'+str(int(celda+1))).value -= shtData.range('AB'+str(int(celda+1))).value
    except: pass
    shtData.range('AB'+str(int(celda+1))+':'+'AC'+str(int(celda+1))).value = ''
        
def cancelarVentaHB(celda):
    orderV = shtData.range('AE'+str(int(celda+1))).value
    if not orderV or orderV == None or orderV == 'None' or orderV == '': orderV = 0
    if esFinde == False: 
        try:
            hb.orders.cancel_order('47352',int(orderV))
            print(f"/// Cancelada Venta  : {int(orderV)} " ,end='\t')
        except: 
            print(f'Error al cancelar VENTA {orderV} con HB')
    try: shtData.range('V'+str(int(celda+1))).value += shtData.range('AD'+str(int(celda+1))).value
    except: pass
    shtData.range('AD'+str(int(celda+1))+':'+'AE'+str(int(celda+1))).value = ''

def soloContinua():
    pass

def namesArs(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA37D'+plazo
    elif nombre[:2] == 'BP': return 'BPOA7'+plazo
    elif nombre[:2] == 'KO': return 'KO'+plazo
    elif nombre[:2] == 'GOGL': return 'GOOGL'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') and (nombre[3:4] == 'D' or nombre[3:4] == 'C'):
        if (nombre[1:2] == 'F' or nombre[1:2] == 'Y'): return nombre[:1]+'20'+nombre[1:3]+plazo
        if (nombre[1:2] == 'M'): return nombre[:1]+'31'+nombre[1:3]+plazo
        if (nombre[1:2] == 'N'): return nombre[:1]+'29'+nombre[1:3]+plazo
        if (nombre[1:2] == 'J'): return nombre[:1]+'18'+nombre[1:3]+plazo
        if (nombre[1:2] == 'G'): return nombre[:1]+'30'+nombre[1:3]+plazo
        if (nombre[1:2] == 'O'): return nombre[:1]+'14'+nombre[1:3]+plazo
        if (nombre[1:2] == 'E'): return nombre[:1]+'31'+nombre[1:3]+plazo
        else: return nombre[:1]+'18'+nombre[1:3]+plazo
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

def cargoXplazo(dicc):
    mejorMep = dicc['mepCI'][0]
    mejorMep24 = dicc['mep24'][0]
    mepArs = namesMep(dicc['arsCImep'][0],' - CI')
    mepArs24 = namesMep(dicc['ars24mep'][0],' - 24hs')
    mepCcl = namesMep(dicc['cclCI'][0],' - CI')
    mepCcl24 = namesMep(dicc['ccl24'][0],' - 24hs')

    '''if mejorMep == 'AL30D - ci': shtData.range('A2:A5').value = ''
    else: 
        shtData.range('A2').value = mejorMep
        shtData.range('A3').value = 'AL30D - CI'
        shtData.range('A4').value = 'AL30 - CI'
        shtData.range('A5').value = namesArs(dicc['mepCI'][0],' - CI')

    if mejorMep24 == 'AL30D - 24hs': shtData.range('A6:A9').value = ''
    else: 
        shtData.range('A6').value = mejorMep24
        shtData.range('A7').value = 'AL30D - 24hs'
        shtData.range('A8').value = 'AL30 - 24hs'
        shtData.range('A9').value = namesArs(dicc['mep24'][0],' - 24hs')'''
    
    if mejorMep == mepArs: shtData.range('A10:A13').value = ''
    else:
        shtData.range('A10').value = dicc['arsCImep'][0]
        shtData.range('A11').value = mepArs
        shtData.range('A12').value = mejorMep
        shtData.range('A13').value = namesArs(dicc['mepCI'][0],' - CI')

    if mejorMep24 == mepArs24: shtData.range('A14:A17').value = ''
    else:
        shtData.range('A14').value = dicc['ars24mep'][0]
        shtData.range('A15').value = mepArs24
        shtData.range('A16').value = mejorMep24
        shtData.range('A17').value = namesArs(dicc['mep24'][0],' - 24hs')

    if mejorMep == mepCcl:  shtData.range('A18:A21').value = ''
    else:
        shtData.range('A18').value = mejorMep
        shtData.range('A19').value = mepCcl
        shtData.range('A20').value = dicc['cclCI'][0]
        shtData.range('A21').value = namesCcl(dicc['mepCI'][0],' - CI')

    if mejorMep24 == mepCcl24: shtData.range('A22:A25').value = ''
    else:
        shtData.range('A22').value = mejorMep24
        shtData.range('A23').value = namesCcl(dicc['mep24'][0],' - 24hs')
        shtData.range('A24').value = dicc['ccl24'][0]
        shtData.range('A25').value = mepCcl24
    shtData.range('A1').value = 'symbol'

def preparaRulo():
    celda,pesos,dolar = listLength,1000,0.01
    tikers = {'cclCI':['',dolar],'ccl24':['',dolar],'mepCI':['',dolar],'mep24':['',dolar],'arsCIccl':['',pesos],'ars24ccl':['',pesos],'arsCImep':['',pesos],'ars24mep':['',pesos]}
    for valor in shtData.range('A'+str(listLength)+':A'+str(allLength)).value:
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
    cargoXplazo(tikers)

def traerADR():
    valorAdr = yf.download(['GGAL','YPF'],period='1d',interval='1d',auto_adjust=False)['Close'].values
    shtData.range('Z61').value = valorAdr[0][0]
    shtData.range('Z63').value = valorAdr[0][1]
    shtData.range('Y62').value = time.strftime("%H:%M:%S")
    
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

def cantidadAuto(nroCelda):
    cantidad = shtData.range('Y'+str(int(nroCelda))).value
    cantidad = 0 if cantidad == None else cantidad
    return abs(int(cantidad))

def stokDisponible(nroCelda):
    stok = shtData.range('U'+str(int(nroCelda))).value
    stok = 0 if not stok or stok == None or stok == 'None' else stok        
    return int(stok)

def hacerTasa(celda,ladoCompra,ladoVenta):
    compra = str(shtData.range('A'+str(int(celda+1))).value).split()
    if len(compra) < 2: pass
    else:
        if compra[2] == 'CI': 
            vende = str(shtData.range('A'+str(int(celda))).value).split()
            if len(vende) < 2: pass
            else:
                if vende[2] == '24hs':
                    nominales = cantidadAuto(celda)
                    stock = stokDisponible(celda)
                    bid = shtData.range(str(ladoCompra)+str(int(celda+1))).value
                    symbol = "MERV - XMEV - " + str(compra[0]) + ' - ' + str(compra[2])
                    print(f'//___/ BUY TASA + {int(nominales)} {compra[0]} {compra[2]} {bid} ',end=' | ')
                    if esFinde == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)

                    ask = shtData.range(str(ladoVenta)+str(int(celda))).value + 20
                    symbol = "MERV - XMEV - " + str(vende[0]) + ' - ' + str(vende[2])
                    if abs(stock) < nominales: nominales = abs(stock)
                    print(f'___/ SELL - {int(nominales)} {vende[0]} {vende[2]} {ask}')
                    if esFinde == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(ask),order_type=pyRofex.OrderType.LIMIT)
                else: pass
        else: pass

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
            if esFinde == False:
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
    try: 
        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range(str(lado)+str(int(celda+1))).value
        if len(nombre) < 2: 
            symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
            if str(tipo).upper() == 'BUY': shtData.range('W'+str(int(celda+1))).value = precio
        else: 
            symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
            if str(tipo).upper() == 'BUY': shtData.range('W'+str(int(celda+1))).value = precio / 100
        
        if str(tipo).upper() == 'BUY':
            print(f'//___/ BUY  /___// + {nominales} {nombre[0]} // precio: {precio}')
            if esFinde == False:
                pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
        else:
            if abs(stock) < nominales: nominales = abs(stock)
            print(f'//___/ SELL /___// - {nominales} {nombre[0]} // precio: {precio}')
            if esFinde == False:
                pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
            
    except: pass

def sellRoll(celda=int):
    try:
        nominales = cantidadAuto(celda)
        nombre = str(shtData.range('A'+str(int(celda))).value).split()
        precio = shtData.range('C'+str(int(celda))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f'//___/ RULO AUT /___// - {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False: 
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

    try:
        nominales = cantidadAuto(celda+1)
        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range('D'+str(int(celda+1))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

    try:
        nominales = cantidadAuto(celda+2)
        nombre = str(shtData.range('A'+str(int(celda+2))).value).split()
        precio = shtData.range('C'+str(int(celda+2))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' - {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

    try:
        nominales = cantidadAuto(celda+3)
        nombre = str(shtData.range('A'+str(int(celda+3))).value).split()
        precio = shtData.range('D'+str(int(celda+3))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' + {nominales} {nombre[0]} {precio}')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

def buyRoll(celda=int):
    try:
        nominales = cantidadAuto(celda)
        nombre = str(shtData.range('A'+str(int(celda))).value).split()
        precio = shtData.range('D'+str(int(celda))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f'//___/ BUY ROLL /___ + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False: 
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

    try:
        nominales = cantidadAuto(celda+1)
        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range('C'+str(int(celda+1))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' - {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

    try:
        nominales = cantidadAuto(celda+2)
        nombre = str(shtData.range('A'+str(int(celda+2))).value).split()
        precio = shtData.range('D'+str(int(celda+2))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

    try:
        nominales = cantidadAuto(celda+3)
        nombre = str(shtData.range('A'+str(int(celda+3))).value).split()
        precio = shtData.range('C'+str(int(celda+3))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' - {nominales} {nombre[0]} {precio} ___//')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

def buyRollPlus(celda=int):
    try:
        nominales = cantidadAuto(celda)
        nombre = str(shtData.range('A'+str(int(celda))).value).split()
        precio = shtData.range('D'+str(int(celda))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f'//___/ BUY ROLL PLUS /___ + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False: 
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

    try:
        nominales = cantidadAuto(celda+1)
        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range('C'+str(int(celda+1))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' - {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

    try:
        nominales = cantidadAuto(celda+2)
        nombre = str(shtData.range('A'+str(int(celda+2))).value).split()
        precio = shtData.range('D'+str(int(celda+2))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

    try:
        nominales = cantidadAuto(celda+3)
        nombre = str(shtData.range('A'+str(int(celda+3))).value).split()
        precio = shtData.range('D'+str(int(celda+3))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' - {nominales} {nombre[0]} {precio} ___//')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

def buyRollMenos(celda=int):
    try:
        nominales = cantidadAuto(celda)
        nombre = str(shtData.range('A'+str(int(celda))).value).split()
        precio = shtData.range('D'+str(int(celda))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f'//___/ BUY ROLL MENOS y PLUS /___ + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False: 
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

    try:
        nominales = cantidadAuto(celda+1)
        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range('C'+str(int(celda+1))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' - {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

    try:
        nominales = cantidadAuto(celda+2)
        nombre = str(shtData.range('A'+str(int(celda+2))).value).split()
        precio = shtData.range('C'+str(int(celda+2))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass

    try:
        nominales = cantidadAuto(celda+3)
        nombre = str(shtData.range('A'+str(int(celda+3))).value).split()
        precio = shtData.range('D'+str(int(celda+3))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        print(f' - {nominales} {nombre[0]} {precio} ___//')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except: pass


def roll():
    celda = 2
    for i in shtData.range('O2:O25').value:
            if str(i).upper() == 'R':
                if celda==2 or celda==6 or celda==8 or celda==10 or celda==14 or celda==18 or celda==22:
                    sellRoll(celda)
            celda += 1
    shtData.range('T1').value = 'ROLL'

def posicionRulo(celda):
    if celda==2 or celda==6 or celda==8 or celda==10 or celda==14 or celda==18 or celda==22:
        return 'ok'

def buscoOperaciones(inicio,fin):
    hora = time.strftime("%H:%M:%S")
    if not shtData.range('T1').value: roll() # RULO AUTOMATICO activado por columna O
    for valor in shtData.range('P'+str(inicio)+':'+'U'+str(fin)).value:
        try:
            if not valor[5] or valor[5] == 0 or hora <= '11:01:00': pass  
            else: 
                nominalDescubierto = True if valor[5] < 0 else False
                cantidad = cantidadAuto(valor[0]+1)
                if cantidad != 0:
                    trailingStop('A'+str((int(valor[0]+1))),cantidad,int(valor[0]),nominalDescubierto,valor[5])
        except: pass

        if valor[1]: # # Columna Q en el excel /////////////////////////////////////////////////////////////////////////////////
            if str(valor[1]).lower() == 'r' and posicionRulo(valor[0]+1) == 'ok': buyRoll(valor[0]+1)
            elif str(valor[1]).lower() == 'r+' and posicionRulo(valor[0]+1) == 'ok': buyRollPlus(valor[0]+1)
            elif str(valor[1]).lower() == 'r-' and posicionRulo(valor[0]+1) == 'ok': buyRollMenos(valor[0]+1)
            elif str(valor[1]).lower() == 'p': getPortfolioHB(hb,'47352',1)
            elif str(valor[1]).lower() == 'm': baseEjercible(valor[0])
            elif str(valor[1]).lower() == 'sm': verificaMariposa(valor[0])
            elif str(valor[1]).lower() == 't': hacerTasa(valor[0],'C','D')
            elif valor[1] == '+': 
                cantidad = cantidadAuto(valor[0]+1)
                operacionRapida(valor[0],'C','BUY',valor[5], cantidad)
            elif str(valor[1]).lower() == 's': 
                cantidad = cantidadAuto(valor[0]+1)
                scalping(valor[0],'C','BUY', valor[5], cantidad)
            else: 
                try: enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0]) # Compra Bid
                except: shtData.range('Q'+str(int(valor[0]+1))).value = ''
            shtData.range('Q'+str(int(valor[0]+1))).value = ''

        if valor[2]: #  Columna R en el excel //////////////////////////////////////////////////////////////////////////////////
            if str(valor[2]).lower() == 'r' and posicionRulo(valor[0]+1) == 'ok': buyRoll(valor[0]+1)
            elif str(valor[2]).lower() == 'p': getPortfolioHB(hb,'47352',1)
            elif str(valor[2]).lower() == 't': hacerTasa(valor[0],'D','D')
            elif valor[2] == '+': 
                cantidad = cantidadAuto(valor[0]+1)
                operacionRapida(valor[0],'D','BUY', valor[5], cantidad)
            elif str(valor[2]).lower() == 's': 
                cantidad = cantidadAuto(valor[0]+1)
                scalping(valor[0],'D','BUY', valor[5], cantidad)
            else: 
                try: enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0]) # Compra Ask
                except: shtData.range('R'+str(int(valor[0]+1))).value = ''
            shtData.range('R'+str(int(valor[0]+1))).value = ''
        
        if valor[3]: # Columna S en el excel ///////////////////////////////////////////////////////////////////////////////////
            if str(valor[3]).lower() == 'r' and posicionRulo(valor[0]+1) == 'ok': buyRoll(valor[0]+1)
            elif str(valor[3]).lower() == 'p': getPortfolioHB(hb,'47352',1)
            elif valor[3] == '-': 
                cantidad = cantidadAuto(valor[0]+1)
                operacionRapida(valor[0],'C','SELL', valor[5], cantidad)
            elif str(valor[3]).lower() == 's': 
                cantidad = cantidadAuto(valor[0]+1)
                scalping(valor[0],'C','SELL', valor[5], cantidad)
            else: 
                try: enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0]) # Vendo Bid
                except: shtData.range('S'+str(int(valor[0]+1))).value = ''
            shtData.range('S'+str(int(valor[0]+1))).value = ''

        if valor[4]: # Columna T en el excel //////////////////////////////////////////////////////////////////////////////////
            if str(valor[4]).lower() == 'r' and posicionRulo(valor[0]+1) == 'ok': buyRoll(valor[0]+1)
            elif str(valor[4]).lower() == 'p': getPortfolioHB(hb,'47352',1)
            elif valor[4] == '-': 
                cantidad = cantidadAuto(valor[0]+1)
                operacionRapida(valor[0],'D','SELL', valor[5], cantidad)
            elif str(valor[4]).lower() == 's': 
                cantidad = cantidadAuto(valor[0]+1)
                scalping(valor[0],'D','SELL', valor[5], cantidad)
            else: 
                try: enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0]) # Vendo Ask
                except: shtData.range('T'+str(int(valor[0]+1))).value = ''
            shtData.range('T'+str(int(valor[0]+1))).value = ''

def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    global reCompra, descubierto
    nombre = str(shtData.range(str(symbol)).value).split()
    precio = shtData.range(str(price)).value
    stock = stokDisponible(int(celda+1))

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
            if esFinde == False:
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
            if esFinde == False: 
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
    
def trailingStop(nombre=str,cantidad=int,nroCelda=int,nominalDescubierto=bool,stock=int):
    global ganancia, reCompra, descubierto
    dolar = 'NO'
    try:
        last = shtData.range('F'+str(int(nroCelda+1))).value
        if not last or last == None or last == 'None': soloContinua()
        costo = shtData.range('W'+str(int(nroCelda+1))).value 
        if not costo or costo == None or costo == 'None': soloContinua()
        apertura = shtData.range('X'+str(int(nroCelda+1))).value
        nombre = str(shtData.range(str(nombre)).value).split()
        stop = shtData.range('X1').value
        r = shtData.range('W1').value
        if len(nombre) < 2: symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs' 
        else : symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])
        bid = shtData.range('C'+str(int(nroCelda+1))).value
        ask = shtData.range('D'+str(int(nroCelda+1))).value
        
        if len(nombre) < 2: # Ingresa si son OPCIONES ///////////////////////////////////////////////////////////////////////////
            ganancia = shtData.range('Z1').value * 10
            if not ganancia: ganancia = 2
            if nominalDescubierto == False :
                if bid >= abs(costo) + ganancia:                         
                    shtData.range('W'+str(int(nroCelda+1))).value = bid

                #shtData.range('V'+str(int(nroCelda+1))).value = round((bid-apertura)*abs(stock)*100,2)

                if not stop and stock > 0 and cantidad > 0:
                    if last <= abs(costo) - (ganancia) and bid >= last: 
                        if not r: print(f'//___/ SELL STOP /___// - {cantidad} {nombre[0]} // precio: {bid} ',end=' ')
                        else: print(f'//___/ SELL STOP /___// - {cantidad} {nombre[0]} // precio: {bid} ')
                        if esFinde == False:
                            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(cantidad)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
                        shtData.range('W'+str(int(nroCelda+1))).value = shtData.range('C'+str(int(nroCelda+1))).value
                        if not r: 
                            bid -= ganancia 
                            bid = round(bid, -1)
                            print(f'____/ BUY STOP /___  + {cantidad} {nombre[0]} // precio: {bid}')
                            if esFinde == False:
                                pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(cantidad)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
                        try: 
                            shtData.range('U'+str(int(nroCelda+1))).value -= abs(cantidad)
                        except: pass

            else: # OPCION VENDIDA EN DESCUBIERTO

                if ask <= abs(costo) - ganancia / 2: 
                    shtData.range('W'+str(int(nroCelda+1))).value = ask

                #shtData.range('V'+str(int(nroCelda+1))).value = round((ask-apertura)*abs(stock)*-100,2)
                if not stop and stock < 0 and cantidad < 0:  
                    if last >= abs(costo)+ (ganancia) and ask <= last: 
                        if not r: print(f'//___/ BUY STOP /___// + {cantidad} {nombre[0]} // precio: {ask}',end=' ')
                        else: print(f'//___/ BUY STOP /___// + {cantidad} {nombre[0]} // precio: {ask} ')
                        if esFinde == False:
                            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(cantidad)), price=float(ask),order_type=pyRofex.OrderType.LIMIT)
                        shtData.range('W'+str(int(nroCelda+1))).value = shtData.range('D'+str(int(nroCelda+1))).value
                        if not r: 
                            ask += ganancia 
                            ask = round(ask, -1)
                            print(f'____/ SELL STOP /___  + {cantidad} {nombre[0]} // precio: {ask}')
                            if esFinde == False:
                                pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(cantidad)), price=float(ask),order_type=pyRofex.OrderType.LIMIT)
                        try: 
                            shtData.range('U'+str(int(nroCelda+1))).value += abs(cantidad)
                        except: pass
        
        else: # Ingresa si son BONOS / LETRAS / ON / CEDEARS ////////////////////////////////////////////////////////////////////
            if time.strftime("%H:%M:%S") > '16:24:00' and str(nombre[2]).lower() == 'CI': 
                if time.strftime("%H:%M:%S") > '17:01:00': pass
                else: pass
            if time.strftime("%H:%M:%S") > '16:56:00' and str(nombre[2]).lower() == '24hs': 
                if time.strftime("%H:%M:%S") > '17:01:00': pass
                else: soloContinua()

            ganancia = shtData.range('Z1').value
            if not ganancia: ganancia = 0.5

            if nombre[0][-1:] == 'D' or nombre[0][-1:] == 'C':
                ganancia /= 100
                dolar = 'SI'

            if bid / 100 > abs(costo) + ganancia / 2: 
                shtData.range('W'+str(int(nroCelda+1))).value = bid/100

            #shtData.range('V'+str(int(nroCelda+1))).value = round(((bid/100)-apertura)*stock,2)
                    
            if not stop and cantidad > 0:
                if last/100 <= abs(costo) - ganancia and bid/100 >= last/100:
                    if not r: 
                        print(f'//___/ SELL x STOP /___ - {cantidad} {nombre[0]} // precio: {bid} ',end=' ')
                    else: 
                        print(f'//___/ SELL x STOP /___ - {cantidad} {nombre[0]} // precio: {bid} ') 
                    if esFinde == False:
                        pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(cantidad)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
                    try:
                        shtData.range('U'+str(int(nroCelda+1))).value -= abs(cantidad) 
                    except: pass
                    shtData.range('W'+str(int(nroCelda+1))).value = shtData.range('C'+str(int(nroCelda+1))).value / 100
                    if not r: 
                        if dolar == 'SI': bid -= ganancia * 10
                        else: bid -= ganancia * 200
                        bid = round(bid, -1)
                        print(f'____/ BUY x STOP /___  + {cantidad} {nombre[0]} // precio: {bid}')
                        if esFinde == False:
                            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(cantidad)), price=float(bid),order_type=pyRofex.OrderType.LIMIT)
    except: pass

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
        valor = shtData.range('AB'+str(int(celda+1))).value
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
        if esFinde == False: 
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)

        nombre = str(shtData.range('A'+str(int(celda))).value).split()
        precio = shtData.range('D'+str(int(celda))).value
        #nominales = shtData.range('Y'+str(int(celda))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
        print(f' + {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)

        nombre = str(shtData.range('A'+str(int(celda+1))).value).split()
        precio = shtData.range('C'+str(int(celda+1))).value
        #nominales = shtData.range('Y'+str(int(celda+1))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
        print(f' - {nominales} {nombre[0]} {precio} ', end='|')
        if esFinde == False: 
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)

        nombre = str(shtData.range('A'+str(int(celda+2))).value).split()
        precio = shtData.range('D'+str(int(celda+2))).value
        #nominales = shtData.range('Y'+str(int(celda+2))).value
        symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs'
        print(f' + {nominales} {nombre[0]} {precio} ')
        if esFinde == False:
            pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(nominales)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
    except:
        print('Falla al intentar una mariposa con: ', shtData.range('A'+str(celda)).value)
        shtData.range('Q'+str(celda)).value = ''
# FIN de MARIPOSAS ---------------------------------------------

vuelta = 0
vueltaPortfolio = 0

while True:
    hora = time.strftime("%H:%M:%S")
    try:
        if not shtData.range('A1').value: preparaRulo()
    except:
        shtData.range('A1').value = 'symbol'

    buscoOperaciones(rangoDesde,rangoHasta)

    if hora > '22:00:30' : 
        print(time.strftime("%H:%M:%S"), 'Mercado local cerrado')
        shtData.range('Q1').value = 'PRECIOS'
        shtData.range('S1').value = 'ADR'
        shtData.range('T1').value = 'ROLL'
        shtData.range('W1').value = 'R'
        shtData.range('X1').value = 'STOP'
        shtData.range('Z1').value = 0.5
        if esFinde == False: 
            pyRofex.close_websocket_connection()
            hb.online.disconnect()
    else:
        if not shtData.range('Q1').value:
            try:
                if esFinde == False: shtData.range('A30').options(index=False, headers=False).value = df_datos   
            except: print('Hubo un error al actualizar excel')

            
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
