
import pyRofex
import xlwings as xw
import time , math
import pandas as pd
import yfinance as yf


wb = xw.Book('D:\\pyHomeBroker\\epgb_pyRofex.xlsb')
shtTickers = wb.sheets('pyRofex')
shtData = wb.sheets('MATRIZ OMS')
#shtData = wb.sheets('Hoja1')
#shtOperaciones = wb.sheets('MATRIZ OMS')

shtData.range('Q1').value = 'PRECIOS'
shtData.range('S1').value = 'ADR'
shtData.range('W1').value = 'TRAILING'
shtData.range('X1').value = 'STOP'
shtData.range('Y1').value = 'VETA OMS'
shtData.range('Z1').value = 0.001
rangoDesde = '2'
rangoHasta = '64'
shtData.range('W'+str(rangoDesde)+':'+'W'+str(rangoHasta)).value = '' 
hoyEs = time.strftime("%A")
    

def diaLaboral():
    if hoyEs == 'Saturday' or hoyEs == 'Sunday':
        return 'Fin de semana'

if diaLaboral():
    print('Es FIN DE SEMANA, no se actualizan los precios locales y no se envian ordenes al broker.')
    esFinde = True
else: 
    esFinde = False

pyRofex._set_environment_parameter("url", "https://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)
pyRofex._set_environment_parameter("ws", "wss://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)

pyRofex.initialize(user="20263866623", password="Bordame01!", account="47352", environment=pyRofex.Environment.LIVE)
print(("online"), time.strftime("%H:%M:%S"))

    
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
allLength = 30 + len(tickers)  - len(caucion)
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

    if mejorMep == 'AL30D - ci': shtData.range('A2:A5').value = ''
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
        shtData.range('A9').value = namesArs(dicc['mep24'][0],' - 24hs')
    
    if mejorMep == mepArs: shtData.range('A10:A13').value = ''
    else:
        shtData.range('A10').value = mejorMep
        shtData.range('A11').value = mepArs
        shtData.range('A12').value = dicc['arsCImep'][0]
        shtData.range('A13').value = namesArs(dicc['mepCI'][0],' - CI')

    if mejorMep24 == mepArs24: shtData.range('A14:A17').value = ''
    else:
        shtData.range('A14').value = mejorMep24
        shtData.range('A15').value = mepArs24
        shtData.range('A16').value = dicc['ars24mep'][0]
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
        shtData.range('A23').value = mepCcl24
        shtData.range('A24').value = dicc['ccl24'][0]
        shtData.range('A25').value = namesCcl(dicc['mep24'][0],' - 24hs')

    shtData.range('A1').value = 'symbol'

def ilRulo():
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

vuelta = 0
def traerADR():
    galiciaADR= yf.download('GGAL',period='1d',interval='1d')['Close'].values
    return galiciaADR[0]

def ruloAutomatico(celda):
    shtData.range('Q'+str(int(celda+1))).value = ""
    if celda+1 == 2 or celda+1 == 6 or celda+1 == 8 or celda+1 == 14 or celda+1 == 18 or celda+1 == 22:
        stockVenta = shtData.range('U'+str(int(celda+1))).value
        if stockVenta != None:
            shtData.range('S'+str(int(celda+1))).value = "-"
            shtData.range('R'+str(int(celda+2))).value = "+"
            shtData.range('S'+str(int(celda+3))).value = "-"
            shtData.range('R'+str(int(celda+4))).value = "+"
            
        else: print('No hay stock disponible para inciar RULO')

def cantidadAuto(nroCelda):
    cantidad = shtData.range('Y'+str(int(nroCelda))).value
    if not cantidad or cantidad == None or cantidad == 'None': 
        cantidad = 1
    return abs(int(cantidad))

def stokDisponible(nroCelda):
    stok = shtData.range('U'+str(int(nroCelda))).value
    if not stok or stok == None or stok == 'None': 
        stok = 0
    return abs(int(stok))

def buscoOperaciones(inicio,fin):
    for valor in shtData.range('P'+str(inicio)+':'+'U'+str(fin)).value:
        try:
            if not shtData.range('W1').value: # Permite TRAILING  ///////////////////////////////////////////////////////////////
                if not valor[5]:  pass
                else: 
                    if valor[5] < 0: vendido = 'si'
                    else: vendido = 'no'
                    if valor[0] < 25: pass
                    else: trailingStop('A'+str((int(valor[0]+1))),cantidadAuto(valor[0]+1),int(valor[0]),vendido)
        except: pass


        if valor[1]: # # Columna Q en el excel /////////////////////////////////////////////////////////////////////////////////
            if str(valor[1]).lower() == 'r': ruloAutomatico(valor[0])
            elif valor[1] == '+': 
                enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
            else: 
                try: 
                    enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0]) # Compra Bid
                except: shtData.range('Q'+str(int(valor[0]+1))).value = ''
            shtData.range('Q'+str(int(valor[0]+1))).value = ''
        
        if valor[2]: #  Columna R en el excel //////////////////////////////////////////////////////////////////////////////////
            if valor[2] == '+': 
                enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
            else: 
                try: 
                    enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0]) # Compra Ask
                except: shtData.range('R'+str(int(valor[0]+1))).value = ''
            shtData.range('R'+str(int(valor[0]+1))).value = ''
        
        if valor[3]: # Columna S en el excel ///////////////////////////////////////////////////////////////////////////////////
            if valor[3] == '-': 
                enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
            else: 
                try: 
                    enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0]) # Vendo Bid
                except: shtData.range('S'+str(int(valor[0]+1))).value = ''
            shtData.range('S'+str(int(valor[0]+1))).value = ''

        if valor[4]: # Columna T en el excel //////////////////////////////////////////////////////////////////////////////////
            if valor[4] == '-': 
                enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
            else: 
                try: 
                    enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0]) # Vendo Ask
                except: shtData.range('T'+str(int(valor[0]+1))).value = ''
            shtData.range('T'+str(int(valor[0]+1))).value = ''

def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    nombre = str(shtData.range(str(symbol)).value).split()
    if len(nombre) == 2: symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[1]) # Es caucho
    elif len(nombre) > 2: symbol = "MERV - XMEV - " + str(nombre[0]) + ' - ' + str(nombre[2])   # Son bonos
    else : symbol = "MERV - XMEV - " + str(nombre[0]) + ' - 24hs' # Son opciones
    precio = shtData.range(str(price)).value
    gastos = float(shtData.range('AB1').value/100)

    if tipo.lower() == 'buy': 
        try: 
            if esFinde == False:
                pyRofex.send_order(ticker=symbol, side=pyRofex.Side.BUY, size=abs(int(size)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
                print(f'______/ BUY   + {int(size)} {symbol} // precio: {precio}') 
        except: 
            shtData.range('Q'+str(int(celda+1))+':'+'R'+str(int(celda+1))).value = ''
            print(f'______/ ERROR en COMPRA. {symbol} // precio: {precio} // + {int(size)}')
        try: shtData.range('V'+str(int(celda+1))).value += abs(int(size))
        except: shtData.range('V'+str(int(celda+1))).value = abs(int(size))
    
    else: # VENTA
        try:
            if esFinde == False: 
                pyRofex.send_order(ticker=symbol, side=pyRofex.Side.SELL, size=abs(int(size)), price=float(precio),order_type=pyRofex.OrderType.LIMIT)
                print(f'______/ SELL  - {int(size)} {symbol} // precio: {precio}')
                gastos /= -1
        except:
            shtData.range('S'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            print(f'______/ ERROR en VENTA. {symbol} // precio: {precio} // {int(size)}')
        try: shtData.range('V'+str(int(celda+1))).value -= abs(int(size))
        except: shtData.range('V'+str(int(celda+1))).value = int(size) / -1

    if str(nombre[0]).upper() == 'GGAL' or str(nombre[0]).upper() == 'GGALD' or len(nombre) < 2 :
        shtData.range('X'+str(int(celda+1))).value = precio * (1 + gastos)
    else: shtData.range('X'+str(int(celda+1))).value = (precio / 100) * (1 + gastos)

def trailingStop(nombre=str,cantidad=int,nroCelda=int,vendido=str):
    try:
        costo = shtData.range('X'+str(int(nroCelda+1))).value 
        if not costo or costo == None or costo == 'None': soloContinua()
        nombre = str(shtData.range(str(nombre)).value).split()
        if str(nombre[0]).upper() == 'GGAL' or str(nombre[0]).upper() == 'GGALD' or len(nombre) < 2: 
            bid = shtData.range('C'+str(int(nroCelda+1))).value
            ask = shtData.range('D'+str(int(nroCelda+1))).value
            last = shtData.range('F'+str(int(nroCelda+1))).value
        else : 
            bid = shtData.range('C'+str(int(nroCelda+1))).value / 100
            ask = shtData.range('D'+str(int(nroCelda+1))).value / 100
            last = shtData.range('F'+str(int(nroCelda+1))).value / 100
        if not last or last == None or last == 'None': soloContinua()

        ganancia = shtData.range('Z1').value
        if not ganancia: ganancia = 0.001

        if len(nombre) < 2: # Ingresa si son OPCIONES ///////////////////////////////////////////////////////////////////////////
            if vendido == 'no': # OPCIONES COMPRADAS +++++++++++++++++
                if bid > abs(costo) * (1 + (ganancia*15)):
                    shtData.range('X'+str(int(nroCelda+1))).value = bid
                    print(f'BUYTRAIL {time.strftime("%H:%M:%S")} {nombre[0]} actual {last} siguiente precio {bid * (1+(ganancia*15))}')
                    if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'BUYTRAIL': pass
                    else: shtData.range('W'+str(int(nroCelda+1))).value = 'BUYTRAIL'
                    
                if not shtData.range('X1').value:
                    if last < abs(costo) * (1 - (ganancia*45)): 
                        if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'STOP':
                            if bid > last * (1-(ganancia*45)):
                                if shtData.range('Y'+str(int(nroCelda+1))).value : 
                                    try: shtData.range('U'+str(int(nroCelda+1))).value -= abs(cantidad)
                                    except: pass
                                    enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),abs(cantidad),nroCelda)
                        else:
                            if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                            else: 
                                print(f'STOP {time.strftime("%H:%M:%S")} {nombre[0]} {last} target salida {costo * (1-(ganancia*45))}')
                                shtData.range('W'+str(int(nroCelda+1))).value = 'STOP'

            else: # OPCIONES VENDIDAS -------------------------------
                if ask < abs(costo) * (1 - (ganancia*15)):
                    shtData.range('X'+str(int(nroCelda+1))).value = ask
                    print(f'SELLTRAIL {time.strftime("%H:%M:%S")} {nombre[0]} actual {last} siguiente precio {costo * (1-(ganancia*15))}')
                    if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'SELLTRAIL': pass
                    else: shtData.range('W'+str(int(nroCelda+1))).value = 'SELLTRAIL'
                    
                if not shtData.range('X1').value:  
                    if last > abs(costo) * (1 + (ganancia*45)): 
                        if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'STOP': 
                            if ask < last * (1-(ganancia*45)):
                                if shtData.range('Y'+str(int(nroCelda+1))).value : 
                                    try: shtData.range('U'+str(int(nroCelda+1))).value += abs(cantidad)
                                    except: pass
                                    enviarOrden('buy','A'+str((int(nroCelda)+1)),'D'+str((int(nroCelda)+1)),abs(cantidad),nroCelda)
                        else:
                            if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                            else: 
                                print(f'STOP {time.strftime("%H:%M:%S")} {nombre[0]} {last} target salida {costo * (1+(ganancia*45))}')
                                shtData.range('W'+str(int(nroCelda+1))).value = 'STOP'

        else: # Ingresa si son BONOS / LETRAS / ON / CEDEARS ////////////////////////////////////////////////////////////////////
            if time.strftime("%H:%M:%S") > '16:24:00' and str(nombre[2]).lower() == 'CI': 
                if time.strftime("%H:%M:%S") > '17:01:00': soloContinua()
                else: 
                    shtData.range('W'+str(int(nroCelda+1))).value = "CLOSED"
                    soloContinua()
            if time.strftime("%H:%M:%S") > '16:56:00' and str(nombre[2]).lower() == '24hs': 
                if time.strftime("%H:%M:%S") > '17:01:00': soloContinua()
                else: 
                    shtData.range('W'+str(int(nroCelda+1))).value = "CLOSED"
                    soloContinua()
            else:
                if bid > abs(costo) * (1 + ganancia):     
                    shtData.range('X'+str(int(nroCelda+1))).value = bid   
                    print(f'TRAILING {time.strftime("%H:%M:%S")} {nombre[0]} {last} precio objetivo {bid * (1+(ganancia))}')     
                    if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'BUYTRAIL': pass
                    else: shtData.range('W'+str(int(nroCelda+1))).value = 'BUYTRAIL'
                
                if not shtData.range('X1').value:
                    if last < abs(costo) * (1 - ganancia*3):
                        if str(shtData.range('W'+str(int(nroCelda+1))).value)=='STOP' and (bid)>(last)*(1-ganancia*1.5):
                            if shtData.range('Y'+str(int(nroCelda+1))).value : 
                                tengoStok = stokDisponible(nroCelda+1)
                                if tengoStok < 1: soloContinua()
                                elif cantidad > tengoStok: cantidad = tengoStok
                                shtData.range('U'+str(int(nroCelda+1))).value -= abs(cantidad)
                                shtData.range('W'+str(int(nroCelda+1))).value = ''
                                enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),abs(cantidad),nroCelda)
                        else: 
                            print(f'STOP {time.strftime("%H:%M:%S")} {nombre[0]} {last} target salida {costo * (1-ganancia*5)}')
                            if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'STOP': soloContinua()
                            else: 
                                shtData.range('W'+str(int(nroCelda+1))).value = 'STOP'
    except: soloContinua()

while True:

    if str(shtData.range('A1').value) != 'symbol': ilRulo()
    buscoOperaciones(rangoDesde,rangoHasta)

    if time.strftime("%H:%M:%S") > '17:01:00': 
        if time.strftime("%H:%M:%S") < '17:01:10':
            print(time.strftime("%H:%M:%S"), 'Mercado local cerrado, continua ADR. ')
            shtData.range('Q1').value = 'PRECIOS'
        
    try: 
        if not shtData.range('Q1').value:
            shtData.range('A30').options(index=False, headers=False).value = df_datos
    except: print('Hubo un error al actualizar excel')
    
    if not shtData.range('S1').value:
        try:
            if vuelta > 10: 
                valorAdr = traerADR()
                shtData.range('Z61').value = valorAdr
                shtData.range('Y62').value = time.strftime("%H:%M:%S")
                vuelta = 0
                if time.strftime("%H:%M:%S") > '17:30:20':
                    print(time.strftime("%H:%M:%S"), 'ADR cerrado. ')
                    pyRofex.close_websocket_connection()
                    break
            else: vuelta += 1
        except: print('ERROR, al cargar el ADR desde yahoo finance')
    #shtOperaciones.range('AI63').options(index=False, headers=False).value = operaciones
    time.sleep(3)


shtData.range('S1').value = 'ADR'
shtData.range('W1').value = 'TRAILING'
shtData.range('X1').value = 'STOP'
shtData.range('Y1').value = 'BROKER'
shtData.range('Z1').value = 0.001
        
        
    


