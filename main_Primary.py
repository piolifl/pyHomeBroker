import pyRofex
import time , math
import pandas as pd
import xlwings as xw
import winsound

wb = xw.Book('.\\epgb.xlsb')
shtTickers = wb.sheets('pyRofex')
shtData = wb.sheets('HomeBroker')

shtData.range('Q1').value = 'PRECIOS'
shtData.range('S1').value = 'OPERAR'
shtData.range('W1').value = 'TRAILING'
shtData.range('X1').value = 'STOP'
shtData.range('Z1').value = 0.001
shtData.range('AB1').value = 0.0025

rangoDesde = '2'
rangoHasta = '90'

pyRofex._set_environment_parameter("url", "https://api.eco.xoms.com.ar/", pyRofex.Environment.LIVE)
pyRofex._set_environment_parameter("ws", "wss://api.eco.xoms.com.ar/", pyRofex.Environment.LIVE)
pyRofex.initialize(user="20263866623", password="M8Khq6gQ_", account="62226", environment=pyRofex.Environment.LIVE)

rng = shtTickers.range('A2:C65').expand() # OPCIONES
tickers = pd.DataFrame(rng.value, columns=['ticker', 'symbol', 'strike'])
largoOpciones = len(tickers)
rng = shtTickers.range('E2:F5').expand() # ACCIONES
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
allLength = len(instruments)

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


#df_order = pd.DataFrame()

def order_error_handler(message):
    print("Error Message Received: {0}".format(message))

def order_exception_handler(e):
    print("Exception Occurred: {0}".format(e.message))

pyRofex.init_websocket_connection(market_data_handler=market_data_handler,
                                  error_handler=error_handler,
                                  exception_handler=exception_handler)

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

def namesArs(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA37D'+plazo
    elif nombre[:2] == 'BP': return 'BPOA7'+plazo
    elif nombre[:2] == 'KO': return 'KO'+plazo
    elif nombre[:2] == 'GOGL': return 'GOOGL'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') and (nombre[3:4] == 'D' or nombre[3:4] == 'C'):
        if (nombre[1:2] == 'F' or nombre[1:2] == 'Y'): return nombre[:1]+'20'+nombre[1:3]+plazo
        if (nombre[1:2] == 'J'): return nombre[:1]+'14'+nombre[1:3]+plazo
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
    celda,pesos,dolar = 93,1000,0.01
    tikers = {'cclCI':['',dolar],'ccl24':['',dolar],'mepCI':['',dolar],'mep24':['',dolar],'arsCIccl':['',pesos],'ars24ccl':['',pesos],'arsCImep':['',pesos],'ars24mep':['',pesos]}
    
    for valor in shtData.range('A'+str(celda)+':A'+str(allLength)).value:
        if not valor: continue
        name = str(valor).split()
        if len(name) < 2: continue
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


###############################################################  ENVIAR ORDENES ################################################    
def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    symbol = shtData.range(str(symbol)).value
    precio = shtData.range(str(price)).value
    if not shtData.range('V'+str(int(celda+1))).value: shtData.range('W'+str(int(celda+1))+':'+'X'+str(int(celda+1))).value = 0
    if tipo.lower() == 'buy': 
        try: 
            pyRofex.send_order_via_websocket(ticker=symbol,side=pyRofex.Side.BUY,size=int(size),order_type=pyRofex.OrderType.LIMIT,price=float(precio)) 
            #pyRofex.send_order(ticker=symbol,side=pyRofex.Side.BUY,size=int(size),order_type=pyRofex.OrderType.LIMIT,price=float(precio)) 
            
            if len(str(symbol).split())<2:
                shtData.range('AD'+str(int(celda+1))).value = float(precio*100)
                try: shtData.range('W'+str(int(celda+1))).value += int(size) * precio*100
                except: shtData.range('W'+str(int(celda+1))).value = int(size) * precio*100
                print(f'______/ BUY  {symbol} // precio: {precio} // + {int(size)}')
            else:
                shtData.range('AD'+str(int(celda+1))).value = float(precio/100)
                try: shtData.range('W'+str(int(celda+1))).value += int(size) * precio/100
                except: shtData.range('W'+str(int(celda+1))).value = int(size) * precio/100
                print(f'______/ BUY  {symbol} // precio: {precio} // + {int(size)}')

            shtData.range('Q'+str(int(celda+1))+':'+'R'+str(int(celda+1))).value = ''
            try: shtData.range('V'+str(int(celda+1))).value += int(size)
            except: shtData.range('V'+str(int(celda+1))).value = int(size)
            shtData.range('AC'+str(int(celda+1))).value = int(size)
        except: 
            shtData.range('Q'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            print(f'______/ ERROR en COMPRA. {symbol} // precio: {precio} // + {int(size)}')

    else: 
        try:
            pyRofex.send_order_via_websocket(ticker=symbol,side=pyRofex.Side.SELL,size=int(size),order_type=pyRofex.OrderType.LIMIT,price=float(precio)) 
            #pyRofex.send_order(ticker=symbol,side=pyRofex.Side.SELL,size=int(size),order_type=pyRofex.OrderType.LIMIT,price=float(precio)) 
            
            if len(str(symbol).split())<2:
                shtData.range('AF'+str(int(celda+1))).value = float(precio*100)
                try: shtData.range('W'+str(int(celda+1))).value -= int(size) * precio*100
                except: shtData.range('W'+str(int(celda+1))).value = int(size) * precio*100
                print(f'______/ SELL {symbol} // precio: {precio} // - {int(size)}')
            else:
                shtData.range('AF'+str(int(celda+1))).value = float(precio/100)
                try: shtData.range('W'+str(int(celda+1))).value -= int(size) * precio/100
                except: shtData.range('W'+str(int(celda+1))).value = int(size) * precio/100
                print(f'______/ SELL {symbol} // precio: {precio} // - {int(size)}')

            shtData.range('S'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            try: shtData.range('V'+str(int(celda+1))).value -= int(size)
            except: shtData.range('V'+str(int(celda+1))).value = int(size)/-1
            shtData.range('AE'+str(int(celda+1))).value = int(size)
        except:
            shtData.range('Q'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            print(f'______/ ERROR en VENTA. {symbol} // precio: {precio} // {int(size)/-1}')


    try: 
        tieneW = shtData.range('W'+str(int(celda+1))).value
        tieneV = shtData.range('V'+str(int(celda+1))).value
        if tieneW != 'TRAILING' or tieneW != 'STOP' or tieneW != '': 
            shtData.range('X'+str(int(celda+1))).value = tieneW / tieneV
        else: 
            shtData.range('W'+str(int(celda+1))).value = ''
            shtData.range('X'+str(int(celda+1))).value = 0
    except: 
        shtData.range('W'+str(int(celda+1))).value = ''
        shtData.range('X'+str(int(celda+1))).value = 0
############################################################### TRAILING STOP #################################################
def trailingStop(nombre=str,cantidad=int,nroCelda=int):
    try:
        nombre = str(shtData.range(str(nombre)).value).split()
        bid = shtData.range('C'+str(int(nroCelda+1))).value
        stock = shtData.range('V'+str(int(nroCelda+1))).value
        last = shtData.range('F'+str(int(nroCelda+1))).value
        costo = shtData.range('X'+str(int(nroCelda+1))).value 
        ganancia = shtData.range('Z1').value
        if not ganancia: ganancia = 0.0005
        if cantidad > stock : cantidad = int(stock)

        if len(nombre) < 2: # Ingresa si son OPCIONES ///////////////////////////////////////////////////////////////////////////
            if bid * 100 > costo * (1 + (ganancia*10)):
                if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'TRAILING': pass
                else: shtData.range('W'+str(int(nroCelda+1))).value = 'TRAILING'
                shtData.range('X'+str(int(nroCelda+1))).value = bid * 100
            if not shtData.range('X1').value:
                if last * 100 < costo * (1 - (ganancia*75)): # Precio baja activo stop y envia orden venta
                    if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'STOP' and bid > last * (1-(ganancia*15)):
                        shtData.range('W'+str(int(nroCelda+1))).value = ''
                        shtData.range('X'+str(int(nroCelda+1))).value = 0
                        if shtData.range('Y'+str(int(nroCelda+1))).value : 
                            enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),cantidad,nroCelda)
                    else:
                        if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                        else:
                            shtData.range('W'+str(int(nroCelda+1))).value = 'STOP'
                            winsound.PlaySound("SystemHand", winsound.SND_ALIAS)      

        else: # Ingresa si son BONOS / LETRAS / ON / CEDEARS ////////////////////////////////////////////////////////////////////
            if time.strftime("%H:%M:%S") > '16:24:50' and str(nombre[2]).lower() == 'ci': 
                shtData.range('W'+str(int(nroCelda+1))).value = "CLOSED"
                pass
            if time.strftime("%H:%M:%S") > '16:56:50' and str(nombre[2]).lower() == '24hs': 
                shtData.range('W'+str(int(nroCelda+1))).value = "CLOSED"
                pass
            else:
                # Rutina, si el precio BID sube modifica precio promedio de compra //////////////////////////////////////////////
                if bid / 100 > costo * (1 + ganancia):             
                    if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'TRAILING': pass
                    else: shtData.range('W'+str(int(nroCelda+1))).value = 'TRAILING'
                    shtData.range('X'+str(int(nroCelda+1))).value = round(bid / 100,5)

                # Si X1 esta vacio, habilita estrategias de ventas  ////////////////////////////////////////////////////////////
                if not shtData.range('X1').value:
                    #  Precio LAST baja, inica estrategia salida vendiendo stock ci en 24hs
                    if last / 100 < costo * (1 - ganancia):
                        if str(nombre[2]).lower() == 'ci':
                            bid2 = shtData.range('C'+str(int(nroCelda+2))).value
                            last2 = shtData.range('F'+str(int(nroCelda+2))).value
                            if str(shtData.range('W'+str(int(nroCelda+1))).value)=='STOP' and (bid2/100)>(last2/100)*(1-ganancia):
                                shtData.range('W'+str(int(nroCelda+1))).value = ''
                                shtData.range('X'+str(int(nroCelda+1))).value = 0
                                try: shtData.range('V'+str(int(nroCelda+1))).value -= cantidad
                                except: shtData.range('V'+str(int(nroCelda+1))).value = cantidad/-1
                                print(f'{time.strftime("%H:%M:%S")} STOP vendo    ',end=' || ')
                                if shtData.range('Y'+str(int(nroCelda+1))).value : 
                                    enviarOrden('sell','A'+str((int(nroCelda)+2)),'C'+str((int(nroCelda)+2)),cantidad,nroCelda+1)
                            else: 
                                if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                                else:
                                    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                                    shtData.range('W'+str(int(nroCelda+1))).value = 'STOP'
                        else:
                            if str(shtData.range('W'+str(int(nroCelda+1))).value)=='STOP' and (bid/100)>(last/100)*(1-ganancia):
                                shtData.range('W'+str(int(nroCelda+1))).value = ''
                                shtData.range('X'+str(int(nroCelda+1))).value = 0
                                print(f'{time.strftime("%H:%M:%S")} STOP vendo    ',end=' || ')
                                if shtData.range('Y'+str(int(nroCelda+1))).value : 
                                    enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),cantidad,nroCelda)
                            else: 
                                if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                                else:
                                    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                                    shtData.range('W'+str(int(nroCelda+1))).value = 'STOP'
    except: pass
############################################################## BUSCA OPERACIONES ###############################################
def buscoOperaciones(inicio,fin):
    for valor in shtData.range('P'+str(inicio)+':'+'V'+str(fin)).value:
        cantidad = shtData.range('Y'+str(int(valor[0]+1))).value
        if cantidad == None: cantidad = 1

        if not shtData.range('W1').value: # Activa TRAILING  ///////////////////////////////////////////////////////////////////
            if not valor[6]: pass
            else:
                try: 
                    if valor[6] > 0: trailingStop('A'+str((int(valor[0])+1)),cantidad,int(valor[0]))
                except: pass

        if valor[1]: # # Columna Q en el excel /////////////////////////////////////////////////////////////////////////////////
            if valor[1] == '+': enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
            else: 
                try: enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0]) # Compra Bid
                except: shtData.range('Q'+str(int(valor[0]+1))).value = ''
            shtData.range('Q'+str(int(valor[0]+1))).value = ''

        if valor[2]: #  Columna R en el excel //////////////////////////////////////////////////////////////////////////////////
            if valor[2] == '+': enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidad,valor[0])
            else: 
                try: enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0]) # Compra Ask
                except: shtData.range('R'+str(int(valor[0]+1))).value = ''
            shtData.range('R'+str(int(valor[0]+1))).value = ''

        if valor[3]: # Columna S en el excel ///////////////////////////////////////////////////////////////////////////////////
            if valor[3] == '-': enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
            else: 
                try: enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0]) # Vendo Bid
                except: shtData.range('S'+str(int(valor[0]+1))).value = ''
            shtData.range('S'+str(int(valor[0]+1))).value = ''

        if valor[4]: # Columna T en el excel //////////////////////////////////////////////////////////////////////////////////
            if valor[4] == '-': enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidad,valor[0])
            else: 
                try: enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0]) # Vendo Ask
                except: shtData.range('T'+str(int(valor[0]+1))).value = ''
            shtData.range('T'+str(int(valor[0]+1))).value = ''
################################################################################################################################

print(("online"), time.strftime("%H:%M:%S"))

while True:
    
    if str(shtData.range('A1').value) != 'symbol': ilRulo()
    if not shtData.range('S1').value : buscoOperaciones(rangoDesde,rangoHasta)
    time.sleep(10)
    try: 
        if not shtData.range('Q1').value:
            shtData.range('A30').options(index=False, headers=False).value = df_datos
            
    except: print("______ ERROR al cargar lista precios en Excel ______ ",time.strftime("%H:%M:%S")) 

    '''if time.strftime("%H:%M:%S") > '17:10:00':
        pyRofex.close_websocket_connection()
        print('Salida por cierre del mercado')
        break'''

