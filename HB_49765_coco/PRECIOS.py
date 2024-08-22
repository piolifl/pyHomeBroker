from pyhomebroker import HomeBroker     
import xlwings as xw                    
import pandas as pd                     
from datetime import date, timedelta
import time
import winsound
import os
import environ
import requests

env = environ.Env()
environ.Env.read_env()
wb = xw.Book('..\\epgb_pyHB.xlsb')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')
shtTest.range('Q1').value = 'BONOS'
shtTest.range('S1').value = 'OPCIONES'
shtTest.range('W1').value = 'TRAILING'
shtTest.range('X1').value = 'STOP'
shtTest.range('Z1').value = 0.001
shtTest.range('AB1').value = 0.0046
rangoDesde = '26'
rangoHasta = '89'

def getOptionsList():
    global allOptions
    rng = shtTickers.range('A2:A61').expand()
    oOpciones = rng.value
    allOptions = pd.DataFrame({'symbol': oOpciones},columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last","change", "open", "high", "low", "previous_close", "turnover", "volume",'operations', 'datetime'])
    allOptions = allOptions.set_index('symbol')
    allOptions['datetime'] = pd.to_datetime(allOptions['datetime'])
    return allOptions
def getAccionesList():
    rng = shtTickers.range('C2:C10').expand()
    oAcciones = rng.value
    ACC = pd.DataFrame({'symbol' : oAcciones}, columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last","change", "open", "high", "low", "previous_close", "turnover", "volume",'operations', 'datetime'])
    ACC = ACC.set_index('symbol')
    ACC['datetime'] = pd.to_datetime(ACC['datetime'])
    return ACC
def getBonosList():
    rng = shtTickers.range('E2:E75').expand()
    oBonos = rng.value
    Bonos = pd.DataFrame({'symbol' : oBonos}, columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last", "change", "open", "high", "low", "previous_close", "turnover", "volume", 'operations', 'datetime'])
    Bonos = Bonos.set_index('symbol')
    Bonos['datetime'] = pd.to_datetime(Bonos['datetime'])
    return Bonos
def getLetrasList():
    rng = shtTickers.range('G2:G20').expand()
    oLetras = rng.value
    Letras = pd.DataFrame({'symbol' : oLetras}, columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last","change", "open", "high", "low", "previous_close", "turnover", "volume",'operations', 'datetime'])
    Letras = Letras.set_index('symbol')
    Letras['datetime'] = pd.to_datetime(Letras['datetime'])
    return Letras
def getONSList():
    rng = shtTickers.range('I2:I50').expand()
    oONS = rng.value
    ONS = pd.DataFrame({'symbol' : oONS}, columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last","change", "open", "high", "low", "previous_close", "turnover", "volume",'operations', 'datetime'])
    ONS = ONS.set_index('symbol')
    ONS['datetime'] = pd.to_datetime(ONS['datetime'])
    return ONS
def getCedearsList():
    rng = shtTickers.range('K2:K50').expand()
    oCedears = rng.value
    Cedears = pd.DataFrame({'symbol' : oCedears}, columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last","change", "open", "high", "low", "previous_close", "turnover", "volume",'operations', 'datetime'])
    Cedears = Cedears.set_index('symbol')
    Cedears['datetime'] = pd.to_datetime(Cedears['datetime'])
    return Cedears
def getPanelGeneralList():
    rng = shtTickers.range('M2:M10').expand()
    oPanelGeneral = rng.value
    PanelGeneral = pd.DataFrame({'symbol' : oPanelGeneral}, columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last","change", "open", "high", "low", "previous_close", "turnover", "volume",'operations', 'datetime'])
    PanelGeneral = PanelGeneral.set_index('symbol')
    PanelGeneral['datetime'] = pd.to_datetime(PanelGeneral['datetime'])
    return PanelGeneral


i = 1
fechas = []
while i < 11:
    fecha = date.today() + timedelta(days=i)
    fechas.extend([fecha])
    i += 1
cauciones = pd.DataFrame({'settlement':fechas}, columns=['settlement', 'bid_amount', 'bid_rate', 'ask_rate', 'ask_amount','last', 'turnover'])
cauciones['settlement'] = pd.to_datetime(cauciones['settlement'])
cauciones = cauciones.set_index('settlement')

options = getOptionsList()
options = options.rename(columns={"bid_size": "bidsize", "ask_size": "asksize"})
ACC = getAccionesList()
bonos = getBonosList()
letras = getLetrasList()
ONS = getONSList()
cedears = getCedearsList()
PanelGeneral = getPanelGeneralList()


everything = pd.concat([ACC, bonos, letras, ONS, cedears, PanelGeneral ])
listLength = len(options) +30
allLength = len(everything) + listLength

def on_options(online, quotes):
    global options
    thisData = quotes
    thisData = thisData.drop(['expiration', 'strike', 'kind'], axis=1)
    thisData['change'] = thisData["change"] / 100
    thisData['datetime'] = pd.to_datetime(thisData['datetime'])
    thisData = thisData.rename(columns={"bid_size": "bidsize", "ask_size": "asksize"})
    options.update(thisData)
    
def on_securities(online, quotes):
    global ACC
    thisData = quotes
    thisData = thisData.reset_index()
    thisData['symbol'] = thisData['symbol'] + ' - ' +  thisData['settlement']
    thisData = thisData.drop(["settlement"], axis=1)
    thisData = thisData.set_index("symbol")
    thisData['change'] = thisData["change"] / 100
    thisData['datetime'] = pd.to_datetime(thisData['datetime'])
    everything.update(thisData)
    
def on_repos(online, quotes):
    global cauciones
    thisData = quotes
    thisData = thisData.reset_index()
    thisData = thisData.set_index("symbol")
    thisData = thisData[['PESOS' in s for s in quotes.index]]
    thisData = thisData.reset_index()
    thisData['settlement'] = pd.to_datetime(thisData['settlement'])
    thisData = thisData.set_index("settlement")
    thisData['last'] = thisData["last"] / 100
    thisData['bid_rate'] = thisData["bid_rate"] / 100
    thisData['ask_rate'] = thisData["ask_rate"] / 100
    thisData = thisData.drop(['open', 'high', 'low', 'volume', 'operations', 'datetime'], axis=1)
    thisData = thisData[['last', 'turnover', 'bid_amount', 'bid_rate', 'ask_rate', 'ask_amount']]
    cauciones.update(thisData)
#--------------------------------------------------------------------------------------------------------------------------------
def getTodos():
    hb.online.connect()
    hb.online.subscribe_options()
    hb.online.subscribe_securities('bluechips', '24hs')   # Acciones del Panel lider - 24hs
    hb.online.subscribe_securities('bluechips', 'SPOT')    # Acciones del Panel lider - spot
    hb.online.subscribe_securities('government_bonds', '24hs') # Bonos - 24hs
    hb.online.subscribe_securities('government_bonds', 'SPOT')  # Bonos - spot
    hb.online.subscribe_securities('cedears', '24hs')      # CEDEARS - 24hs
    hb.online.subscribe_securities('cedears', 'SPOT')      # CEDEARS - spot
    #hb.online.subscribe_securities('general_board', '24hs') # Acciones del Panel general - 24hs
    #hb.online.subscribe_securities('general_board', 'SPOT') # Acciones del Panel general - spot
    hb.online.subscribe_securities('short_term_government_bonds', '24hs')  # LETRAS - 24hs
    hb.online.subscribe_securities('short_term_government_bonds', 'SPOT')   # LETRAS - spot
    hb.online.subscribe_securities('corporate_bonds', '24hs')  # Obligaciones Negociables - 24hs
    hb.online.subscribe_securities('corporate_bonds', 'SPOT')  # Obligaciones Negociables - spot
    hb.online.subscribe_repos()

def login():
    hb.auth.login(dni=str(os.environ.get('dni')), 
    user=str(os.environ.get('user')),  
    password=str(os.environ.get('password')),
    raise_exception=True)

hb = HomeBroker(int(os.environ.get('broker')), on_options=on_options, on_securities=on_securities, on_repos=on_repos)
login()
getTodos()

def getPortfolio(hb, comitente):
    try:
        shtTest.range('U'+str(rangoDesde)+':'+'U'+str(rangoHasta)).value = ''
        payload = {'comitente': str(comitente),
        'consolida': '0',
        'proceso': '22',
        'fechaDesde': None,
        'fechaHasta': None,
        'tipo': None,
        'especie': None,
        'comitenteMana': None}
        
        if os.environ.get('name') == 'COCOS.CAPITAL':
            portfolio = requests.post("https://cocoscap.com/Consultas/GetConsulta", cookies=hb.auth.cookies, json=payload).json()
        else: 
            portfolio = requests.post("https://clientes.bcch.org.ar/Consultas/GetConsulta", cookies=hb.auth.cookies, json=payload).json()

        for i in portfolio['Result']['Activos'][0]['Subtotal'][0]['APERTURA']:
            if i['IMPO'] != None: print(i['DETA'],i['IMPO'] )

        subtotal = [ i['Subtotal'] for i in portfolio["Result"]["Activos"][0:] ]
        for i in subtotal[0:]:
            if i[0]['NERE'] != 'Pesos':  
                subtotal = [ ( x['NERE'],x['CAN0'],x['CANT']) for x in i[0:] if x['CANT'] != None]
                for x in subtotal:
                    for valor in shtTest.range('A'+str(rangoDesde)+':'+'P'+str(rangoHasta)).value:
                        if not valor[0]: continue
                        ticker = str(valor[0]).split()
                        if x[0] == ticker[0]: 
                            shtTest.range('U'+str(int(valor[15]+1))).value = x[2]
                            if not shtTest.range('V'+str(int(valor[15]+1))).value:
                                if len(ticker) < 2: 
                                    shtTest.range('X'+str(int(valor[15]+1))).value = x[1]
                                else:
                                    try: shtTest.range('X'+str(int(valor[15]+1))).value = x[1] /100
                                    except: shtTest.range('X'+str(int(valor[15]+1))).value = x[1]
    except: pass

#--------------------------------------------------------------------------------------------------------------------------------
print(time.strftime("%H:%M:%S"),f"Logueo correcto en: {os.environ.get('name')} cuenta: {int(os.environ.get('account_id'))}")


#--------------------------------------------------------------------------------------------------------------------------------
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
    mepArs = namesMep(dicc['arsCImep'][0],' - spot')
    mepArs24 = namesMep(dicc['ars24mep'][0],' - 24hs')
    mepCcl = namesMep(dicc['cclCI'][0],' - spot')
    mepCcl24 = namesMep(dicc['ccl24'][0],' - 24hs')

    if mejorMep == 'AL30D - spot': shtTest.range('A2:A5').value = ''
    else: 
        shtTest.range('A2').value = mejorMep
        shtTest.range('A3').value = 'AL30D - spot'
        shtTest.range('A4').value = 'AL30 - spot'
        shtTest.range('A5').value = namesArs(dicc['mepCI'][0],' - spot')

    if mejorMep24 == 'AL30D - 24hs': shtTest.range('A6:A9').value = ''
    else: 
        shtTest.range('A6').value = mejorMep24
        shtTest.range('A7').value = 'AL30D - 24hs'
        shtTest.range('A8').value = 'AL30 - 24hs'
        shtTest.range('A9').value = namesArs(dicc['mep24'][0],' - 24hs')
    
    if mejorMep == mepArs: shtTest.range('A10:A13').value = ''
    else:
        shtTest.range('A10').value = mejorMep
        shtTest.range('A11').value = mepArs
        shtTest.range('A12').value = dicc['arsCImep'][0]
        shtTest.range('A13').value = namesArs(dicc['mepCI'][0],' - spot')

    if mejorMep24 == mepArs24: shtTest.range('A14:A17').value = ''
    else:
        shtTest.range('A14').value = mejorMep24
        shtTest.range('A15').value = mepArs24
        shtTest.range('A16').value = dicc['ars24mep'][0]
        shtTest.range('A17').value = namesArs(dicc['mep24'][0],' - 24hs')

    if mejorMep == mepCcl:  shtTest.range('A18:A21').value = ''
    else:
        shtTest.range('A18').value = mejorMep
        shtTest.range('A19').value = mepCcl
        shtTest.range('A20').value = dicc['cclCI'][0]
        shtTest.range('A21').value = namesCcl(dicc['mepCI'][0],' - spot')

    if mejorMep24 == mepCcl24: shtTest.range('A22:A25').value = ''
    else:
        shtTest.range('A22').value = mejorMep24
        shtTest.range('A23').value = mepCcl24
        shtTest.range('A24').value = dicc['ccl24'][0]
        shtTest.range('A25').value = namesCcl(dicc['mep24'][0],' - 24hs')

    shtTest.range('A1').value = 'symbol'

def ilRulo():
    celda,pesos,dolar = listLength+2,1000,0.01
    tikers = {'cclCI':['',dolar],'ccl24':['',dolar],'mepCI':['',dolar],'mep24':['',dolar],'arsCIccl':['',pesos],'ars24ccl':['',pesos],'arsCImep':['',pesos],'ars24mep':['',pesos]}
    
    for valor in shtTest.range('A'+str(celda)+':A'+str(allLength)).value:
        if not valor: continue
        name = str(valor).split()
        
        if str(name[2]).lower() == 'spot':

            if str(name[0][-1:]).upper()=='C':

                arsC = shtTest.range('AA'+str(celda)).value
                if not arsC: arsC = 1000
                if arsC > tikers['arsCIccl'][1]: tikers['arsCIccl'] = [namesArs(name[0],' - spot'),arsC]

                ccl = shtTest.range('Z'+str(celda)).value
                if not ccl: ccl = 0.01
                if ccl > tikers['cclCI'][1]: tikers['cclCI'] = [valor,ccl]

            if str(name[0][-1:]).upper()=='D': 

                arsM = shtTest.range('AA'+str(celda)).value
                if not arsM: arsM = 1000
                if arsM > tikers['arsCImep'][1]: tikers['arsCImep'] = [namesArs(name[0],' - spot'),arsM]

                mep = shtTest.range('Z'+str(celda)).value
                if not mep: mep = 0.01
                if mep > tikers['mepCI'][1]: tikers['mepCI'] = [valor,mep]


        if str(name[2]) == '24hs':
            if str(name[0][-1:]).upper()=='C':
                arsC = shtTest.range('AA'+str(celda)).value
                if not arsC: arsC = 1000
                if arsC > tikers['ars24ccl'][1]: tikers['ars24ccl'] = [namesArs(name[0],' - 24hs'),arsC]
                ccl = shtTest.range('Z'+str(celda)).value
                if not ccl: ccl = 0.01
                if ccl > tikers['ccl24'][1]: tikers['ccl24'] = [valor,ccl]

            if str(name[0][-1:]).upper()=='D': 
                arsM = shtTest.range('AA'+str(celda)).value
                if not arsM: arsM = 1000
                if arsM > tikers['ars24mep'][1]: tikers['ars24mep'] = [namesArs(name[0],' - 24hs'),arsM]
                mep = shtTest.range('Z'+str(celda)).value
                if not mep: mep = 0.01
                if mep > tikers['mep24'][1]: tikers['mep24'] = [valor,mep]
        celda +=1
    cargoXplazo(tikers)

def cancelaCompra(celda):
    try:
        orderC = shtTest.range('AB'+str(int(celda+1))).value
        if not orderC: orderC = 0
        hb.orders.cancel_order(int(os.environ.get('account_id')),int(orderC))
        shtTest.range('V'+str(int(celda+1))).value -= shtTest.range('AC'+str(int(celda+1))).value
        shtTest.range('AB'+str(int(celda+1))+':'+'AD'+str(int(celda+1))).value = ''
        print(f" /// Cancela Compra nro: {int(orderC)} ",time.strftime("%H:%M:%S"))
    except: print(time.strftime("%H:%M:%S"),'______ ERROR al cancelar orden.')

def cancelarVenta(celda):
    try:
        orderV = shtTest.range('AE'+str(int(celda+1))).value
        if not orderV: orderV = 0
        hb.orders.cancel_order(int(os.environ.get('account_id')),int(orderV))
        shtTest.range('V'+str(int(celda+1))).value += shtTest.range('AF'+str(int(celda+1))).value
        shtTest.range('AE'+str(int(celda+1))+':'+'AG'+str(int(celda+1))).value = ''
        print(f" /// Cancela Venta nro : {int(orderV)} ",time.strftime("%H:%M:%S"))
    except: print(time.strftime("%H:%M:%S"),'______ ERROR al cancelar orden.')

def cancelarTodo(desde,hasta):
    try:
        hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
        shtTest.range('AB'+str(desde)+':'+'AH'+str(hasta)).value = ''
        print(" /// Todas las ordenes activas canceladas ",time.strftime("%H:%M:%S"))
    except: print(time.strftime("%H:%M:%S"),'______ ERROR al cancelar orden.')

###############################################################  ENVIAR ORDENES ################################################    
def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    global orderC, orderV
    orderC, orderV = 0,0
    symbol = str(shtTest.range(str(symbol)).value).split()
    precio = shtTest.range(str(price)).value
    if not shtTest.range('V'+str(int(celda+1))).value: shtTest.range('W'+str(int(celda+1))+':'+'X'+str(int(celda+1))).value = 0
    if tipo.lower() == 'buy': 
        try: 
            if len(symbol) < 2:
                orderC = hb.orders.send_buy_order(symbol[0],'24hs', float(precio), int(size))
                shtTest.range('AD'+str(int(celda+1))).value = float(precio)
                try: shtTest.range('W'+str(int(celda+1))).value += int(size) * precio*100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio*100
                print(f'______/ BUY  opcion {symbol[0]} 24hs // precio: {precio} // + {int(size)}') 
            else:
                orderC = hb.orders.send_buy_order(symbol[0],symbol[2], float(precio), int(size))
                shtTest.range('AD'+str(int(celda+1))).value = float(precio/100)
                try: shtTest.range('W'+str(int(celda+1))).value += int(size) * precio/100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio/100
                print(f'______/ BUY  {symbol[0]} {symbol[2]} // precio: {round(precio/100,4)} // + {int(size)}')
            # ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            shtTest.range('Q'+str(int(celda+1))+':'+'R'+str(int(celda+1))).value = ''
            try: shtTest.range('V'+str(int(celda+1))).value += int(size)
            except: shtTest.range('V'+str(int(celda+1))).value = int(size)
            shtTest.range('AB'+str(int(celda+1))).value = orderC
            shtTest.range('AC'+str(int(celda+1))).value = int(size)
        except: 
            shtTest.range('Q'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            print(f'______/ ERROR en COMPRA. {symbol[0]} // precio: {precio} // + {int(size)}')
    else: 
        try:
            if len(symbol) < 2:
                orderV = hb.orders.send_sell_order(symbol[0],'24hs', float(precio), int(size))
                shtTest.range('AG'+str(int(celda+1))).value = float(precio)
                try: shtTest.range('W'+str(int(celda+1))).value -= int(size) * precio*100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio*100
                print(f'______/ SELL opcion {symbol[0]} 24hs // precio: {precio} // - {int(size)}')
            else:
                orderV = hb.orders.send_sell_order(symbol[0],symbol[2], float(precio), int(size))
                shtTest.range('AG'+str(int(celda+1))).value = float(precio/100)
                try: shtTest.range('W'+str(int(celda+1))).value -= int(size) * precio/100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio/100
                print(f'______/ SELL {symbol[0]} {symbol[2]} // precio: {round(precio/100,4)} // - {int(size)}')
            # ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            shtTest.range('S'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            try: shtTest.range('V'+str(int(celda+1))).value -= int(size)
            except: shtTest.range('V'+str(int(celda+1))).value = int(size)/-1
            shtTest.range('AE'+str(int(celda+1))).value = orderV
            shtTest.range('AF'+str(int(celda+1))).value = int(size)
        except:
            shtTest.range('Q'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            print(f'______/ ERROR en VENTA. {symbol[0]} // precio: {precio} // {int(size)/-1}')
    try: 
        tieneW = shtTest.range('W'+str(int(celda+1))).value
        tieneV = shtTest.range('V'+str(int(celda+1))).value
        if tieneW != 'TRAILING' or tieneW != 'STOP' or tieneW != '': 
            shtTest.range('X'+str(int(celda+1))).value = tieneW / tieneV
        else: 
            shtTest.range('W'+str(int(celda+1))).value = ''
            shtTest.range('X'+str(int(celda+1))).value = 0
    except: 
        shtTest.range('W'+str(int(celda+1))).value = ''
        shtTest.range('X'+str(int(celda+1))).value = 0
############################################################### TRAILING STOP #################################################
def trailingStop(nombre=str,cantidad=int,nroCelda=int):
    try:
        nombre = str(shtTest.range(str(nombre)).value).split()
        bid = shtTest.range('C'+str(int(nroCelda+1))).value
        stock = shtTest.range('V'+str(int(nroCelda+1))).value
        last = shtTest.range('F'+str(int(nroCelda+1))).value
        costo = shtTest.range('X'+str(int(nroCelda+1))).value 
        ganancia = shtTest.range('Z1').value
        if not ganancia: ganancia = 0.0005
        if cantidad > stock : cantidad = int(stock)

        if len(nombre) < 2: # Ingresa si son OPCIONES ///////////////////////////////////////////////////////////////////////////
            if bid * 100 > costo * (1 + (ganancia*10)):
                if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'TRAILING': pass
                else: shtTest.range('W'+str(int(nroCelda+1))).value = 'TRAILING'
                shtTest.range('X'+str(int(nroCelda+1))).value = bid * 100
            if not shtTest.range('X1').value:
                if last * 100 < costo * (1 - (ganancia*75)): # Precio baja activo stop y envia orden venta
                    if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP' and bid > last * (1-(ganancia*15)):
                        shtTest.range('W'+str(int(nroCelda+1))).value = ''
                        shtTest.range('X'+str(int(nroCelda+1))).value = 0
                        if shtTest.range('Y'+str(int(nroCelda+1))).value : 
                            enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),cantidad,nroCelda)
                    else:
                        if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                        else:
                            shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP'
                            winsound.PlaySound("SystemHand", winsound.SND_ALIAS)      

        else: # Ingresa si son BONOS / LETRAS / ON / CEDEARS ////////////////////////////////////////////////////////////////////
            if time.strftime("%H:%M:%S") > '16:24:50' and str(nombre[2]).lower() == 'spot': 
                shtTest.range('W'+str(int(nroCelda+1))).value = "CLOSED"
                pass
            if time.strftime("%H:%M:%S") > '16:56:50' and str(nombre[2]).lower() == '24hs': 
                shtTest.range('W'+str(int(nroCelda+1))).value = "CLOSED"
                pass
            else:
                # Rutina, si el precio BID sube modifica precio promedio de compra //////////////////////////////////////////////
                if bid / 100 > costo * (1 + ganancia):             
                    if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'TRAILING': pass
                    else: shtTest.range('W'+str(int(nroCelda+1))).value = 'TRAILING'
                    shtTest.range('X'+str(int(nroCelda+1))).value = round(bid / 100,5)

                # Si X1 esta vacio, habilita estrategias de ventas  ////////////////////////////////////////////////////////////
                if not shtTest.range('X1').value:
                    #  Precio LAST baja, inica estrategia salida vendiendo stock spot en 24hs
                    if last / 100 < costo * (1 - ganancia):
                        if str(nombre[2]).lower() == 'spot':
                            bid2 = shtTest.range('C'+str(int(nroCelda+2))).value
                            last2 = shtTest.range('F'+str(int(nroCelda+2))).value
                            if str(shtTest.range('W'+str(int(nroCelda+1))).value)=='STOP' and (bid2/100)>(last2/100)*(1-ganancia):
                                shtTest.range('W'+str(int(nroCelda+1))).value = ''
                                shtTest.range('X'+str(int(nroCelda+1))).value = 0
                                try: shtTest.range('V'+str(int(nroCelda+1))).value -= cantidad
                                except: shtTest.range('V'+str(int(nroCelda+1))).value = cantidad/-1
                                print(f'{time.strftime("%H:%M:%S")} STOP vendo    ',end=' || ')
                                if shtTest.range('Y'+str(int(nroCelda+1))).value : 
                                    enviarOrden('sell','A'+str((int(nroCelda)+2)),'C'+str((int(nroCelda)+2)),cantidad,nroCelda+1)
                            else: 
                                if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                                else:
                                    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                                    shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP'
                        else:
                            if str(shtTest.range('W'+str(int(nroCelda+1))).value)=='STOP' and (bid/100)>(last/100)*(1-ganancia):
                                shtTest.range('W'+str(int(nroCelda+1))).value = ''
                                shtTest.range('X'+str(int(nroCelda+1))).value = 0
                                print(f'{time.strftime("%H:%M:%S")} STOP vendo    ',end=' || ')
                                if shtTest.range('Y'+str(int(nroCelda+1))).value : 
                                    enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),cantidad,nroCelda)
                            else: 
                                if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                                else:
                                    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                                    shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP'
    except: pass
############################################################## BUSCA OPERACIONES ###############################################
def buscoOperaciones(inicio,fin):
    for valor in shtTest.range('P'+str(inicio)+':'+'V'+str(fin)).value:

        cantidad = shtTest.range('Y'+str(int(valor[0]+1))).value
        if cantidad == None: cantidad = 1

        if not shtTest.range('W1').value: # Activa TRAILING  ///////////////////////////////////////////////////////////////////
            if not valor[6]: pass
            else:
                try: 
                    if valor[6] > 0: trailingStop('A'+str((int(valor[0])+1)),cantidad,int(valor[0]))
                except: pass

        if valor[1]: # # Columna Q en el excel /////////////////////////////////////////////////////////////////////////////////
            if str(valor[1]).lower() == 'c': cancelaCompra(valor[0])
            elif str(valor[1]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[1] == '+': enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
            elif str(valor[1]).upper() == 'P': getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0]) # Compra Bid
                except: shtTest.range('Q'+str(int(valor[0]+1))).value = ''
            shtTest.range('Q'+str(int(valor[0]+1))).value = ''

        if valor[2]: #  Columna R en el excel //////////////////////////////////////////////////////////////////////////////////
            if str(valor[2]).lower() == 'c': cancelaCompra(valor[0])
            elif str(valor[2]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[2] == '+': enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidad,valor[0])
            elif str(valor[2]).upper() == 'P': getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0]) # Compra Ask
                except: shtTest.range('R'+str(int(valor[0]+1))).value = ''
            shtTest.range('R'+str(int(valor[0]+1))).value = ''

        if valor[3]: # Columna S en el excel ///////////////////////////////////////////////////////////////////////////////////
            if str(valor[3]).lower() == 'v': cancelarVenta(valor[0])
            elif str(valor[3]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[3] == '-': enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
            elif str(valor[3]).upper() == 'P': getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0]) # Vendo Bid
                except: shtTest.range('S'+str(int(valor[0]+1))).value = ''
            shtTest.range('S'+str(int(valor[0]+1))).value = ''

        if valor[4]: # Columna T en el excel //////////////////////////////////////////////////////////////////////////////////
            if str(valor[4]).lower() == 'v': cancelarVenta(valor[0])
            elif str(valor[4]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[4] == '-': enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidad,valor[0])
            elif str(valor[4]).upper() == 'P': getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0]) # Vendo Ask
                except: shtTest.range('T'+str(int(valor[0]+1))).value = ''
            shtTest.range('T'+str(int(valor[0]+1))).value = ''
############################################################ CARGA BUCLE EN EXCEL ##############################################
while True:

    if time.strftime("%H:%M:%S") > '17:01:00': 
        if time.strftime("%H:%M:%S") > '17:10:00': pass
        else:
            try: getPortfolio(hb, os.environ.get('account_id'))
            except: pass
            break
    time.sleep(2)

    try: 
        if not shtTest.range('Q1').value:
            shtTest.range('A'+str(listLength)).options(index=True,header=False).value = everything
            try: shtTest.range('AJ2').options(index=True, header=False).value = cauciones
            except: print("______ ERROR al cargar cauciones en Excel ______ ",time.strftime("%H:%M:%S")) 
    except: print("______ ERROR al cargar Bonos/Letras en Excel ______ ",time.strftime("%H:%M:%S")) 

    try:
        if not shtTest.range('S1').value: 
            shtTest.range('A30').options(index=True,header=False).value=options  
    except: print("______ ERROR al cargar OPCIONES en Excel ______ ",time.strftime("%H:%M:%S")) 
        
    if str(shtTest.range('A1').value) != 'symbol': ilRulo()
    
try: 
    hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
    hb.online.disconnect()
except: pass
print(time.strftime("%H:%M:%S"), 'Mercado cerrado. ')
shtTest.range('Q1').value = 'BONOS'
shtTest.range('S1').value = 'OPCIONES'
shtTest.range('W1').value = 'TRAILING'
shtTest.range('X1').value = 'STOP'
shtTest.range('Y1').value = 'BROKER'

#[ ]><   \n
