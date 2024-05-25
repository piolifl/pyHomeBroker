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
wb = xw.Book('.\\epgb_pyHB.xlsx')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')
shtTest.range('Q1').value = 'BONOS'
shtTest.range('S1').value = 'OPCIONES'
shtTest.range('W1').value = 'TRAILING'
shtTest.range('X1').value = 'STOP'
shtTest.range('Y1').value = -10
shtTest.range('Z1').value = 0.0005
shtTest.range('AB1').value = 0.0001
rangoDesde = '26'
rangoHasta = '59'

def getBonosList():
    rng = shtTickers.range('E2:E145').expand()
    oBonos = rng.value
    Bonos = pd.DataFrame({'symbol' : oBonos}, columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last", "change", "open", "high", "low", "previous_close", "turnover", "volume", 'operations', 'datetime'])
    Bonos = Bonos.set_index('symbol')
    Bonos['datetime'] = pd.to_datetime(Bonos['datetime'])
    return Bonos
def getOptionsList():
    global allOptions
    rng = shtTickers.range('A2:A35').expand()
    oOpciones = rng.value
    allOptions = pd.DataFrame({'symbol': oOpciones},columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last","change", "open", "high", "low", "previous_close", "turnover", "volume",'operations', 'datetime'])
    allOptions = allOptions.set_index('symbol')
    allOptions['datetime'] = pd.to_datetime(allOptions['datetime'])
    return allOptions

i = 1
fechas = []
while i < 11:
    fecha = date.today() + timedelta(days=i)
    fechas.extend([fecha])
    i += 1
cauciones = pd.DataFrame({'settlement':fechas}, columns=['settlement', 'bid_amount', 'bid_rate', 'ask_rate', 'ask_amount','last', 'turnover'])
cauciones['settlement'] = pd.to_datetime(cauciones['settlement'])
cauciones = cauciones.set_index('settlement')

bonos = getBonosList()
options = getOptionsList()
options = options.rename(columns={"bid_size": "bidsize", "ask_size": "asksize"})
everything = bonos
listLength = len(options) +30

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
    #hb.online.subscribe_securities('bluechips', '24hs')    # Acciones del Panel lider - 24hs
    hb.online.subscribe_securities('bluechips', '24hs')   # Acciones del Panel lider - 24hs
    hb.online.subscribe_securities('bluechips', 'SPOT')    # Acciones del Panel lider - spot
    #hb.online.subscribe_securities('government_bonds', '24hs')  # Bonos - 24hs
    hb.online.subscribe_securities('government_bonds', '24hs') # Bonos - 24hs
    hb.online.subscribe_securities('government_bonds', 'SPOT')  # Bonos - spot
    #hb.online.subscribe_securities('cedears', '24hs')      # CEDEARS - 24hs
    # hb.online.subscribe_securities('cedears', '24hs')      # CEDEARS - 24hs
    #hb.online.subscribe_securities('cedears', 'SPOT')      # CEDEARS - spot
    # hb.online.subscribe_securities('general_board', '24hs') # Acciones del Panel general - 24hs
    # hb.online.subscribe_securities('general_board', '24hs') # Acciones del Panel general - 24hs
    # hb.online.subscribe_securities('general_board', 'SPOT') # Acciones del Panel general - spot
    #hb.online.subscribe_securities('short_term_government_bonds', '24hs')   # LETRAS - 24hs
    hb.online.subscribe_securities('short_term_government_bonds', '24hs')  # LETRAS - 24hs
    hb.online.subscribe_securities('short_term_government_bonds', 'SPOT')   # LETRAS - spot
    #hb.online.subscribe_securities('corporate_bonds', '24hs')  # Obligaciones Negociables - 24hs
    # hb.online.subscribe_securities('corporate_bonds', '24hs')  # Obligaciones Negociables - 24hs
    #hb.online.subscribe_securities('corporate_bonds', 'SPOT')  # Obligaciones Negociables - spot
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
    payload = {'comitente': str(comitente),
     'consolida': '0',
     'proceso': '22',
     'fechaDesde': None,
     'fechaHasta': None,
     'tipo': None,
     'especie': None,
     'comitenteMana': None}
    
    portfolio = requests.post("https://cocoscap.com/Consultas/GetConsulta", cookies=hb.auth.cookies, json=payload).json()
    print()
    subtotal = [ (i['DETA'],i['IMPO']) for i in portfolio["Result"]["Totales"]["Detalle"] ]
    print(subtotal)
    subtotal = [ i['Subtotal'] for i in portfolio["Result"]["Activos"][0:] ]
    for i in subtotal[0:]:
        if i[0]['NERE'] != 'Pesos':  
            subtotal = [ ( x['NERE'],x['CAN0'],x['CANT'],' || ',x['PCIO'],x['GTOS']) for x in i[0:] if x['CANT'] != None]
            for x in subtotal: print(x)

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
        if (nombre[1:2] == 'L'): return nombre[:1]+'26'+nombre[1:3]+plazo
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
    if mejorMep == 'AL30D - spot': shtTest.range('A2').value = ''
    else: shtTest.range('A2').value = mejorMep
    shtTest.range('A3').value = 'AL30D - spot'
    shtTest.range('A4').value = 'AL30 - spot'
    shtTest.range('A5').value = namesArs(dicc['mepCI'][0],' - spot')

    mejorMep = dicc['mep24'][0]
    if mejorMep == 'AL30D - 24hs': shtTest.range('A6').value = ''
    else: shtTest.range('A6').value = mejorMep
    shtTest.range('A7').value = 'AL30D - 24hs'
    shtTest.range('A8').value = 'AL30 - 24hs'
    shtTest.range('A9').value = namesArs(dicc['mep24'][0],' - 24hs')
    
    shtTest.range('A10').value = dicc['mepCI'][0]
    shtTest.range('A11').value = namesMep(dicc['arsCImep'][0],' - spot')
    shtTest.range('A12').value = dicc['arsCImep'][0]
    shtTest.range('A13').value = namesArs(dicc['mepCI'][0],' - spot')
    shtTest.range('A14').value = dicc['mep24'][0]
    shtTest.range('A15').value = namesMep(dicc['ars24mep'][0],' - 24hs')
    shtTest.range('A16').value = dicc['ars24mep'][0]
    shtTest.range('A17').value = namesArs(dicc['mep24'][0],' - 24hs')

    shtTest.range('A18').value = dicc['mepCI'][0]
    shtTest.range('A19').value = namesMep(dicc['cclCI'][0],' - spot')
    shtTest.range('A20').value = dicc['cclCI'][0]
    shtTest.range('A21').value = namesCcl(dicc['mepCI'][0],' - spot')
    shtTest.range('A22').value = dicc['mep24'][0]
    shtTest.range('A23').value = namesMep(dicc['ccl24'][0],' - 24hs')
    shtTest.range('A24').value = dicc['ccl24'][0]
    shtTest.range('A25').value = namesCcl(dicc['mep24'][0],' - 24hs')

    shtTest.range('A1').value = 'symbol'
    winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)

def ilRulo():
    celda,pesos,dolar = 64,1000,0.01
    tikers = {'cclCI':['',dolar],'ccl24':['',dolar],'mepCI':['',dolar],'mep24':['',dolar],'arsCIccl':['',pesos],'ars24ccl':['',pesos],'arsCImep':['',pesos],'ars24mep':['',pesos]}
    
    for valor in shtTest.range('A64:A201').value:
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

############################################ ENVIAR ORDENES ################################################    
def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    global orderC, orderV
    orderC, orderV = 0,0
    symbol = str(shtTest.range(str(symbol)).value).split()
    precio = shtTest.range(str(price)).value
    recompro = float(shtTest.range('Y1').value)
    if not shtTest.range('V'+str(int(celda+1))).value: shtTest.range('V'+str(int(celda+1))+':'+'X'+str(int(celda+1))).value = 0
    if tipo.lower() == 'buy': 
        try: 
            if len(symbol) < 2:
                if str(shtTest.range('X1').value) == 'REC': 
                    if not recompro: shtTest.range('Y1').value = -1
                    else:  precio += recompro / 10
                    shtTest.range('X1').value = ''
                    print(f'{time.strftime("%H:%M:%S")} RECOMPRA ',end=' || ')
                orderC = hb.orders.send_buy_order(symbol[0],'24hs', float(precio), int(size))
                shtTest.range('AE'+str(int(celda+1))).value = float(precio)
                try: shtTest.range('V'+str(int(celda+1))).value += int(size)
                except: shtTest.range('V'+str(int(celda+1))).value = int(size)
                try: shtTest.range('W'+str(int(celda+1))).value += int(size) * precio*100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio*100
                print(f'BUY  {symbol[0]} 24hs // precio: {precio} // + {int(size)} // orden: {orderC}')
            else:
                if str(shtTest.range('X1').value) == 'REC': 
                    if not recompro: shtTest.range('Y1').value = -1
                    else:  precio += recompro * 10
                    shtTest.range('X1').value = ''
                    print(f'{time.strftime("%H:%M:%S")} RECOMPRA ',end=' || ')
                orderC = hb.orders.send_buy_order(symbol[0],symbol[2], float(precio), int(size))
                shtTest.range('AE'+str(int(celda+1))).value = float(precio/100)
                try: shtTest.range('V'+str(int(celda+1))).value += int(size)
                except: shtTest.range('V'+str(int(celda+1))).value = int(size)
                try: shtTest.range('W'+str(int(celda+1))).value += int(size) * precio/100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio/100
                print(f'BUY  {symbol[0]} {symbol[2]} // precio {round(precio/100,4)} // + {int(size)} // orden: {orderC}')
            shtTest.range('AB'+str(int(celda+1))).value = orderC
            shtTest.range('AD'+str(int(celda+1))).value = int(size)
            shtTest.range('AF'+str(int(celda+1))).value = str(time.strftime("%H:%M:%S"))
        except: 
            winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
            shtTest.range('Q'+str(int(celda+1))+':'+'U'+str(int(celda+1))).value = ''
            print('Error en COMPRA.')
    else: 
        try:
            if len(symbol) < 2:
                orderV = hb.orders.send_sell_order(symbol[0],'24hs', float(precio), int(size))
                shtTest.range('AE'+str(int(celda+1))).value = float(precio)
                try: shtTest.range('V'+str(int(celda+1))).value -= int(size)
                except: shtTest.range('V'+str(int(celda+1))).value = int(size)
                try: shtTest.range('W'+str(int(celda+1))).value -= int(size) * precio*100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio*100
                print(f'SELL {symbol[0]} 24hs // precio: {precio} // - {int(size)} // orden: {orderV}')
            else:
                orderV = hb.orders.send_sell_order(symbol[0],symbol[2], float(precio), int(size))
                shtTest.range('AE'+str(int(celda+1))).value = float(precio/100)
                try: shtTest.range('V'+str(int(celda+1))).value -= int(size)
                except: shtTest.range('V'+str(int(celda+1))).value = int(size)
                try: shtTest.range('W'+str(int(celda+1))).value -= int(size) * precio/100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio/100
                print(f'SELL {symbol[0]} {symbol[2]} // precio: {round(precio/100,4)} // - {int(size)} // orden: {orderV}')
            shtTest.range('AC'+str(int(celda+1))).value = orderV
            shtTest.range('AD'+str(int(celda+1))).value = int(size)
            shtTest.range('AF'+str(int(celda+1))).value = str(time.strftime("%H:%M:%S"))
        except:
            winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
            shtTest.range('Q'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            print('Error en VENTA.')
            
    try: shtTest.range('X'+str(int(celda+1))).value=shtTest.range('W'+str(int(celda+1))).value / shtTest.range('V'+str(int(celda+1))).value
    except: pass
    shtTest.range('Q'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
############################################ TRAILING STOP ################################################
def trailingStop(nombre=str,cantidad=int,nroCelda=int):
    try:
        nombre = str(shtTest.range(str(nombre)).value).split()
        bid = float(shtTest.range('C'+str(int(nroCelda+1))).value)
        bid_size = int(shtTest.range('B'+str(int(nroCelda+1))).value)
        stock = int(shtTest.range('V'+str(int(nroCelda+1))).value)
        last = float(shtTest.range('F'+str(int(nroCelda+1))).value)
        costo = float(shtTest.range('X'+str(int(nroCelda+1))).value) 
        ganancia = shtTest.range('Z1').value
        if not ganancia: 
            ganancia = 0.0005
            shtTest.range('Z1').value = 0.0005
        if cantidad > stock : cantidad = stock
        if cantidad > bid_size : cantidad = bid_size

        if len(nombre) < 2: # OPCIONES _____________________________________________________________________________________
            if bid * 100 > costo * (1 + (ganancia*10)): # Precio sube activo trailing y sube % ganancia 
                shtTest.range('W'+str(int(nroCelda+1))).value = 'TRAILING'
                shtTest.range('X'+str(int(nroCelda+1))).value = bid * 100
            if not shtTest.range('X1').value:
                if last * 100 < costo * (1 - (ganancia*75)): # Precio baja activo stop y envia orden venta
                    queTiene = shtTest.range('W'+str(int(nroCelda+1))).value
                    if str(queTiene) == "CLOSED" : pass
                    else:
                        if str(queTiene) == 'STOP' and bid>last*(1-(ganancia*15)):
                            print(f'{time.strftime("%H:%M:%S")} STOP     ',end=' || ')
                            if shtTest.range('Y'+str(int(nroCelda+1))).value : shtTest.range('X1').value = 'REC'
                            shtTest.range('W'+str(int(nroCelda+1))).value = ''
                            shtTest.range('X'+str(int(nroCelda+1))).value = bid * 100
                            enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),cantidad,nroCelda)
                        else: shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP'  

        else: # BONOS / LETRAS / ON / CEDEARS _______________________________________________________________________________
            if time.strftime("%H:%M:%S") > '16:24:50' and str(nombre[2]).lower() == 'spot': 
                shtTest.range('W'+str(int(nroCelda+1))).value = "CLOSED"
                pass
            if time.strftime("%H:%M:%S") > '16:56:50' and str(nombre[2]).lower() == '24hs': 
                shtTest.range('W'+str(int(nroCelda+1))).value = "CLOSED"
                pass
            if bid / 100 > costo * (1 + ganancia): # Precio sube activo trailing y sube % ganancia               
                shtTest.range('W'+str(int(nroCelda+1))).value = 'TRAILING'
                shtTest.range('X'+str(int(nroCelda+1))).value = round(bid / 100,5)
            if not shtTest.range('X1').value:
                if last / 100 < costo * (1 - ganancia): # Precio baja activo stop y envia orden venta
                    if str(shtTest.range('W'+str(int(nroCelda+1))).value)=='STOP' and (bid/100)>(last/100)*(1-ganancia):
                        print(f'{time.strftime("%H:%M:%S")} STOP     ',end=' || ')
                        if shtTest.range('Y'+str(int(nroCelda+1))).value : shtTest.range('X1').value = 'REC'
                        shtTest.range('W'+str(int(nroCelda+1))).value = ''
                        shtTest.range('X'+str(int(nroCelda+1))).value = round(bid / 100,5)
                        enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),cantidad,nroCelda)
                    else: shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP' 
    except: pass
############################################ BUSCA OPERACIONES ###############################################
def buscoOperaciones(inicio,fin):
    for valor in shtTest.range('P'+str(inicio)+':'+'V'+str(fin)).value:
        cantidad = shtTest.range('Y'+str(int(valor[0]+1))).value
        if cantidad == None: cantidad = 1
        if valor[1]: # COMPRAR precio BID __________________________________________________________________________________
            try:   enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0])
            except: 
                shtTest.range('Q'+str(int(valor[0]+1))).value = ''
                winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
        if valor[2]: # COMPRAR precio ASK _________________________________________________________________________________
            try:  enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0])
            except: 
                shtTest.range('R'+str(int(valor[0]+1))).value = ''
                winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
        if valor[3]: # VENDER precio BID __________________________________________________________________________________
            try:  enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0])
            except: 
                shtTest.range('S'+str(int(valor[0]+1))).value = ''
                winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
        if valor[4]: # VENDER precio ASK ___________________________________________________________________________________
            try:  enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0])
            except: 
                shtTest.range('T'+str(int(valor[0]+1))).value = ''
                winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
        if valor[5]: # CANCELAR todas las ordenes ______________________________________________________________________________
            try: 
                stock = shtTest.range('V'+str(int(valor[0]+1))).value
                cancela = shtTest.range('AD'+str(int(valor[0]+1))).value
                if cancela == None: cancela = 1
                if str(valor[5]).lower() == 'c':
                    orderC = shtTest.range('AB'+str(int(valor[0]+1))).value
                    if orderC == None: orderC = '0000'
                    hb.orders.cancel_order(int(os.environ.get('account_id')),int(orderC))
                    shtTest.range('U'+str(int(valor[0]+1))).value = ''
                    if stock != None: shtTest.range('V'+str(int(valor[0]+1))).value -= shtTest.range('AD'+str(int(valor[0]+1))).value
                    shtTest.range('X'+str(int(valor[0]+1))).value = 0
                    shtTest.range('AB'+str(int(valor[0]+1))).value = ''
                    print(f" // Orden compra: {orderC} fue cancelada x {cancela} // ",time.strftime("%H:%M:%S"))

                elif str(valor[5]).lower() == 'v': 
                    orderV = shtTest.range('AC'+str(int(valor[0]+1))).value
                    if orderV == None : orderV = '0000'
                    hb.orders.cancel_order(int(os.environ.get('account_id')),int(orderV))
                    shtTest.range('U'+str(int(valor[0]+1))).value = ''
                    if stock != None: shtTest.range('V'+str(int(valor[0]+1))).value += shtTest.range('AD'+str(int(valor[0]+1))).value
                    shtTest.range('X'+str(int(valor[0]+1))).value = 0
                    shtTest.range('AC'+str(int(valor[0]+1))).value = ''
                    print(f" // Orden venta : {orderV} fue cancelada x {cancela} // ",time.strftime("%H:%M:%S"))

                elif str(valor[5]).lower() == 'x': 
                    hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
                    shtTest.range('U'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = ''
                    shtTest.range('AB'+str(inicio)+':'+'AF'+str(fin)).value = ''
                    print(" // Todas las ordenes activas canceladas - ",time.strftime("%H:%M:%S") )
            except: 
                winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                shtTest.range('U'+str(int(valor[0]+1))).value = ''
                print(time.strftime("%H:%M:%S"),'Error al cancelar orden.')

            if valor[5] == '-' or valor[5] == '+': # buy//sell usando puntas ______________________________________________________
                if valor[5] == '-':enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidad,valor[0])
                else: enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
                shtTest.range('U'+str(int(valor[0]+1))).value = ''

            if str(valor[5]).upper() == 'P': # PRORTFOLIO  ________________________________________________________________________
                try: getPortfolio(hb, os.environ.get('account_id'))
                except: 
                    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                    print("______ error al traer portfolio ______ ",time.strftime("%H:%M:%S"))
                shtTest.range('U'+str(int(valor[0]+1))).value = ''

        if not shtTest.range('W1').value: # Activa TRAILING  ______________________________________________________________________
            if not valor[6]: pass
            else:
                try: 
                    if valor[6] > 0: trailingStop('A'+str((int(valor[0])+1)),cantidad,int(valor[0]))
                except: pass
                
        if str(shtTest.range('X1').value).upper() == 'REC': # Activa RECOMPRA AUTOMATICA _____________________________________
            try:  enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
            except: 
                winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                print(time.strftime("%H:%M:%S"), 'Error RECOMPRA Automatica.')
########################################### CARGA BUCLE EN EXCEL ##########################################
while True:

    if not shtTest.range('M1').value:
        try: getPortfolio(hb, os.environ.get('account_id'))
        except: 
            winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
            print("______ error al traer portfolio ______ ",time.strftime("%H:%M:%S"))
        shtTest.range('M1').value = 'volume'
    
    if time.strftime("%H:%M:%S") > '17:01:00': 
        getPortfolio(hb, os.environ.get('account_id'))
        break
    if str(shtTest.range('A1').value) != 'symbol': ilRulo()

    try: 
        if not shtTest.range('Q1').value:
            shtTest.range('A'+str(listLength)).options(index=True,header=False).value = everything
            try: shtTest.range('AH2').options(index=True, header=False).value = cauciones
            except: print("______ error al cargar cauciones en Excel ______ ",time.strftime("%H:%M:%S")) 
    except: 
        winsound.PlaySound("SystemHand", winsound.SND_ALIAS) 
        print("______ error al cargar Bonos/Letras en Excel ______ ",time.strftime("%H:%M:%S")) 

    try:
        if not shtTest.range('S1').value: 
            shtTest.range('A30').options(index=True,header=False).value=options  
       #shtTest.range('A26').options(index=True, header=False).value = everything
       #shtTest.range('A' + str(listLength)).options(index=True, header=False).value = options
    except: 
        winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
        print("______ error al cargar OPCIONES en Excel ______ ",time.strftime("%H:%M:%S")) 
        
    buscoOperaciones(rangoDesde,rangoHasta)
    time.sleep(2)
    
try: hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
except: pass
print(time.strftime("%H:%M:%S"), 'Mercado cerrado. ')
shtTest.range('Q1').value = 'BONOS'
shtTest.range('S1').value = 'OPCIONES'
shtTest.range('W1').value = 'TRAILING'
shtTest.range('X1').value = 'STOP'



#[ ]><   \n
#print("\nimprimir en linea nueva")