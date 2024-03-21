from pyhomebroker import HomeBroker     
import xlwings as xw                    
import pandas as pd                     
from datetime import date, timedelta
import time
import winsound
import os
import environ

env = environ.Env()
environ.Env.read_env()
wb = xw.Book('D:\\pyHomeBroker\\epgb_pyHB.xlsx')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')
shtTest.range('Q1').value = 'PRC'
shtTest.range('R1').value ='TRAIL'
shtTest.range('S1').value ='STOP'
shtTest.range('T1').value = 0.002
shtTest.range('U1').value = 5
shtTest.range('V1').value = 0
shtTest.range('W1').value = 1


def getOptionsList():
    global allOptions
    rng = shtTickers.range('A2:A35').expand()
    oOpciones = rng.value
    allOptions = pd.DataFrame({'symbol': oOpciones},columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last","change", "open", "high", "low", "previous_close", "turnover", "volume",'operations', 'datetime'])
    allOptions = allOptions.set_index('symbol')
    allOptions['datetime'] = pd.to_datetime(allOptions['datetime'])
    return allOptions

def getBonosList():
    rng = shtTickers.range('E2:E115').expand()
    oBonos = rng.value
    Bonos = pd.DataFrame({'symbol' : oBonos}, columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last", "change", "open", "high", "low", "previous_close", "turnover", "volume", 'operations', 'datetime'])
    Bonos = Bonos.set_index('symbol')
    Bonos['datetime'] = pd.to_datetime(Bonos['datetime'])
    return Bonos

def getAccionesList():
    rng = shtTickers.range('C2:C70').expand()
    oAcciones = rng.value
    ACC = pd.DataFrame({'symbol' : oAcciones}, columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last", "change", "open", "high", "low", "previous_close", "turnover", "volume", 'operations', 'datetime'])
    ACC = ACC.set_index('symbol')
    ACC['datetime'] = pd.to_datetime(ACC['datetime'])
    return ACC

i = 1
fechas = []
while i < 8:
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
#-------------------------------------------------------------------------------------------------------
def getGrupos():
    hb.online.connect()
    hb.online.subscribe_options()
    hb.online.subscribe_securities('bluechips', '48hs')    # Acciones del Panel lider - 48hs
    # hb.online.subscribe_securities('bluechips', '24hs')   # Acciones del Panel lider - 24hs
    hb.online.subscribe_securities('bluechips', 'SPOT')    # Acciones del Panel lider - spot
    hb.online.subscribe_securities('government_bonds', '48hs')  # Bonos - 48hs
    # hb.online.subscribe_securities('government_bonds', '24hs') # Bonos - 24hs
    hb.online.subscribe_securities('government_bonds', 'SPOT')  # Bonos - spot
    # hb.online.subscribe_securities('cedears', '48hs')      # CEDEARS - 48hs
    # hb.online.subscribe_securities('cedears', '24hs')      # CEDEARS - 24hs
    # hb.online.subscribe_securities('cedears', 'SPOT')      # CEDEARS - spot
    # hb.online.subscribe_securities('general_board', '48hs') # Acciones del Panel general - 48hs
    # hb.online.subscribe_securities('general_board', '24hs') # Acciones del Panel general - 24hs
    # hb.online.subscribe_securities('general_board', 'SPOT') # Acciones del Panel general - spot
    hb.online.subscribe_securities('short_term_government_bonds', '48hs')   # LETRAS - 48hs
    # hb.online.subscribe_securities('short_term_government_bonds', '24hs')  # LETRAS - 24hs
    hb.online.subscribe_securities('short_term_government_bonds', 'SPOT')   # LETRAS - spot
    hb.online.subscribe_securities('corporate_bonds', '48hs')  # Obligaciones Negociables - 48hs
    # hb.online.subscribe_securities('corporate_bonds', '24hs')  # Obligaciones Negociables - 24hs
    hb.online.subscribe_securities('corporate_bonds', 'SPOT')  # Obligaciones Negociables - spot
    hb.online.subscribe_repos()

hb = HomeBroker(int(os.environ.get('broker')),
                on_options=on_options,
                on_securities=on_securities,
                on_repos=on_repos)

hb.auth.login(dni=str(os.environ.get('dni')),
              user=str(os.environ.get('user')),
              password=str(os.environ.get('password')),
              raise_exception=True)

getGrupos()

#-------------------------------------------------------------------------------------------------------
print(time.strftime("%H:%M:%S"),f"Logueo correcto en cuenta: {int(os.environ.get('account_id'))} broker: {os.environ.get('name')}")

def namesArs(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA37D'+plazo
    elif nombre[:2] == 'BP': return 'BPOA7'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') and (nombre[3:4] == 'D' or nombre[3:4] == 'C'):
        if (nombre[1:2] == 'F' or nombre[1:2] == 'Y'): return nombre[:1]+'20'+nombre[1:3]+plazo
        else: return nombre[:1]+'18'+nombre[1:3]+plazo
    elif (nombre[:2] == 'AL' or nombre[:2] == 'GD' or nombre[:2] == 'AE') and (nombre[4:5] == 'D' or nombre[4:5] == 'C'):
        return nombre[:4]+plazo
    else: return nombre[:4]+'O'+plazo

def namesCcl(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA7DC'+plazo
    elif nombre[:2] == 'BP': return 'BPA7C'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') :
        if nombre[3:4] == 'D': return nombre[:3]+'C'+plazo
        else: return nombre[:1]+nombre[3:5]+'C'+plazo
    elif (nombre[:2] == 'AL' or nombre[:2] == 'GD' or nombre[:2] == 'AE') and (nombre[4:5] == 'D' or nombre[4:5] == ' '):
        return nombre[:4]+'C'+plazo
    else: return nombre[:4]+'C'+plazo

def namesMep(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA7DD'+plazo
    elif nombre[:2] == 'BP': return 'BPA7D'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') :
        if nombre[3:4] == 'C': return nombre[:3]+'D'+plazo
        else: return nombre[:1]+nombre[3:5]+'D'+plazo
    elif (nombre[:2] == 'AL' or nombre[:2] == 'GD' or nombre[:2] == 'AE') and (nombre[4:5] == 'D' or nombre[4:5] == ' '):
        return nombre[:4]+'D'+plazo
    else: return nombre[:4]+'D'+plazo

def cargoXplazo(dicc):
    shtTest.range('A1').value = 'symbol'
    if time.strftime("%H:%M:%S") > '16:26:00':
        # Carga el mejor pagador de MEP 48hs  
        shtTest.range('A2').value = dicc['mep48'][0]
        shtTest.range('A3').value = 'AL30D - 48hs'
        shtTest.range('A4').value = 'AL30 - 48hs'
        shtTest.range('A5').value = namesArs(dicc['mep48'][0],' - 48hs')# Ticker en ARS para reiniciar rulo en 48hs
        shtTest.range('A6').value = dicc['mep48'][0]
        shtTest.range('A7').value = 'GD30D - 48hs'
        shtTest.range('A8').value = 'GD30 - 48hs'
        shtTest.range('A9').value = namesArs(dicc['mep48'][0],' - 48hs')# Ticker en ARS para reiniciar rulo en 48hs
        shtTest.range('A10').value = dicc['mep48'][0]
        shtTest.range('A11').value = namesMep(dicc['ars48mep'][0],' - 48hs') #  mep
        shtTest.range('A12').value = dicc['ars48mep'][0] #  ars
        shtTest.range('A13').value = namesArs(dicc['mep48'][0],' - 48hs')# Ticker en ARS para reiniciar rulo en 48hs
        shtTest.range('A14').value = 'AL30C - 48hs'
        shtTest.range('A15').value = 'GD30C - 48hs'
        shtTest.range('A16').value = 'GD30D - 48hs'
        shtTest.range('A17').value = 'AL30D - 48hs'
        shtTest.range('A18').value = dicc['mep48'][0] # mep
        shtTest.range('A19').value = namesMep(dicc['ccl48'][0],' - 48hs')
        shtTest.range('A20').value = dicc['ccl48'][0] # ccl
        shtTest.range('A21').value = namesCcl(dicc['mep48'][0],' - 48hs')
    else:
        # Carga el mejor pagador de MEP CI  
        shtTest.range('A2').value = dicc['mepCI'][0]
        shtTest.range('A3').value = 'AL30D - spot'
        shtTest.range('A4').value = 'AL30 - spot'
        shtTest.range('A5').value = namesArs(dicc['mepCI'][0],' - spot')
        shtTest.range('A6').value = dicc['mepCI'][0]
        shtTest.range('A7').value = 'GD30D - spot'
        shtTest.range('A8').value = 'GD30 - spot'
        shtTest.range('A9').value = namesArs(dicc['mepCI'][0],' - spot')
        shtTest.range('A10').value = dicc['mepCI'][0]
        shtTest.range('A11').value = namesMep(dicc['arsCImep'][0],' - spot')
        shtTest.range('A12').value = dicc['arsCImep'][0]
        shtTest.range('A13').value = namesArs(dicc['mepCI'][0],' - spot')
        shtTest.range('A14').value = 'AL30C - spot'
        shtTest.range('A15').value = 'GD30C - spot'
        shtTest.range('A16').value = 'GD30D - spot'
        shtTest.range('A17').value = 'AL30D - spot'
        shtTest.range('A18').value = dicc['mepCI'][0] # mep
        shtTest.range('A19').value = namesMep(dicc['cclCI'][0],' - spot')
        shtTest.range('A20').value = dicc['cclCI'][0] # ccl
        shtTest.range('A21').value = namesCcl(dicc['mepCI'][0],' - spot')
    winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)

def ilRulo():
    celda,pesos,dolar = 64,1000,0
    tikers = {'cclCI':['',dolar],'ccl48':['',dolar],'mepCI':['',dolar],'mep48':['',dolar],'arsCIccl':['',pesos],'ars48ccl':['',pesos],'arsCImep':['',pesos],'ars48mep':['',pesos]}
    for valor in shtTest.range('A64:A165').value:
        if not valor: continue
        arsM = shtTest.range('AA'+str(celda)).value
        if arsM == None: arsM = 1000
        arsC = arsM
        ccl = shtTest.range('Z'+str(celda)).value
        if ccl == None: ccl = 0
        mep = ccl
        if (valor[7:8] == 's' or valor[8:9] == 's'):
            if valor[3:4] == 'C' or valor[4:5] == 'C': 
                if arsC > tikers['arsCIccl'][1]: tikers['arsCIccl'] = [namesArs(valor[:5],' - spot'),arsC]
                if ccl > tikers['cclCI'][1]: tikers['cclCI'] = [valor,ccl]
            if valor[3:4] == 'D' or valor[4:5] == 'D':
                if arsM > tikers['arsCImep'][1]: tikers['arsCImep'] = [namesArs(valor[:5],' - spot'),arsM]
                if mep > tikers['mepCI'][1]: tikers['mepCI'] = [valor,mep]
        if (valor[7:9]=='48' or valor[8:10]=='48'):
            if valor[3:4] == 'C' or valor[4:5] == 'C': 
                if arsC > tikers['ars48ccl'][1]: tikers['ars48ccl'] = [namesArs(valor[:5],' - 48hs'),arsC]
                if ccl > tikers['ccl48'][1]: tikers['ccl48'] = [valor,ccl]
            if valor[3:4] == 'D' or valor[4:5] == 'D': 
                if arsM > tikers['ars48mep'][1]: tikers['ars48mep'] = [namesArs(valor[:5],' - 48hs'),arsM]
                if mep > tikers['mep48'][1]: tikers['mep48'] = [valor,mep]
        celda +=1
    cargoXplazo(tikers)

############################################ ENVIAR ORDENES ################################################    
def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    global orderC, orderV
    symbol = str(shtTest.range(str(symbol)).value).split()
    por = int(shtTest.range('W1').value)
    precio = shtTest.range(str(price)).value
    menosRecompra = float(shtTest.range('U1').value)
    if not shtTest.range('V'+str(int(celda+1))).value: shtTest.range('V'+str(int(celda+1))+':'+'X'+str(int(celda+1))).value = 0
    if tipo.lower() == 'buy': 
        if len(symbol) < 2:
            if str(shtTest.range('R1').value) == 'REC': 
                if not menosRecompra: 
                    precio -= 1
                    shtTest.range('U1').value = 10
                else:  precio -= menosRecompra / 10
                shtTest.range('R1').value = ''
            orderC = hb.orders.send_buy_order(symbol[0],'24hs', float(precio), int(size))
            print(f'Buy  {symbol[0]} // cantidad: + {int(size)} // precio: {precio}')
            try: shtTest.range('V'+str(int(celda+1))).value += int(size)
            except: shtTest.range('V'+str(int(celda+1))).value = int(size)
            try: shtTest.range('W'+str(int(celda+1))).value += int(size) * precio*100
            except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio*100
        else:
            if str(shtTest.range('R1').value) == 'REC': 
                if not menosRecompra: 
                    precio -= 100
                    shtTest.range('U1').value = 10
                else:  precio -= menosRecompra * 10
                shtTest.range('R1').value = ''
            orderC = hb.orders.send_buy_order(symbol[0],symbol[2], float(precio), int(size*por))
            print(f'Buy  {symbol[0]} {symbol[2]} // cantidad: + {int(size*por)} // precio {round(precio/100,2)}')
            try: shtTest.range('V'+str(int(celda+1))).value += int(size*por)
            except: shtTest.range('V'+str(int(celda+1))).value = int(size*por)
            try: shtTest.range('W'+str(int(celda+1))).value += int(size*por) * precio/100
            except: shtTest.range('W'+str(int(celda+1))).value = int(size*por) * precio/100
    else: 
        if len(symbol) < 2:
            orderV = hb.orders.send_sell_order(symbol[0],'24hs', float(precio), int(size))
            print(f'Sell {symbol[0]} // cantidad: - {int(size)} // precio: {precio}')
            try: shtTest.range('V'+str(int(celda+1))).value -= int(size)
            except: shtTest.range('V'+str(int(celda+1))).value = int(size)
            try: shtTest.range('W'+str(int(celda+1))).value -= int(size) * precio*100
            except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio*100
        else:
            orderV = hb.orders.send_sell_order(symbol[0],symbol[2], float(precio), int(size*por))
            print(f'Sell {symbol[0]} {symbol[2]} // cantidad: - {int(size*por)} // precio: {round(precio/100,2)}')
            try: shtTest.range('V'+str(int(celda+1))).value -= int(size*por)
            except: shtTest.range('V'+str(int(celda+1))).value = int(size*por)
            try: shtTest.range('W'+str(int(celda+1))).value -= int(size*por) * precio/100
            except: shtTest.range('W'+str(int(celda+1))).value = int(size*por) * precio/100

    shtTest.range('X'+str(int(celda+1))).value=shtTest.range('W'+str(int(celda+1))).value / shtTest.range('V'+str(int(celda+1))).value
    shtTest.range('Q'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
    winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
############################################ TRAILING STOP ################################################
def trailingStop(nombre=str,cantidad=int,nroCelda=int):
    try:
        nombre = str(shtTest.range(str(nombre)).value).split()
        bid = float(shtTest.range('C'+str(int(nroCelda+1))).value)
        bid_size = int(shtTest.range('B'+str(int(nroCelda+1))).value)
        stock = int(shtTest.range('V'+str(int(nroCelda+1))).value)
        last = float(shtTest.range('F'+str(int(nroCelda+1))).value)
        costo = float(shtTest.range('X'+str(int(nroCelda+1))).value) 
        try: ganancia = float(shtTest.range('T1').value)
        except:
            shtTest.range('T1').value = 0.002
            ganancia = float(shtTest.range('T1').value)
        if cantidad > stock : cantidad = stock
        if cantidad > bid_size : cantidad = bid_size
        if len(nombre) < 2: #TRAILING sobre opciones financieras
            if bid * 100 > costo * (1 + (ganancia*25)): # Precio sube activo trailing y sube % ganancia 
                shtTest.range('W'+str(int(nroCelda+1))).value = 'TRAILING'
                shtTest.range('X'+str(int(nroCelda+1))).value = bid * 100
            else: shtTest.range('W'+str(int(nroCelda+1))).value = ''
            if not shtTest.range('S1').value:
                if last * 100 < costo * (1 - (ganancia*10)): # Precio baja activo stop y envia orden venta
                    if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP' and bid>last*(1-(ganancia*10)):
                        print(f'{time.strftime("%H:%M:%S")} STOP ',end=' || ')
                        shtTest.range('R1').value = 'REC'
                        shtTest.range('W'+str(int(nroCelda+1))).value = ''
                        shtTest.range('X'+str(int(nroCelda+1))).value = bid * 100
                        enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),cantidad,nroCelda)
                    else: shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP'  
        else: #TRAILING sobre bonos / letras / ons
            if bid / 100 > costo * (1 + ganancia): # Precio sube activo trailing y sube % ganancia               
                shtTest.range('W'+str(int(nroCelda+1))).value = 'TRAILING'
                shtTest.range('X'+str(int(nroCelda+1))).value = round(bid / 100,5)
            else: shtTest.range('W'+str(int(nroCelda+1))).value = ''
            if not shtTest.range('S1').value:
                if last / 100 < costo * (1 - ganancia): # Precio baja activo stop y envia orden venta
                    if str(shtTest.range('W'+str(int(nroCelda+1))).value)=='STOP' and (bid/100)>(last/100)*(1-ganancia):
                        print(f'{time.strftime("%H:%M:%S")} STOP ',end=' || ')
                        shtTest.range('R1').value = 'REC'
                        shtTest.range('W'+str(int(nroCelda+1))).value = ''
                        shtTest.range('X'+str(int(nroCelda+1))).value = round(bid / 100,5)
                        enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),cantidad,nroCelda)
                    else: shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP' 
    except: pass
########################################### CARGA BUCLE EN EXCEL ##########################################
while True:
    time.sleep(2)
    if time.strftime("%H:%M:%S") > '17:03:00': break 
    if str(shtTest.range('A1').value) != 'symbol': ilRulo()
    try:
        if not shtTest.range('Q1').value:
            shtTest.range('A30').options(index=True,header=False).value=options
            shtTest.range('A'+str(listLength)).options(index=True,header=False).value = everything
            shtTest.range('AE2').options(index=True, header=False).value = cauciones
       #shtTest.range('A26').options(index=True, header=False).value = everything
       #shtTest.range('A' + str(listLength)).options(index=True, header=False).value = options
    except: print("_____ error al cargar datos en Excel !!! ",time.strftime("%H:%M:%S"))
    
    for valor in shtTest.range('P22:V59').value:
        if not shtTest.range('R1').value: # Activa TRAILING  __________________________________________
            try: stock = int(valor[6])
            except: stock = 0
            if stock > 0:
                if not shtTest.range('Y'+str(int(valor[0]+1))).value: cantidad = 1
                else: cantidad = int(shtTest.range('Y'+str(int(valor[0]+1))).value)
                trailingStop('A'+str((int(valor[0])+1)),cantidad,valor[0])
        if str(shtTest.range('R1').value).upper() == 'REC': # Activa RECOMPRA AUTOMATICA _____________
            try:   enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
            except: 
                shtTest.range('R1').value = ''
                print('Error no se ejecuta Recompra Automatica. Corregir valor en celda Y')
        if valor[1]: # COMPRAR precio BID _________________________________________________________________
            try:   enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0])
            except: shtTest.range('Q'+str(int(valor[0]+1))).value = ''
        if valor[2]: # COMPRAR precio ASK _______________________________________________________________
            try:  enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0])
            except: shtTest.range('R'+str(int(valor[0]+1))).value = ''
        if valor[3]: # VENDER precio BID ________________________________________________________________
            try:  enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0])
            except: shtTest.range('S'+str(int(valor[0]+1))).value = ''
        if valor[4]: # VENDER precio ASK ________________________________________________________________
            try:  enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0])
            except: shtTest.range('T'+str(int(valor[0]+1))).value = ''
        if valor[5]:
            try: # CANCELAR todas las ordenes _____________________________________________________________
                if str(valor[5]).lower() == 'c': 
                    hb.orders.cancel_order(int(os.environ.get('account_id')),orderC)
                    shtTest.range('U'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = ''
                    print("Orden compra fue cancelada")
                    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                elif str(valor[5]).lower() == 'v': 
                    hb.orders.cancel_order(int(os.environ.get('account_id')),orderV)
                    shtTest.range('U'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = ''
                    print("Orden venta fue cancelada")
                    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                elif str(valor[5]).lower() == 'x': 
                    hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
                    shtTest.range('U'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = ''
                    print("Todas las ordenes activas canceladas")
                    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
            except: 
                shtTest.range('U'+str(int(valor[0]+1))).value = ''
                print('Error, al cancelar orden.')
            if valor[5] == '-' or valor[5] == '+': # buy//sell usando puntas ______________________________
                try: 
                    cantidad = int(shtTest.range('V'+str(int(valor[0]+1))).value)
                except: cantidad = 1
                if valor[5] == '-':enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
                else: enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidad,valor[0])
                shtTest.range('U'+str(int(valor[0]+1))).value = ''
print(time.strftime("%H:%M:%S"), 'Mercado cerrado.')
  
#[ ]><   \n
