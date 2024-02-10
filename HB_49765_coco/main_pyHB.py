from pyhomebroker import HomeBroker     
import xlwings as xw                    
import pandas as pd                     
from datetime import date, timedelta
import time
import os
import environ

env = environ.Env()
environ.Env.read_env()

wb = xw.Book('D:\pyHomeBroker\epgb_pyHB.xlsx')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')
shtTest.range('Q1').value = 'N'
shtTest.range('R1').value = 0.05
shtTest.range('U1:V1').value  = 0
shtTest.range('Y30:Z53').value  = 0
shtTest.range('S1').value ='N'
shtTest.range('T1').value ='N'
shtTest.range('W1').value  = 1
'''
def getOptionsList():
    global allOptions
    rng = shtTickers.range('A2:A25').expand()
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
'''
#-------------------------------------------------------------------------------------------------------
print(time.strftime("%H:%M:%S"),f"Logueo en cuenta: {int(os.environ.get('account_id'))} en: {os.environ.get('name')}")

def namesArs(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA37D'+plazo
    elif nombre[:2] == 'BP': return 'BPO27'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') and (nombre[3:4] == 'D' or nombre[3:4] == 'C'):
        if (nombre[1:2] == 'F' or nombre[1:2] == 'Y'): return nombre[:1]+'20'+nombre[1:3]+plazo
        else: return nombre[:1]+'18'+nombre[1:3]+plazo
    elif (nombre[:2] == 'AL' or nombre[:2] == 'GD' or nombre[:2] == 'AE') and (nombre[4:5] == 'D' or nombre[4:5] == 'C'):
        return nombre[:4]+plazo
    else: return nombre[:4]+'O'+plazo

def namesCcl(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA7DC'+plazo
    elif nombre[:2] == 'BP': return 'BP27C'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') :
        if nombre[3:4] == 'D': return nombre[:3]+'C'+plazo
        else: return nombre[:1]+nombre[3:5]+'C'+plazo
    elif (nombre[:2] == 'AL' or nombre[:2] == 'GD' or nombre[:2] == 'AE') and (nombre[4:5] == 'D' or nombre[4:5] == ' '):
        return nombre[:4]+'C'+plazo
    else: return nombre[:4]+'C'+plazo

def namesMep(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA7DD'+plazo
    elif nombre[:2] == 'BP': return 'BP27D'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') :
        if nombre[3:4] == 'C': return nombre[:3]+'D'+plazo
        else: return nombre[:1]+nombre[3:5]+'D'+plazo
    elif (nombre[:2] == 'AL' or nombre[:2] == 'GD' or nombre[:2] == 'AE') and (nombre[4:5] == 'D' or nombre[4:5] == ' '):
        return nombre[:4]+'D'+plazo
    else: return nombre[:4]+'D'+plazo

def cargoPlanilla(dicc):
    if dicc['mepCI'][1] != 10000:
        shtTest.range('A22').value = dicc['mepCI'][0]    
        shtTest.range('Y22').value = dicc['mepCI'][1]
        shtTest.range('Z22').value = namesArs(dicc['mepCI'][0],' - spot')
        shtTest.range('AA22').value =namesCcl(dicc['mepCI'][0],' - spot')
    if dicc['mep48'][1] != 10000:    
        shtTest.range('A23').value = dicc['mep48'][0]
        shtTest.range('Y23').value = dicc['mep48'][1]
        shtTest.range('Z23').value = namesArs(dicc['mep48'][0],' - 48hs')
        shtTest.range('AA23').value =namesCcl(dicc['mep48'][0],' - 48hs')
    if dicc['cclCI'][1] != 10000:
        shtTest.range('A24').value = dicc['cclCI'][0]
        shtTest.range('Y24').value = dicc['cclCI'][1]
        shtTest.range('Z24').value = namesArs(dicc['cclCI'][0],' - spot')
        shtTest.range('AA24').value =namesMep(dicc['cclCI'][0],' - spot')
    if dicc['ccl48'][1] != 10000:
        shtTest.range('A25').value = dicc['ccl48'][0]
        shtTest.range('Y25').value = dicc['ccl48'][1]
        shtTest.range('Z25').value = namesArs(dicc['ccl48'][0],' - 48hs')
        shtTest.range('AA25').value =namesMep(dicc['ccl48'][0],' - 48hs')

    if dicc['arsCImep'][1] != 100:
        shtTest.range('A26').value = dicc['arsCImep'][0]
        shtTest.range('Y26').value = dicc['arsCImep'][1]
        shtTest.range('Z26').value = namesMep(dicc['arsCImep'][0],' - spot')
        shtTest.range('AA26').value =namesCcl(dicc['arsCImep'][0],' - spot')
    if dicc['ars48mep'][1] != 100:
        shtTest.range('A27').value = dicc['ars48mep'][0]
        shtTest.range('Y27').value = dicc['ars48mep'][1]
        shtTest.range('Z27').value = namesMep(dicc['ars48mep'][0],' - 48hs')
        shtTest.range('AA27').value =namesCcl(dicc['ars48mep'][0],' - 48hs')
    if dicc['arsCIccl'][1] != 100:
        shtTest.range('A28').value = dicc['arsCIccl'][0]
        shtTest.range('Y28').value = dicc['arsCIccl'][1]
        shtTest.range('Z28').value = namesCcl(dicc['arsCIccl'][0],' - spot')
        shtTest.range('AA28').value =namesMep(dicc['arsCIccl'][0],' - spot')
    if dicc['ars48ccl'][1] != 100:
        shtTest.range('A29').value = dicc['ars48ccl'][0]
        shtTest.range('Y29').value = dicc['ars48ccl'][1]
        shtTest.range('Z29').value = namesCcl(dicc['ars48ccl'][0],' - 48hs')
        shtTest.range('AA29').value =namesMep(dicc['ars48ccl'][0],' - 48hs') 

def limpio():
    shtTest.range('A10:A17').value = ''
    shtTest.range('A22:A29').value = ''
    shtTest.range('Y22:AA29').value = ''

def cargoXplazo(dicc):
    if time.strftime("%H:%M:%S") > '16:26:00':
        shtTest.range('A10').value = dicc['mep48'][0] # mep
        shtTest.range('A11').value = namesMep(dicc['ars48mep'][0],' - 48hs') #  mep
        shtTest.range('A12').value = dicc['ars48mep'][0] #  ars
        shtTest.range('A13').value = namesArs(dicc['mep48'][0],' - 48hs') # ars
        shtTest.range('A14').value = dicc['mep48'][0] # mep
        shtTest.range('A15').value = namesMep(dicc['ccl48'][0],' - 48hs')
        shtTest.range('A16').value = dicc['ccl48'][0] # ccl
        shtTest.range('A17').value = namesCcl(dicc['mep48'][0],' - 48hs')
    else:
        shtTest.range('A10').value = dicc['mepCI'][0]
        shtTest.range('A11').value = namesMep(dicc['arsCImep'][0],' - spot')
        shtTest.range('A12').value = dicc['arsCImep'][0]
        shtTest.range('A13').value = namesArs(dicc['mepCI'][0],' - spot') 
        shtTest.range('A14').value = dicc['mepCI'][0] # mep
        shtTest.range('A15').value = namesMep(dicc['cclCI'][0],' - spot')
        shtTest.range('A16').value = dicc['cclCI'][0] # ccl
        shtTest.range('A17').value = namesCcl(dicc['mepCI'][0],' - spot')

def ilRulo():
    shtTest.range('A1').value = 'symbol'
    limpio()
    celda,pesos,dolar = 46,1000,0
    tikers = {'cclCI':['',dolar],'ccl48':['',dolar],'mepCI':['',dolar],'mep48':['',dolar],'arsCIccl':['',pesos],'ars48ccl':['',pesos],'arsCImep':['',pesos],'ars48mep':['',pesos]}
    
    for valor in shtTest.range('A46:A153').value:
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
    cargoPlanilla(tikers)
    
def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    global orderC,orderV
    symbol = shtTest.range(str(symbol)).value.split()
    mas = shtTest.range('U1').value
    multiplo = float(shtTest.range('W1').value)
    precio = (shtTest.range(str(price)).value + mas)
    precioV = (precio - (mas * 2))
    size *= multiplo
    orderC, orderV = 0,0
    if shtTest.range('V'+str(int(celda+1))).value == 0: 
        shtTest.range('W'+str(int(celda+1))+':'+'X'+str(int(celda+1))).value = 0
    if tipo.lower() == 'buy': 
        if len(symbol) < 2:
            #orderC = hb.orders.send_buy_order(symbol[0],'24hs', float(precio),int(size))
            print(f'Buy  {symbol[0]} 24hs // +{int(size)} // {precio} // ${int(precio*100*size)} // order {orderC}')
            shtTest.range('V'+str(int(celda+1))).value += int(size)
            shtTest.range('W'+str(int(celda+1))).value += int(size) * float(precio)*100
        else:
            #orderC = hb.orders.send_buy_order(symbol[0],symbol[2],float(precio),int(size))
            print(f'Buy  {symbol[0]} {symbol[2]} // +{int(size)} // {precio} // ${int(precio/100*size)} // order {orderC}')
            shtTest.range('V'+str(int(celda+1))).value += int(size)
            shtTest.range('W'+str(int(celda+1))).value += int(size) * float(precio)/100
    else: 
        if len(symbol) < 2:
            #orderV = hb.orders.send_sell_order(symbol[0],'24hs', float(precioV),int(size))
            print(f'Sell {symbol[0]} 24hs // -{int(size)} // {precioV} // ${int(precioV*100*size)} // order {orderV}')
            shtTest.range('V'+str(int(celda+1))).value -= int(size)
            shtTest.range('W'+str(int(celda+1))).value -= int(size) * float(precioV)*100
        else:
            #orderV = hb.orders.send_sell_order(symbol[0],symbol[2],float(precioV),int(size))
            print(f'Sell {symbol[0]} {symbol[2]} // -{int(size)} // {precioV} // ${int(precioV/100*size)} // order {orderV}')
            shtTest.range('V'+str(int(celda+1))).value -= int(size)
            shtTest.range('W'+str(int(celda+1))).value -= int(size) * float(precioV)/100
    shtTest.range('Q'+str(int(valor[0]+1))+':'+'U'+str(int(valor[0]+1))).value = 0
    if shtTest.range('V'+str(int(celda+1))).value == 0:
        shtTest.range('X'+str(int(celda+1))).value = shtTest.range('W'+str(int(celda+1))).value / 1
    else: 
        shtTest.range('X'+str(int(celda+1))).value = shtTest.range('W'+str(int(celda+1))).value / shtTest.range('V'+str(int(celda+1))).value

####################################### TRAILING STOP ################################################
def trailing(nombre=str,cantidad=int,nroCelda=int):
    try:
        bid = shtTest.range('C'+str(int(nroCelda+1))).value
        costo = shtTest.range('X'+str(int(nroCelda+1))).value 
        nombre = shtTest.range(str(nombre)).value.split()
        ganancia = shtTest.range('R1').value
        if bid * 100 > costo * (1 + ganancia)  :

            if shtTest.range('Q1').value == 'T':
                #orderV = hb.orders.send_sell_order(nombre[0],'24hs', float(bid),int(cantidad))
                print('vendo x trailing', nombre[0],'24hs', float(bid),int(cantidad))
                shtTest.range('Q1').value = ''

            shtTest.range('Z'+str(int(nroCelda+1))).value = costo
            shtTest.range('X'+str(int(nroCelda+1))).value = bid * 100     
    except: pass
#########################################################################################################

while True:
    '''if time.strftime("%H:%M:%S") > '17:06:00': break 
    try:
       #shtTest.range('A26').options(index=True, header=False).value = everything
       #shtTest.range('A' + str(listLength)).options(index=True, header=False).value = options
       if shtTest.range('T1').value!='N':shtTest.range('A30').options(index=True,header=False).value=options
       if shtTest.range('S1').value!='N':shtTest.range('A'+str(listLength)).options(index=True,header=False).value = everything
       shtTest.range('AE2').options(index=True, header=False).value = cauciones
       if time.strftime("%H:%M:%S") <= '10:45:00': continue
    except: print("_____ error al cargar datos en Excel !!! ",time.strftime("%H:%M:%S"))'''

    if shtTest.range('A1').value != 'symbol': ilRulo()
    #time.sleep(10)
    for valor in shtTest.range('P30:V53').value:
        if shtTest.range('Q1').value != 'N' and valor[6] != 0: 
            trailing('A'+str((int(valor[0])+1)),valor[6],valor[0])
        
        if valor[1] != 0: # COMPRAR precio BID ___________________________________________________________
            try: 
                enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0])
            except: shtTest.range('Q'+str(int(valor[0]+1))+':'+'T'+str(int(valor[0]+1))).value = 0
        elif valor[2] != 0: # COMPRAR precio ASK _________________________________________________________
            try: 
                enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0])
            except: shtTest.range('Q'+str(int(valor[0]+1))+':'+'T'+str(int(valor[0]+1))).value = 0
        elif valor[3] != 0: # VENDER precio BID __________________________________________________________
            try:
                enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0])
            except: shtTest.range('Q'+str(int(valor[0]+1))+':'+'T'+str(int(valor[0]+1))).value = 0
        elif valor[4] != 0: # VENDER precio ASK __________________________________________________________
            try:
                enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0])
            except: shtTest.range('Q'+str(int(valor[0]+1))+':'+'T'+str(int(valor[0]+1))).value = 0

        # CANCELAR todas las ordenes _____________________________________________________________________
        try: 
            if valor[5] == 'c' or valor[5] == 'C': 
                #hb.orders.cancel_order(int(os.environ.get('account_id')),orderC)
                shtTest.range('Q'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = 0
                print(f"Orden compra {orderC} fue cancelada")
            elif valor[5] == 'v' or valor[5] == 'V': 
                #hb.orders.cancel_order(int(os.environ.get('account_id')),orderV)
                shtTest.range('Q'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = 0
                print(f"Orden venta {orderV} fue cancelada")
            elif valor[5] == 'x' or valor[5] == 'X': 
                #hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
                shtTest.range('Q'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = 0
                print("Todas las ordenes activas canceladas")
        except: 
            shtTest.range('U'+str(int(valor[0]+1))).value = 0
            print('Error, al cancelar orden.')

        # mundo RULOS en automaticoPuntass _______________________________________________________________
        '''elif valor[5] == '-':
            try:
                shtTest.range('W1').value  = 1
                cantidad= int(shtTest.range('Y'+str(int(valor[0]+1))).value)
                enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
                shtTest.range('U'+str(int(valor[0]+1))).value = 0
            except: shtTest.range('U'+str(int(valor[0]+1))).value = 0
        
        elif valor[5] == '+':
            try:
                shtTest.range('W1').value  = 1
                enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidad,valor[0])
                shtTest.range('U'+str(int(valor[0]+1))).value = 0
            except: shtTest.range('U'+str(int(valor[0]+1))).value = 0
            '''
# print(time.strftime("%H:%M:%S"), 'Mercado cerrado.')
  
#[ ]><   \n
