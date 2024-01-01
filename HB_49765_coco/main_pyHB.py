from pyhomebroker import HomeBroker     
import xlwings as xw                    
import pandas as pd                     
from datetime import date, timedelta
import time
import os
import environ

env = environ.Env()
environ.Env.read_env()

print(time.strftime("%H:%M:%S"),'Abriendo el archvo de Excel ...')

wb = xw.Book('D:\pyHomeBroker\epgb_pyHB.xlsx')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')
shtTest.range('U1:V1').value  = 0
shtTest.range('W1').value  = 10
shtTest.range('Q2:X29').value  = 0

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
listLength = len(options) +26

print(time.strftime("%H:%M:%S"),'Se preparan tickers para solicitar precios ...')

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
    # hb.online.subscribe_options()
    # hb.online.subscribe_securities('bluechips', '48hs')    # Acciones del Panel lider - 48hs
    # hb.online.subscribe_securities('bluechips', '24hs')   # Acciones del Panel lider - 24hs
    # hb.online.subscribe_securities('bluechips', 'SPOT')    # Acciones del Panel lider - spot
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
                #on_options=on_options,
                on_securities=on_securities,
                on_repos=on_repos)

hb.auth.login(dni=str(os.environ.get('dni')),
              user=str(os.environ.get('user')),
              password=str(os.environ.get('password')),
              raise_exception=True)

getGrupos()

def salida():
    hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
    exc = xw.apps.active
    exc.quit()
    hb.online.disconnect()
    exit()
#-------------------------------------------------------------------------------------------------------
print(time.strftime("%H:%M:%S"),f"Logueo correcto COCOS.CAPITAL nro cuenta: {int(os.environ.get('account_id'))}")

def nameArs(name):
    if name == 'SE4D': name = 'S18E4'
    elif name == 'SE4C': name = 'S18E4'
    elif name == 'MRCA': name = 'MRCAO'
    elif name == 'CLSI': name = 'CLSIO'
    elif name == 'BA7D': name = 'BA37D'
    elif name == 'AL30C': name = 'AL30'
    elif name == 'AL30D': name = 'AL30'
    elif name == 'AE38C': name = 'AE38'
    elif name == 'AE38D': name = 'AE38'
    elif name == 'AL29C': name = 'AL29'
    elif name == 'AL29D': name = 'AL29'
    elif name == 'AL35C': name = 'AL35'
    elif name == 'AL35D': name = 'AL35'
    elif name == 'AL41C': name = 'AL41'
    elif name == 'AL41D': name = 'AL41'
    elif name == 'GD29C': name = 'GD29'
    elif name == 'GD29D': name = 'GD29'
    elif name == 'GD35C': name = 'GD35'
    elif name == 'GD35D': name = 'GD35'
    elif name == 'GD38C': name = 'GD38'
    elif name == 'GD38D': name = 'GD38'
    elif name == 'GD41C': name = 'GD41'
    elif name == 'GD41D': name = 'GD41'
    elif name == 'GD46C': name = 'GD46'
    elif name == 'GD46D': name = 'GD46'
    return name
def nameCcl(name):
    if name == 'S18E4': name = 'SE4C'
    elif name == 'S18E4': name = 'SE4C'
    elif name == 'MRCAO': name = 'MRCAC'
    elif name == 'CLSIO': name = 'CLSIC'
    elif name == 'BA37D': name = 'BA7DC'
    elif name == 'AL30': name = 'AL30C'
    elif name == 'GD30': name = 'GD30C'
    elif name == 'AE38': name = 'AE38C'
    elif name == 'AL29': name = 'AL29C'
    elif name == 'AL35': name = 'AL35C'
    elif name == 'AL41': name = 'AL41C'
    elif name == 'GD29': name = 'GD29C'
    elif name == 'GD35': name = 'GD35C'
    elif name == 'GD38': name = 'GD38C'
    elif name == 'GD41': name = 'GD41C'
    elif name == 'GD46': name = 'GD46C'
    return name
def nameMep(name):
    if name == 'S18E4': name = 'SE4D'
    elif name == 'S18E4': name = 'SE4D'
    elif name == 'MRCAO': name = 'MRCAD'
    elif name == 'CLSIO': name = 'CLSID'
    elif name == 'BA37D': name = 'BA7DD'
    elif name == 'AL30': name = 'AL30D'
    elif name == 'GD30': name = 'GD30D'
    elif name == 'AE38': name = 'AE38D'
    elif name == 'AL29': name = 'AL29D'
    elif name == 'AL35': name = 'AL35D'
    elif name == 'AL41': name = 'AL41D'
    elif name == 'GD29': name = 'GD29D'
    elif name == 'GD35': name = 'GD35D'
    elif name == 'GD38': name = 'GD38D'
    elif name == 'GD41': name = 'GD41D'
    elif name == 'GD46': name = 'GD46D'
    return name
   
def ilRulo():
    celda,pesos,dolar = 46,100,10000
    tikers = {
        'cclCI':['tiker',dolar],'ccl48':['tiker',dolar],
        'mepCI':['tiker',dolar],'mep48':['tiker',dolar],
        'arsCI':['tiker',pesos],'ars48':['tiker',pesos]
        }
    for valor in shtTest.range('A46:A141').value:
        ars = shtTest.range('AA'+str(celda)).value
        ccl = shtTest.range('Z'+str(celda)).value
        mep = shtTest.range('Z'+str(celda)).value
        if ars != None and ccl != None and mep != None:
            if (valor[7:8] == 's' or valor[8:9] == 's'):
                if ars > tikers['arsCI'][1]: tikers['arsCI'] = [nameArs(valor[:4])+' - spot',int(ars)]
                if valor[3:4] == 'C' or valor[4:5] == 'C': 
                    if ccl < tikers['cclCI'][1]: tikers['cclCI'] = [nameCcl(valor[:4])+' - spot',int(ccl)]
                if valor[3:4] == 'D' or valor[4:5] == 'D':
                    if mep < tikers['mepCI'][1]: tikers['mepCI'] = [nameMep(valor[:4])+' - spot',int(mep)]
            if (valor[7:9]=='48' or valor[8:10]=='48'):
                if ars > tikers['ars48'][1]: tikers['ars48'] = [nameArs(valor[:4])+' - 48hs',int(ars)]
                if valor[3:4] == 'C' or valor[4:5] == 'C': 
                    if ccl < tikers['ccl48'][1]: tikers['ccl48'] = [nameCcl(valor[:4])+' - 48hs',int(ccl)]
                if valor[3:4] == 'D' or valor[4:5] == 'D': 
                    if mep < tikers['mep48'][1]: tikers['mep48'] = [nameMep(valor[:4])+' - 48hs',int(mep)]
        celda +=1
    print(time.strftime("%H:%M:%S"),tikers)
    # Carga de tikers en planilla excel
    shtTest.range('A2').value = tikers['mepCI'][0]                  
    shtTest.range('A3').value = tikers['cclCI'][0][:4]+'D - spot'   
    shtTest.range('A4').value = tikers['cclCI'][0]                   
    shtTest.range('A5').value = tikers['mepCI'][0][:4]+'C - spot'
    shtTest.range('A6').value = tikers['mepCI'][0]                  
    shtTest.range('A7').value = tikers['arsCI'][0][:4]+'D - spot'   
    shtTest.range('A8').value = tikers['arsCI'][0]                   
    shtTest.range('A9').value = tikers['mepCI'][0][:4]+' - spot'
    shtTest.range('A10').value = tikers['cclCI'][0]                  
    shtTest.range('A11').value = tikers['arsCI'][0][:4]+'C - spot'   
    shtTest.range('A12').value = tikers['arsCI'][0]                   
    shtTest.range('A13').value = tikers['cclCI'][0][:4]+' - spot'

    shtTest.range('A14').value = tikers['mep48'][0]                  
    shtTest.range('A15').value = tikers['ccl48'][0][:4]+'D - 48hs'   
    shtTest.range('A16').value = tikers['ccl48'][0]                   
    shtTest.range('A17').value = tikers['mep48'][0][:4]+'C - 48hs'
    shtTest.range('A18').value = tikers['mep48'][0]                  
    shtTest.range('A19').value = tikers['ars48'][0][:4]+'D - 48hs'   
    shtTest.range('A20').value = tikers['ars48'][0]                   
    shtTest.range('A21').value = tikers['mep48'][0][:4]+' - 48hs'
    shtTest.range('A22').value = tikers['ccl48'][0]                  
    shtTest.range('A23').value = tikers['ars48'][0][:4]+'C - 48hs'   
    shtTest.range('A24').value = tikers['ars48'][0]                   
    shtTest.range('A25').value = tikers['ccl48'][0][:4]+' - 48hs'

def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    symbol = shtTest.range(str(symbol)).value.split()
    mas = shtTest.range('U1').value
    multiplo = float(shtTest.range('W1').value)
    precio = (shtTest.range(str(price)).value + mas)
    precioV = (precio - (mas * 2))
    size *= multiplo
    order = 0
    
    if shtTest.range('V'+str(int(celda+1))).value == 0: 
        shtTest.range('W'+str(int(celda+1))+':'+'X'+str(int(celda+1))).value = 0
    if tipo.lower() == 'buy': 
        if len(symbol) < 2:
            order = hb.orders.send_buy_order(symbol[0],'24hs', float(precio),int(size))
            print(f'Buy  {symbol[0]} 24hs // +{int(size)} // {precio} // ${int(precio*100*size)} // order {order}')
            shtTest.range('V'+str(int(celda+1))).value += int(size)
            shtTest.range('W'+str(int(celda+1))).value += int(size) * float(precio)*100
        else:
            order = hb.orders.send_buy_order(symbol[0],symbol[2],float(precio),int(size))
            print(f'Buy  {symbol[0]} {symbol[2]} // +{int(size)} // {precio} // ${int(precio/100*size)} // order {order}')
            shtTest.range('V'+str(int(celda+1))).value += int(size)
            shtTest.range('W'+str(int(celda+1))).value += int(size) * float(precio)/100
    else: 
        if len(symbol) < 2:
            order = hb.orders.send_sell_order(symbol[0],'24hs', float(precioV),int(size))
            print(f'Sell {symbol[0]} 24hs // -{int(size)} // {precioV} // ${int(precioV*100*size)} // order {order}')
            shtTest.range('V'+str(int(celda+1))).value -= int(size)
            shtTest.range('W'+str(int(celda+1))).value -= int(size) * float(precioV)*100
        else:
            order = hb.orders.send_sell_order(symbol[0],symbol[2],float(precioV),int(size))
            print(f'Sell {symbol[0]} {symbol[2]} // -{int(size)} // {precioV} // ${int(precioV/100*size)} // order {order}')
            shtTest.range('V'+str(int(celda+1))).value -= int(size)
            shtTest.range('W'+str(int(celda+1))).value -= int(size) * float(precioV)/100
    shtTest.range('Q'+str(int(valor[0]+1))+':'+'U'+str(int(valor[0]+1))).value = 0
    if shtTest.range('V'+str(int(celda+1))).value == 0:
        shtTest.range('X'+str(int(celda+1))).value = shtTest.range('W'+str(int(celda+1))).value / 1
    else: 
        shtTest.range('X'+str(int(celda+1))).value = shtTest.range('W'+str(int(celda+1))).value / shtTest.range('V'+str(int(celda+1))).value

print(time.strftime("%H:%M:%S"),"Cargando los precios en la planilla")

while True:
    
    try:
       #shtTest.range('A26').options(index=True, header=False).value = everything
       #shtTest.range('A' + str(listLength)).options(index=True, header=False).value = options
       shtTest.range('A26').options(index=True, header=False).value = options
       shtTest.range('A' + str(listLength)).options(index=True, header=False).value = everything
       shtTest.range('AE2').options(index=True, header=False).value = cauciones
       if time.strftime("%H:%M:%S") <= '10:45:00': continue
       if time.strftime("%H:%M:%S") > '17:05:00': salida() 
    except: print("Error al escribir datos, reconectando Excel ... ",time.strftime("%H:%M:%S"))

    if shtTest.range('Q1').value != 1:
        ilRulo()
        shtTest.range('Q1').value = 1

    for valor in shtTest.range('P2:U29').value:
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
        elif valor[5] == 'c' or valor[5] == 'C': 
            try:
                hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
                shtTest.range('Q'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = 0
                print("Todas las ordenes activas canceladas, verificar multiplo favor")
            except:
                print("Error no fue posible cancelar todas las ordenes activas...")

        # mundo RULOS en automaticoPuntass _______________________________________________________________
        elif valor[5] == '-':
            try:
                shtTest.range('W1').value  = 1
                cantidad= int(shtTest.range('Y'+str(int(valor[0]+1))).value)
                enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
                shtTest.range('U'+str(int(valor[0]+1))).value = 0
                break
            except: shtTest.range('U'+str(int(valor[0]+1))).value = 0
        
        elif valor[5] == '+':
            try:
                shtTest.range('W1').value  = 1
                enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidad,valor[0])
                shtTest.range('U'+str(int(valor[0]+1))).value = 0
                break
            except: shtTest.range('U'+str(int(valor[0]+1))).value = 0

        
#[ ]><   \n
