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
shtTest.range('W1').value  = 100
shtTest.range('Q2:X25').value  = 0

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
    global cauciones, caucionesD
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
    #hb.online.subscribe_options()
    #hb.online.subscribe_securities('bluechips', '48hs')    # Acciones del Panel lider - 48hs
    # hb.online.subscribe_securities('bluechips', '24hs')   # Acciones del Panel lider - 24hs
    #hb.online.subscribe_securities('bluechips', 'SPOT')    # Acciones del Panel lider - spot
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

    for valor in shtTest.range('P2:U25').value:
        if valor[1] != 0: # COMPRAR precio BID ___________________________________________________________
            try: 
                enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0])
                break
            except: shtTest.range('Q'+str(int(valor[0]+1))+':'+'T'+str(int(valor[0]+1))).value = 0
        elif valor[2] != 0: # COMPRAR precio ASK _________________________________________________________
            try: 
                enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0])
                break
            except: shtTest.range('Q'+str(int(valor[0]+1))+':'+'T'+str(int(valor[0]+1))).value = 0
        elif valor[3] != 0: # VENDER precio BID __________________________________________________________
            try:
                enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0])
                break
            except: shtTest.range('Q'+str(int(valor[0]+1))+':'+'T'+str(int(valor[0]+1))).value = 0
        elif valor[4] != 0: # VENDER precio ASK __________________________________________________________
            try:
                enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0])
                break
            except: shtTest.range('Q'+str(int(valor[0]+1))+':'+'T'+str(int(valor[0]+1))).value = 0
        elif valor[5] == 'c' or valor[5] == 'C': # CANCELAR todas las ordenes ____________________________
            try:
                hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
                shtTest.range('Q'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = 0
                print("Todas las ordenes activas canceladas, verificar multiplo favor")
                break
            except:
                print("Error no fue posible cancelar todas las ordenes activas...")
        
        # mundo RULOS en automaticoPuntass _______________________________________________________________
        elif valor[5] == 4:
            try:
                shtTest.range('W1').value  = 1
                cantidad= int(shtTest.range('Y'+str(int(valor[0]+1))).value)
                enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidad,valor[0])
                shtTest.range('U'+str(int(valor[0]+1))).value = 0
                break
            except: shtTest.range('U'+str(int(valor[0]+1))).value = 0
        
        elif valor[5] == 3:
            try:
                shtTest.range('W1').value  = 1
                cantidad= int(shtTest.range('Y'+str(int(valor[0]+1))).value)
                enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
                shtTest.range('U'+str(int(valor[0]+1))).value = 0
                break
            except: shtTest.range('U'+str(int(valor[0]+1))).value = 0

        elif valor[5] == 2:
            try:
                shtTest.range('W1').value  = 1
                enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidad,valor[0])
                shtTest.range('U'+str(int(valor[0]+1))).value = 0
                break
            except: shtTest.range('U'+str(int(valor[0]+1))).value = 0
        
        elif valor[5] == 1:
            try:
                shtTest.range('W1').value  = 1
                enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
                shtTest.range('U'+str(int(valor[0]+1))).value = 0
                break
            except: shtTest.range('U'+str(int(valor[0]+1))).value = 0
        


# r2zGLem7KtxtE4b // git
#[ ]><   \n
