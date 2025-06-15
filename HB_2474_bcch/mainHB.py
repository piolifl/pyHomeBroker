from pyhomebroker import HomeBroker     
import xlwings as xw                    
import pandas as pd                  
from datetime import date, timedelta
import time
import os
import environ
import requests
import yfinance as yf

env = environ.Env()
environ.Env.read_env()
wb = xw.Book('..\\epgb.xlsb')
shtData = wb.sheets('HOME')
shtTickers = wb.sheets('Tickers')

shtData.range('A1').value = 'symbol'
shtData.range('Q1').value = 'PRC'
shtData.range('R1').value = 'ADR'
shtData.range('S1').value = 'D'
shtData.range('T1').value = 'ROLL'
shtData.range('W1').value = 'STOP'
shtData.range('X1').value = 'SCP'
shtData.range('Y1').value = os.environ.get('name')
shtData.range('Z1').value = 0.5

rangoDesde = '28'
rangoHasta = '60'

hoyEs = time.strftime("%A")

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
    #hb.online.subscribe_securities('cedears', '24hs')      # CEDEARS - 24hs
    #hb.online.subscribe_securities('cedears', 'SPOT')      # CEDEARS - spot
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

def diaLaboral():
    if hoyEs == 'Saturday' or hoyEs == 'Sunday':
        return 'Fin de semana'

def getPortfolio(hb, comitente, tipo):
    try:
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
        
        elif os.environ.get('name') == 'VETA':
            portfolio = requests.post("https://cuentas.vetacapital.com.ar/Consultas/GetConsulta", cookies=hb.auth.cookies, json=payload).json()
        
        else: 
            portfolio = requests.post("https://clientes.bcch.org.ar/Consultas/GetConsulta", cookies=hb.auth.cookies, json=payload).json()

        shtData.range('U26:V'+str(rangoHasta)).value = ''
        try: 
            shtData.range('M1').value = portfolio['Result']['Activos'][0]['Subtotal'][0]['APERTURA'][1]['ACUM']
            print('ARS:', portfolio['Result']['Activos'][0]['Subtotal'][0]['APERTURA'][1]['ACUM'], end=' || ')
        except: shtData.range('M1').value = 0
        try: 
            shtData.range('O1').value = portfolio['Result']['Activos'][0]['Subtotal'][2]['APERTURA'][1]['ACUM']
            print('USD MEP:', portfolio['Result']['Activos'][0]['Subtotal'][2]['APERTURA'][1]['ACUM'], ' || ',time.strftime("%H:%M:%S") )
        except: shtData.range('O1').value = 0

        subtotal = [ i['Subtotal'] for i in portfolio["Result"]["Activos"][0:] ]

        for i in subtotal[0:]:
            if i[0]['NERE'] != 'Pesos':  
                subtotal = [ ( x['NERE'],x['CAN0'],x['CANT']) for x in i[0:] if x['CANT'] != None]
                for x in subtotal:
                    for valor in shtData.range('A26:P'+str(rangoHasta)).value:
                        if not valor[0]: continue
                        ticker = str(valor[0]).split()
                        if ticker[0][-1:] == 'D' or ticker[0][-1:] == 'C':  
                            if x[0] == ticker[0][:-1]: 
                                shtData.range('U'+str(int(valor[15]+1))).value = int(x[2])
                        else:
                            if x[0] == ticker[0]: 
                                
                                shtData.range('U'+str(int(valor[15]+1))).value = int(x[2])
                                hayW = shtData.range('W'+str(int(valor[15]+1))).value

                                if tipo == 1:

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
                                            else: shtData.range('W'+str(int(valor[15]+1))).value = valor[5] / 100
    except: pass
#--------------------------------------------------------------------------------------------------------------------------------
if diaLaboral():
    print('Es FIN DE SEMANA, sin logueo y no se actualizan los precios en la planilla.')
    esFinde = True
else: 
    esFinde = False
    login()
    getTodos()
    print(time.strftime("%H:%M:%S"),f"Logueo correcto en: {os.environ.get('name')} cuenta: {int(os.environ.get('account_id'))}")

shtData.range('Y1').value = os.environ.get('name')

#--------------------------------------------------------------------------------------------------------------------------------

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
    cargoXplazo(tikers,monedaInicial)

def traerADR():
    #valorAdr = yf.download(['GGAL'],period='1d',interval='1d',auto_adjust=False)['Close'].values
    valorAdr = yf.download(['GGAL'],period='1d',interval='1d',auto_adjust=False)['Close'].values
    shtData.range('Z61').value = valorAdr[0][0]
    '''max = yf.download(['GGAL'],period='1d',interval='1d',auto_adjust=False)['High'].values
    min = yf.download(['GGAL'],period='1d',interval='1d',auto_adjust=False)['Low'].values
    shtData.range('AB61').value = max[0][0]
    shtData.range('AB62').value = min[0][0]'''
    shtData.range('Y62').value = time.strftime("%H:%M:%S")





def cancelaCompra(celda):
    orderC = shtData.range('AB'+str(int(celda+1))).value
    if not orderC or orderC == None or orderC == 'None' or orderC == '': orderC = 0
    if esFinde == False: 
        try: 
            hb.orders.cancel_order(int(os.environ.get('account_id')),int(orderC))
            print(f"/// Cancelada Compra : {int(orderC)} ",end='\t')
        except: pass
    try: shtData.range('V'+str(int(celda+1))).value -= shtData.range('AC'+str(int(celda+1))).value
    except: pass
    shtData.range('AB'+str(int(celda+1))+':'+'AD'+str(int(celda+1))).value = ''
        
def cancelarVenta(celda):
    orderV = shtData.range('AE'+str(int(celda+1))).value
    if not orderV or orderV == None or orderV == 'None' or orderV == '': orderV = 0
    if esFinde == False: 
        try:
            hb.orders.cancel_order(int(os.environ.get('account_id')),int(orderV))
            print(f"/// Cancelada Venta  : {int(orderV)} " ,end='\t')
        except: pass
    try: shtData.range('V'+str(int(celda+1))).value += shtData.range('AF'+str(int(celda+1))).value
    except: pass
    shtData.range('AE'+str(int(celda+1))+':'+'AG'+str(int(celda+1))).value = ''

def cancelarTodo(desde,hasta):
    if esFinde == False:
        try:  
            hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
            print("/// Todas las ordenes activas canceladas ")
        except: pass
    shtData.range('AB'+str(desde)+':'+'AH'+str(hasta)).value = ''

def cantidadAuto(nroCelda):
    cantidad = shtData.range('Y'+str(int(nroCelda))).value
    if not cantidad or cantidad == None or cantidad == 'None': 
        cantidad = 0
    return abs(int(cantidad))

def soloContinua():
    pass

def stokDisponible(nroCelda):
    stok = shtData.range('U'+str(int(nroCelda))).value
    if not stok or stok == None or stok == 'None': 
        stok = 0
    return abs(int(stok))


###############################################################  ENVIAR ORDENES ################################################    
def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    global orderC, orderV
    orderC, orderV = hoyEs,hoyEs
    symbol = str(shtData.range(str(symbol)).value).split()
    precio = shtData.range(str(price)).value
    if tipo.lower() == 'buy': 
        try: 
            if len(symbol) < 2:
                if esFinde == False: orderC = hb.orders.send_buy_order(symbol[0],'24hs', float(precio), abs(int(size)))
                shtData.range('AD'+str(int(celda+1))).value = float(precio)
                shtData.range('X'+str(int(celda+1))).value = precio
                print(f'        ______/ BUY  opcion + {int(size)} {symbol[0]} // precio: {precio} // {orderC}') 
            else:
                if esFinde == False: orderC = hb.orders.send_buy_order(symbol[0],symbol[2], float(precio), abs(int(size)))
                shtData.range('AD'+str(int(celda+1))).value = float(precio/100)
                shtData.range('X'+str(int(celda+1))).value = precio / 100
                print(f'        ______/ BUY + {int(size)} {symbol[0]} {symbol[2]} // precio: {round(precio/100,4)} // {orderC}')
        except: 
            shtData.range('Q'+str(int(celda+1))+':'+'R'+str(int(celda+1))).value = ''
            print(f'        ______/ ERROR en COMPRA. {symbol[0]} // precio: {precio} // + {int(size)}')

        shtData.range('Q'+str(int(celda+1))+':'+'R'+str(int(celda+1))).value = ''
        try: shtData.range('V'+str(int(celda+1))).value += abs(int(size))
        except: shtData.range('V'+str(int(celda+1))).value = abs(int(size))
        shtData.range('AB'+str(int(celda+1))).value = orderC
        shtData.range('AC'+str(int(celda+1))).value = abs(int(size))
    
    else: # VENTA
        try:
            if len(symbol) < 2:
                if esFinde == False: orderV = hb.orders.send_sell_order(symbol[0],'24hs', float(precio), abs(int(size)))
                shtData.range('AG'+str(int(celda+1))).value = float(precio)
                shtData.range('X'+str(int(celda+1))).value = precio
                print(f'______/ SELL opcion - {int(size)} {symbol[0]} // precio: {precio} // {orderV}')
            else:
                if esFinde == False: orderV = hb.orders.send_sell_order(symbol[0],symbol[2], float(precio), abs(int(size)))
                shtData.range('AG'+str(int(celda+1))).value = float(precio/100)
                shtData.range('X'+str(int(celda+1))).value = precio /100
                print(f'______/ SELL - {int(size)} {symbol[0]} {symbol[2]} // precio: {round(precio/100,4)} // {orderV}')
        except:
            shtData.range('S'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            print(f'______/ ERROR en VENTA. {symbol[0]} // precio: {precio} // {int(size)}')

        shtData.range('S'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
        try: shtData.range('V'+str(int(celda+1))).value -= abs(int(size))
        except: shtData.range('V'+str(int(celda+1))).value = int(size) / -1
        shtData.range('AE'+str(int(celda+1))).value = orderV
        shtData.range('AF'+str(int(celda+1))).value = abs(int(size))
############################################################### TRAILING STOP #################################################
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
            if vendido == 'no':
                if bid > abs(costo) * (1 + (ganancia*75)):
                    shtData.range('X'+str(int(nroCelda+1))).value = bid
                    print(f'TRAILING {nombre[0]} siguiente precio {bid * (1+(ganancia*75))}', time.strftime("%H:%M:%S"))
                    if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'BUYTRAIL': pass
                    else: shtData.range('W'+str(int(nroCelda+1))).value = 'BUYTRAIL'
                    
                if not shtData.range('X1').value:
                    if last < abs(costo) * (1 - (ganancia*75)): 
                        if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'STOP':
                            if bid > last * (1-(ganancia*45)):
                                if shtData.range('Y'+str(int(nroCelda+1))).value : 
                                    try: shtData.range('U'+str(int(nroCelda+1))).value -= abs(cantidad)
                                    except: pass
                                    enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),abs(cantidad),nroCelda)
                        else:
                            if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                            else: 
                                print(f'STOP {nombre[0]} x {cantidad} precio {bid} target salida {costo * (1-(ganancia*75))}', time.strftime("%H:%M:%S"))
                                shtData.range('W'+str(int(nroCelda+1))).value = 'STOP'
            else:
                if ask < abs(costo) * (1 - (ganancia*75)):
                    shtData.range('X'+str(int(nroCelda+1))).value = ask
                    print(f'TRAILING {nombre[0]} siguiente precio {costo * (1-(ganancia*75))}', time.strftime("%H:%M:%S"))
                    if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'SELLTRAIL': pass
                    else: shtData.range('W'+str(int(nroCelda+1))).value = 'SELLTRAIL'
                    
                if not shtData.range('X1').value:  
                    if last > abs(costo) * (1 + (ganancia*75)): 
                        if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'STOP': 
                            
                            if ask < last * (1-(ganancia*15)):
                                if shtData.range('Y'+str(int(nroCelda+1))).value : 
                                    try: shtData.range('U'+str(int(nroCelda+1))).value += abs(cantidad)
                                    except: pass
                                    enviarOrden('buy','A'+str((int(nroCelda)+1)),'D'+str((int(nroCelda)+1)),abs(cantidad),nroCelda)
                        else:
                            if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                            else: 
                                print(f'STOP {nombre[0]} target salida {costo * (1-(ganancia*75))}', time.strftime("%H:%M:%S"))
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
                    print(f'TRAILING {nombre[0]} precio objetivo {bid * (1+(ganancia))}', time.strftime("%H:%M:%S"))     
                    if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'BUYTRAIL': pass
                    else: shtData.range('W'+str(int(nroCelda+1))).value = 'BUYTRAIL'
                
                if not shtData.range('X1').value:
                    if last < abs(costo) * (1 - ganancia*3):
                        if str(shtData.range('W'+str(int(nroCelda+1))).value)=='STOP' and (bid)>(last)*(1-ganancia*1.5):
                            if shtData.range('Y'+str(int(nroCelda+1))).value : 
                                tengoStok = stokDisponible(nroCelda+1)
                                if tengoStok < 1: soloContinua()
                                elif cantidad > tengoStok: cantidad = tengoStok
                                print(f'{time.strftime("%H:%M:%S")} STOP venta    ',end=' || ')
                                shtData.range('U'+str(int(nroCelda+1))).value -= abs(cantidad)
                                shtData.range('W'+str(int(nroCelda+1))).value = ''
                                enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),abs(cantidad),nroCelda)
                        else: 
                            print(f'STOP {nombre[0]} target salida {costo * (1-ganancia*5)}', time.strftime("%H:%M:%S"))
                            if str(shtData.range('W'+str(int(nroCelda+1))).value) == 'STOP': soloContinua()
                            else: 
                                shtData.range('W'+str(int(nroCelda+1))).value = 'STOP'
    except: soloContinua()
############################################################## BUSCA OPERACIONES ###############################################
def buscoOperaciones(inicio,fin):
    for valor in shtData.range('P'+str(inicio)+':'+'U'+str(fin)).value:
        try:
            if not shtData.range('W1').value: # Permite TRAILING  ///////////////////////////////////////////////////////////////
                if not valor[5]:  pass
                else: 
                    if valor[5] < 0: vendido = 'si'
                    else: vendido = 'no'
                    trailingStop('A'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),int(valor[0]),vendido)
        except: pass

        if valor[1]: # # Columna Q en el excel /////////////////////////////////////////////////////////////////////////////////
            if str(valor[1]).lower() == 'c': cancelaCompra(valor[0])
            elif str(valor[1]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[1] == '+': 
                enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
            elif str(valor[1]).upper() == 'P': 
                if esFinde == False: getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: 
                    if shtData.range('AB'+str(int(valor[0]+1))).value: cancelaCompra(valor[0]) # CANCELA oreden compra anterior
                    enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0]) # Compra Bid
                except: shtData.range('Q'+str(int(valor[0]+1))).value = ''
            shtData.range('Q'+str(int(valor[0]+1))).value = ''

        if valor[2]: #  Columna R en el excel //////////////////////////////////////////////////////////////////////////////////
            if str(valor[2]).lower() == 'c': cancelaCompra(valor[0])
            elif str(valor[2]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[2] == '+': 
                enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
            elif str(valor[2]).upper() == 'P': 
                if esFinde == False: getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: 
                    if shtData.range('AB'+str(int(valor[0]+1))).value: cancelaCompra(valor[0])
                    enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0]) # Compra Ask
                except: shtData.range('R'+str(int(valor[0]+1))).value = ''
            shtData.range('R'+str(int(valor[0]+1))).value = ''

        if valor[3]: # Columna S en el excel ///////////////////////////////////////////////////////////////////////////////////
            if str(valor[3]).lower() == 'v': cancelarVenta(valor[0])
            elif str(valor[3]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[3] == '-': 
                enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
            elif str(valor[3]).upper() == 'P': 
                if esFinde == False: getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: 
                    if shtData.range('AE'+str(int(valor[0]+1))).value: cancelarVenta(valor[0])
                    enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0]) # Vendo Bid
                except: shtData.range('S'+str(int(valor[0]+1))).value = ''
            shtData.range('S'+str(int(valor[0]+1))).value = ''

        if valor[4]: # Columna T en el excel //////////////////////////////////////////////////////////////////////////////////
            if str(valor[4]).lower() == 'v': cancelarVenta(valor[0])
            elif str(valor[4]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[4] == '-': 
                enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
            elif str(valor[4]).upper() == 'P': 
                if esFinde == False: getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: 
                    if shtData.range('AE'+str(int(valor[0]+1))).value: cancelarVenta(valor[0]) # CANCELA oreden venta anterior
                    enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0]) # Vendo Ask
                except: shtData.range('T'+str(int(valor[0]+1))).value = ''
            shtData.range('T'+str(int(valor[0]+1))).value = ''
############################################################ CARGA BUCLE EN EXCEL ##############################################
broker = str(shtData.range('Y1').value).upper()

vuelta = 0
vueltaPortfolio = 0

while True:

    if time.strftime("%H:%M:%S") > '17:01:00': 
        if time.strftime("%H:%M:%S") > '17:05:00': pass
        else: break
    
    try:
        preparar =  shtData.range('A1').value
        if preparar != 'symbol': 
            shtData.range('T1').value = 'ROLL'
            preparaRulo(preparar)
    except:
        print('error al preparar los Rulos ')
        shtData.range('A1').value = 'symbol'


    if broker == 'BCCH': 
        buscoOperaciones(rangoDesde,rangoHasta)
    
    time.sleep(2)

    try: 
        if not shtData.range('Q1').value and esFinde == False:
            shtData.range('A31').options(index=True,header=False).value=options
            shtData.range('A'+str(listLength+1)).options(index=True,header=False).value = everything
            try: shtData.range('AJ2').options(index=True, header=False).value = cauciones
            except: print("______ ERROR al cargar cauciones en Excel ______ ",time.strftime("%H:%M:%S")) 
    except: print("______ ERROR al cargar Bonos/Letras en Excel ______ ",time.strftime("%H:%M:%S")) 
 
    try:
        if vuelta > 10: 
            valorAdr = traerADR()
            shtData.range('Z61').value = valorAdr
            vuelta = 0
        else: vuelta += 1
    except: print('ERROR, al cargar el ADR desde yahoo finance')
    
try: 
    hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
    hb.online.disconnect()
except: pass

print(time.strftime("%H:%M:%S"), 'Mercado cerrado. ')

shtData.range('A1').value = 'symbol'
shtData.range('Q1').value = 'PRC'
shtData.range('R1').value = 'ADR'
shtData.range('S1').value = 'D'
shtData.range('T1').value = 'ROLL'
shtData.range('W1').value = 'STOP'
shtData.range('X1').value = 'SCP'
shtData.range('Y1').value = os.environ.get('name')
shtData.range('Z1').value = 0.5

#[ ]><   \n
