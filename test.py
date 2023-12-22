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

def getBonosList():
    rng = shtTickers.range('E2:E115').expand()
    oBonos = rng.value
    Bonos = pd.DataFrame({'symbol' : oBonos}, columns=["symbol", "bid_size", "bid", "ask", "ask_size", 
                                                       "last", "change", "open", "high", "low", 
                                                       "previous_close", "turnover", "volume", 
                                                       'operations', 'datetime'])
    Bonos = Bonos.set_index('symbol')
    Bonos['datetime'] = pd.to_datetime(Bonos['datetime'])
    return Bonos

bonos = getBonosList()

def on_securities(online, quotes):
    global ACC
    thisData = quotes
    thisData = thisData.reset_index()
    thisData['symbol'] = thisData['symbol'] + ' - ' +  thisData['settlement']
    thisData = thisData.drop(["settlement"], axis=1)
    thisData = thisData.set_index("symbol")
    thisData['change'] = thisData["change"] / 100
    thisData['datetime'] = pd.to_datetime(thisData['datetime'])
    bonos.update(thisData)

hb = HomeBroker(int(os.environ.get('broker')),
                on_securities=on_securities)

hb.auth.login(dni=str(os.environ.get('dni')),
              user=str(os.environ.get('user')),
              password=str(os.environ.get('password')),
              raise_exception=True)

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
    # hb.online.subscribe_securities('short_term_government_bonds', '48hs')   # LETRAS - 48hs
    # hb.online.subscribe_securities('short_term_government_bonds', '24hs')  # LETRAS - 24hs
    # hb.online.subscribe_securities('short_term_government_bonds', 'SPOT')   # LETRAS - spot
    # hb.online.subscribe_securities('corporate_bonds', '48hs')  # Obligaciones Negociables - 48hs
    # hb.online.subscribe_securities('corporate_bonds', '24hs')  # Obligaciones Negociables - 24hs
    # hb.online.subscribe_securities('corporate_bonds', 'SPOT')  # Obligaciones Negociables - spot
hb.online.subscribe_repos()


print(bonos)

ccl = [ tiker for tiker in bonos if tiker[0][4:5]=='C']
mep = [ tiker for tiker in bonos if tiker[0][4:5]=='D']

print(mep)
