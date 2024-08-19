import gspread
import os
import pandas as pd
import pyRofex
import time
import datetime as dt
from pprint import pprint
import requests

today = dt.date.today()
today_year = today.year
today_month = today.month
today_day = today.day
today_yearmonthday = today.strftime('%Y%m%d')

holidays = requests.get(f'http://nolaborables.com.ar/api/v2/feriados/{today_year}').json()
holidays_dict = {key: [] for key in range(1, 13)}

for holiday in holidays:
    holiday_month = holiday['mes']
    holiday_day = holiday['dia']
    holidays_dict[holiday_month].append(holiday_day)

if today_day in holidays_dict[today_month]:
    quit()

end_time = dt.time(17, 0, 30, 0)  # Configurar con el horario de cierre de rueda, ojo que es horario de la pc en la que corre

Registro = gspread.service_account(filename=os.path.dirname(os.path.abspath(__file__)) + '/gspread/[XXXXX].json')  # Dirigir a las credenciales correspondientes del spreadsheet. La hoja se tiene que llamar "Registro" (o modificar los nombres en el código)
Panel = Registro.open("Registro").worksheet("Panel")  # El spreadsheet tiene que tener una hoja "Panel", donde van a ir los precios
Tickers = Registro.open("Registro").worksheet("Tickers")  # El spreadsheet tiene que tener una hoja "tickers" con una columna "A" donde van todos los tickers a 48hs y CI, y una columna "B" donde van los tickers a 24hs.
print('Google spreadsheets connected')

Panel.batch_clear(["3:3002"])

tickers = Tickers.col_values(1)
tickers24 = Tickers.col_values(2)

pyRofex._set_environment_parameter("url", "https://api.eco.xoms.com.ar/", pyRofex.Environment.LIVE)  # Configurar con la URL que corresponda al broker
pyRofex._set_environment_parameter("ws", "wss://api.eco.xoms.com.ar/", pyRofex.Environment.LIVE) # Configurar con el WS que corresponda al broker

try:
    pyRofex.initialize(user="20263866623", password="M8Khq6gQ_", account="62226", environment=pyRofex.Environment.LIVE)

    # pyRofex.initialize(user=[XXXXX],
    #                    password=[XXXXX],
    #                    account=[XXXXX],
    #                    environment=pyRofex.Environment.REMARKET)

# Completar credenciales de acceso según corresponda, se recomienda usar archivo aparte para guardar credenciales

except pyRofex.components.exceptions.ApiException:
    print(f'\npyRofex environment could not be initialized. Authentication failed.'
          '\nCheck login credentials: Incorrect User or Password. (APIException)')
    quit()
print(f'\npyRofex environment successfully initialized')

instruments = pyRofex.get_detailed_instruments()['instruments']

letras_cficode = "DYXTXR"
call_cficode = 'OCASPS'
put_cficode = 'OPASPS'

letras = list()
opciones = list()
vencimientos = list()

for instrument in instruments:
    if instrument['cficode'] == letras_cficode:
        append = instrument['securityDescription'].split(' - ')[2]
        letras.append(append)
    if instrument['cficode'] == call_cficode or instrument['cficode'] == put_cficode:
        if "Galicia" in instrument['underlying']:
            vencimiento = instrument['maturityDate']
            vencimientos.append(vencimiento)

letras = list(dict.fromkeys(letras))
letras.sort()

vencimientos = list(dict.fromkeys(vencimientos))
vencimientos.sort()
vencimientos[:] = [x for x in vencimientos if x >= today_yearmonthday]

for instrument in instruments:
    if instrument['cficode'] == call_cficode or instrument['cficode'] == put_cficode:
        if "Galicia" in instrument['underlying'] or "Comercial" in instrument['underlying'] or "YPF Merval" in instrument['underlying']:
            if instrument['maturityDate'] == vencimientos[0] or instrument['maturityDate'] == vencimientos[1]:
                append = instrument['securityDescription'].split(' - ')[2]
                opciones.append(append)

# Acá arriba elegir de qué activos se quiere tener las opciones.

opciones = list(dict.fromkeys(opciones))
opciones.sort()

tickers = tickers + letras
tickers24 = tickers24 + opciones

instruments_formatted = []
for ticker in tickers:
    instruments_formatted.append('MERV - XMEV - ' + ticker + ' - 48hs')
    instruments_formatted.append('MERV - XMEV - ' + ticker + ' - CI')
for ticker in tickers24:
    instruments_formatted.append('MERV - XMEV - ' + ticker + ' - 24hs')

instruments_raw = pyRofex.get_all_instruments()['instruments']
all_instruments = list()

for instrument_dict in instruments_raw:
    if instrument_dict['instrumentId']['symbol'].split(' - ')[0] == 'MERV':
        all_instruments.append(instrument_dict['instrumentId']['symbol'])

instruments_to_be_removed = list()
symbol_not_found = 0
for instrument in instruments_formatted:
    if instrument not in all_instruments:
        print(f"Instrument {instrument} is not in the API's instrument list")
        instruments_to_be_removed.append(instrument)
        symbol_not_found = 1

if symbol_not_found == 1:
    for remove_this_instrument in instruments_to_be_removed:
        instruments_formatted.remove(remove_this_instrument)
else:
    print(f"\nAll instruments to be subscribed are in the API's instrument list\n")

index_list = [item.replace('MERV - XMEV - ', '') for item in instruments_formatted]
index_list = ['Ticker - Plazo'] + index_list
index_list = [[el] for el in index_list]

Panel.update('A2', index_list)

prices = pd.DataFrame(columns=["Bid_size", "Bid", "Ask", "Ask_size", "Last", "Last_size", 'Nominal_volume',
                               'Effective_volume'], index=instruments_formatted).fillna(0)
prices.index.name = "Instrumento"


msg_date_time = ''

def market_data_handler(message):
    global prices, msg_date_time

    # print(f"Market data received for {message['instrumentId']['symbol'].replace('MERV - XMEV - ', '')} at "
    #       f"{datetime.fromtimestamp(message['timestamp']/1000)}")

    msg_datetime = dt.datetime.fromtimestamp(message['timestamp'] / 1000)
    msg_date_time = msg_datetime.strftime("%m/%d/%Y %H:%M:%S")
    msg_time_time = msg_datetime.time()

    if message['marketData']['LA']:
        prices.loc[message['instrumentId']['symbol'], 'Last'] = message['marketData']['LA']['price']
        prices.loc[message['instrumentId']['symbol'], 'Last_size'] = message['marketData']['LA']['size']
    else:
        prices.loc[message['instrumentId']['symbol'], 'Last'] = 0
        prices.loc[message['instrumentId']['symbol'], 'Last_size'] = 0

    if message['marketData']['OF']:
        prices.loc[message['instrumentId']['symbol'], 'Ask'] = message['marketData']['OF'][0]['price']
        prices.loc[message['instrumentId']['symbol'], 'Ask_size'] = message['marketData']['OF'][0]['size']
    else:
        prices.loc[message['instrumentId']['symbol'], 'Ask'] = 0
        prices.loc[message['instrumentId']['symbol'], 'Ask_size'] = 0

    if message['marketData']['BI']:
        prices.loc[message['instrumentId']['symbol'], 'Bid'] = message['marketData']['BI'][0]['price']
        prices.loc[message['instrumentId']['symbol'], 'Bid_size'] = message['marketData']['BI'][0]['size']
    else:
        prices.loc[message['instrumentId']['symbol'], 'Bid'] = 0
        prices.loc[message['instrumentId']['symbol'], 'Bid_size'] = 0

    if message['marketData']['NV']:
        prices.loc[message['instrumentId']['symbol'], 'Nominal_volume'] = message['marketData']['NV']
    else:
        prices.loc[message['instrumentId']['symbol'], 'Nominal_volume'] = 0

    if message['marketData']['EV']:
        prices.loc[message['instrumentId']['symbol'], 'Effective_volume'] = message['marketData']['EV']
    else:
        prices.loc[message['instrumentId']['symbol'], 'Effective_volume'] = 0


def error_handler(message):
    print(f"\n>>>>>>Error message received at {dt.datetime.now()}:")
    pprint(message)
    pyRofex.close_websocket_connection()
    quit()


def exception_handler(message):
    print(f"\n>>>>>>Exception occurred at {dt.datetime.now()}:")
    pprint(message)
    pyRofex.close_websocket_connection()
    quit()

pyRofex.init_websocket_connection(market_data_handler=market_data_handler,
                                  error_handler=error_handler,
                                  exception_handler=exception_handler)

entries = [pyRofex.MarketDataEntry.BIDS, pyRofex.MarketDataEntry.OFFERS, pyRofex.MarketDataEntry.LAST,
           pyRofex.MarketDataEntry.NOMINAL_VOLUME, pyRofex.MarketDataEntry.TRADE_EFFECTIVE_VOLUME]

half_list = round(len(instruments_formatted)/2)

for x in [instruments_formatted[:half_list], instruments_formatted[half_list:]]:
    pyRofex.market_data_subscription(
        tickers=x,
        entries=entries
    )

print('Websocket connection successfully initialized for:')
pprint(index_list)

while True:

    if dt.datetime.fromtimestamp(time.time()).time() > end_time:
        pyRofex.close_websocket_connection()
        quit()

    try:
        Panel.update('D1', msg_date_time)
        Panel.update('B2', [prices.columns.tolist()] + prices.values.tolist())
    except:
        pass
    time.sleep(1)