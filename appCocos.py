from pycocos import Cocos


# Bordame1+

app = Cocos(email="miryam.borda79@gmail.com", password="Bordame1+")


# Get the available portfolio with the current market valuation
#print(app.my_portfolio())

# Get the available funds
print(app.funds_available())
'''
{'total': {'ars': 116967.9535, 'usd': 90.05881391949549}, 'tickers': [{'id_instrument': None, 'id_security': None, 'short_ticker': None, 'ticker': 'AR$', 'type': None, 'subtype': None, 'instrument_name': None, 'instrument_short_name': None, 'instrument_code': None, 'quantity': 116967.9535, 'amount': 116967.9535, 'amount_usd': 90.05881391949549, 'last': 1, 'variation': None, 'logo_file_name': 'ars.jpg'}, {'id_instrument': None, 'id_security': None, 'short_ticker': None, 'ticker': 'US$ MEP', 'type': None, 'subtype': None, 'instrument_name': None, 'instrument_short_name': None, 'instrument_code': None, 'quantity': 0, 'amount': 0, 'amount_usd': 0, 'last': 1298.79518072, 'variation': None, 'logo_file_name': 'usd.jpg'}]}
{'CI': {'ars': 116967.9535, 'usd': 0, 'ext': 0}, '24hs': {'ars': 116967.9535, 'usd': 0, 'ext': 0}, '48hs': {'ars': 116967.9535, 'usd': 0, 'ext': 0}}
'''


'''long_ticker = app.long_ticker(ticker="GFGV29581O", 
                              settlement=app.settlements.T1, 
                              currency=app.currencies.PESOS, 
                              segment=app.segments.OPTIONS)

print(long_ticker)'''

precios0 = app.get_instrument_snapshot(ticker="AL30", segment=app.segments.DEFAULT)
print('AL30-0001-C-CT-ARS',precios0)

'''
AL30-0001-C-CT-ARS [
{'short_ticker': 'AL30D', 'long_ticker': 'AL30D-0002-C-CT-USD', 'instrument_code': 'AL30', 
'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2030', 
'instrument_short_name': 'Argentina 2030', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 
'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 
'settlement_days': 1, 'currency': 'USD', 'price_factor': 100, 'contract_size': 1, 
'min_lot_size': 1, 'id_security': 3286, 'tick_size': 0.001, 
'date': '2024-08-15', 'open': 49.75, 'high': 50.29, 'low': 49.629, 
'close': 50.29, 'prev_close': 49.85, 'last': 50.29, 'bid': 49.6, 'ask': 50.5, 
'bids': [{'size': 28, 'price': 49.6}, {'size': 89, 'price': 47.21}, 
{'size': 100, 'price': 46.5}, {'size': 91, 'price': 45.5}, {'size': 29, 'price': 45}], 
'asks': [{'size': 610, 'price': 50.5}, {'size': 3981, 'price': 50.6}, 
{'size': 771, 'price': 50.67}, {'size': 623, 'price': 50.75}, {'size': 3822, 'price': 51}], 
'turnover': 31662123.3, 'volume': 63372141, 'variation': 0.00882648, 'term': '24hs', 
'id_tick_size_rule': 'BYMA_FIXED_INCOME', 'is_favorite': True, 'newTerm': 1}, 
{'short_ticker': 'AL30', 'long_ticker': 
'AL30-0001-C-CT-ARS', 'instrument_code': 'AL30', 'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2030', 'instrument_short_name': 'Argentina 2030', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 0, 'currency': 'ARS', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 3258, 'tick_size': 10, 'date': '2024-08-15', 'open': 63240, 'high': 64290, 'low': 62700, 'close': 64290, 'prev_close': 62880, 'last': 64290, 'bid': 62900, 'ask': 65000, 'bids': [{'size': 707, 'price': 62900}, {'size': 157, 'price': 62800}, {'size': 49, 'price': 62500}, {'size': 2, 'price': 62000}, {'size': 50, 'price': 57850}], 'asks': [{'size': 30000, 'price': 65000}, {'size': 921, 'price': 67700}], 'turnover': 203362015292.1023, 'volume': 320284290, 'variation': 0.02242366, 'term': 'CI', 'id_tick_size_rule': 'BYMA_FIXED_INCOME', 'is_favorite': True, 'newTerm': 0}, {'short_ticker': 'AL30D', 'long_ticker': 'AL30D-0001-C-CT-USD', 'instrument_code': 'AL30', 'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2030', 'instrument_short_name': 'Argentina 2030', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 0, 'currency': 'USD', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 3285, 'tick_size': 0.01, 'date': '2024-08-15', 'open': 49.67, 'high': 50.3, 'low': 49.611, 'close': 50.22, 'prev_close': 49.73, 'last': 50.22, 'bid': 46.9, 'ask': 50.5, 'bids': [{'size': 500, 'price': 46.9}, {'size': 200, 'price': 46.5}, {'size': 500, 'price': 46.3}, {'size': 200, 'price': 45}], 'asks': [{'size': 1788, 'price': 50.5}, {'size': 1, 'price': 51}, {'size': 5107, 'price': 51.5}, {'size': 300, 'price': 52}, {'size': 9142, 'price': 53}], 'turnover': 142057226.51, 'volume': 284706831, 'variation': 0.00985321, 'term': 'CI', 'id_tick_size_rule': 'BYMA_FIXED_INCOME', 'is_favorite': True, 'newTerm': 0}, {'short_ticker': 'AL30', 'long_ticker': 'AL30-0002-C-CT-ARS', 'instrument_code': 'AL30', 'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2030', 'instrument_short_name': 'Argentina 2030', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 1, 'currency': 'ARS', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 3259, 'tick_size': 10, 'date': '2024-08-15', 'open': 63300, 'high': 64190, 'low': 63030, 'close': 64080, 'prev_close': 63250, 'last': 64080, 'bid': 63900, 'ask': 65000, 'bids': [{'size': 1711, 'price': 63900}, {'size': 243, 'price': 63400}, {'size': 10, 'price': 63000}, {'size': 9590, 'price': 62400}, {'size': 55, 'price': 62300}], 'asks': [{'size': 30000, 'price': 65000}, {'size': 200, 'price': 65900}, {'size': 1914, 'price': 67500}, {'size': 942, 'price': 70000}, {'size': 5000, 'price': 70900}], 'turnover': 99202277592.0001, 'volume': 155987883, 'variation': 0.01312253, 'term': '24hs', 'id_tick_size_rule': 'BYMA_FIXED_INCOME', 'is_favorite': True, 'newTerm': 1}]
'''

print()
print()
precios1 = app.get_instrument_snapshot(ticker="GFGV29581O", segment=app.segments.DEFAULT)
print('GFGV29581O-0002-O-CT-ARS',precios1)





'''
# Send a withdrawal order of 1000 pesos
app.withdraw_funds(currency=app.currencies.PESOS, 
                   amount="1000", 
                   cbu_cvu="0000003100070922163640")

# Get the long ticker for AL30 with T+2 settlement
long_ticker = app.long_ticker(ticker="AL30", 
                              settlement=app.settlements.T0, 
                              currency=app.currencies.PESOS)

# Send a buy order for 200 AL30 bonds with T+2 settlement at $9000. By default, all orders are *LIMIT* orders.
order = app.submit_buy_order(long_ticker=long_ticker, 
                             quantity="200", 
                             price="9000")

# Cancel an order by order_id
app.cancel_order(order_number=order['Orden'])

# Get the quoteboard for "Acciones panel Lideres", T+2 settlement, traded in Pesos
app.instrument_list_snapshot(instrument_type=app.instrument_types.ACCIONES, 
                             instrument_subtype=app.instrument_subtypes.LIDERES, 
                             settlement=app.settlements.T2, 
                             currency=app.currencies.PESOS, 
                             segment=app.segments.DEFAULT)


'''