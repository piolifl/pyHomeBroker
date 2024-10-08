import pycocos
import time


# Bordame1+

app = pycocos.Cocos("miryam.borda79@gmail.com","Bordame1+")


mep = app.get_dolar_mep_info()
lista_px = app.get_instrument_list_snapshot(instrument_type=app.instrument_types.BONOS, instrument_subtype=app.instrument_subtypes.ARS, settlement=app.settlements.T2, currency=app.currencies.PESOS, segment=app.segments.DEFAULT)

# Get the available portfolio with the current market valuation
#print(app.my_portfolio())

# Get the available funds
while True:
    print()
    print(mep)

    time.sleep(10)





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