# Importar la biblioteca pyCocos
import pycocos

# Crear una instancia del cliente y realizar la autenticación
app = pycocos.Cocos('email', 'password')


historico_precios = app.get_daily_history(long_ticker="AL30–0003-C-CT-ARS", date_from="2023–01–01")

# Obtener información y última cotización de un instrumento
precios = app.get_instrument_snapshot(ticker="AL30", segment=app.segments.DEFAULT)

# Obtener una lista de instrumentos
lista_px = app.get_instrument_list_snapshot(instrument_type=app.instrument_types.BONOS, instrument_subtype=app.instrument_subtypes.ARS, settlement=app.settlements.T2, currency=app.currencies.PESOS, segment=app.segments.DEFAULT)

# Obtener la lista de instrumentos recomendados
recomendados = app.get_recommended_tickers()

# Obtener la lista de instrumentos favoritos de la cuenta
favoritos = app.get_favorites_tickers()

# Buscar un instrumento
resultado = app.search_ticker(query="AL3")

# Obtener las reglas de precios para cada grupo de instrumentos
reglas = app.instruments_rules()

# Obtener las combinaciones posibles de tipos y subtipos para los paneles de instrumentos
tipos = app.instrument_types_and_subtypes()

# Obtener el estado de las diferentes formas de liquidación
# 0 = CI, 1 = 24 hs., 2 = 48hs, 3 = BOT dolar mep
status = app.market_status()

# Obtener información sobre los diferentes Dólar MEP
mep = app.get_dolar_mep_info()
open_mep = app.get_open_dolar_mep_info()


# IMPORTANTE: el envio de ordenes a precio de mercado aun no fue desarrollado.

# Envío de una orden de compra con precio limitado
orden = app.submit_buy_order(long_ticker="AL30–0003-C-CT-ARS", quantity="1000", price="14500", order_type=app.order_types.LIMIT)

# Envío de una orden de venta con precio limitado
orden = app.submit_sell_order(long_ticker="AL30–0003-C-CT-ARS", quantity="1000", price="14500", order_type=app.order_types.LIMIT)

# IMPORTANTE: estos metodos validan que la cuenta tenga suficiente 
# saldo antes de enviar la orden.

# Colocar una orden de caución# por ejemplo 50.000 pesos a 7 dias a 85% de tasa anual.
orden = app.place_repo_order(currency=app.currencies.PESOS, amount=50000, term=7, rate=85)

# Al enviar satisfactoriamente una orden, el servidor responde con el nro 
# de orden, es necesario conservarlo para consultar el estado y/o cancelarla

# Cancelar una orden
cancel = app.cancel_order(order_number="12023")

# Consultar el estado de una orden
estado = app.order_status(order_number="12023")

# Consultar el estado de todas las órdenes del día
estado = app.order_status()


# Obtener los datos de la cuenta comitente
data = app.my_data()

# Obtener los datos de las cuentas bancarias asociadas
cuentas = app.my_bank_accounts()

# Obtener información actual del portfolio y saldos en diferentes monedas
portfolio = app.my_portfolio()

# Consultar poder de compra
poder_compra = app.funds_available()

# Consultar la cantidad de valores negociables disponibles
poder_venta = app.stocks_available(long_ticker="AL30-0003-C-CT-ARS")

# Obtener un historial de movimientos de cuenta corriente
movimientos = app.account_activity(date_from="2023-01-01", 
date_to="2023-06-30")

# Consultar el rendimiento de la cartera
performance_diaria = app.portfolio_performance(timeframe=app.performance_timeframes.DAILY)
performance_historica = app.portfolio_performance(timeframe=app.performance_timeframes.HISTORICAL)

# Dar de alta una nueva cuenta bancaria
respuesta = app.submit_new_bank_account(cbu="1234567890123456789", cuit="20123456787", currency=app.currencies.PESOS)

# Solicitar extracción de fondos
respuesta = app.withdraw_funds(currency=app.currencies.PESOS, amount="1000", cbu_cvu="1234567890123456789")





# Cerrar sesión
app.logout()


# Generar long ticker para AL30 en pesos con liquidacion en 48 hs. 
long_ticker = app.long_ticker(ticker="AL30", settlement=app.settlements.T1, currency=app.currencies.PESOS)

# Generar long ticker para GOOGLD (dolares) con liquidacion en CI
long_ticker = app.long_ticker(ticker="GOOGLD", settlement=app.settlements.T0, currency=app.currencies.USD)

# Generar long ticker para opciones de GGAL GFGC500.JU
long_ticker = app.long_ticker(ticker="GFGC500.JU", settlement=app.settlements.T1, currency=app.currencies.PESOS, segment=app.segments.OPTIONS)
