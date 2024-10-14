import pycocos

app = pycocos.Cocos("miryam.borda79@gmail.com","Bordame1+")

print('Se loguea correctamente !')


precios = app.get_instrument_snapshot(ticker="AL30", segment=app.segments.DEFAULT)

# Generar long ticker para AL30 en pesos con liquidacion en 24 hs. 
long_ticker = app.long_ticker(
    ticker="AL30", 
    settlement=app.settlements.T1, 
    currency=app.currencies.PESOS)

# Generar long ticker para GOOGLD (dolares) con liquidacion en CI
long_ticker = app.long_ticker(
    ticker="GOOGLD", 
    settlement=app.settlements.T0, 
    currency=app.currencies.USD)

# Generar long ticker para opciones de GGAL GFGC500.JU
long_ticker_ = app.long_ticker(
    ticker="GFGV46581O", 
    settlement=app.settlements.T1, 
    currency=app.currencies.PESOS, 
    segment=app.segments.OPTIONS) #GFGV46581O-0002-O-CT-ARS


precios = app.get_instrument_snapshot(ticker="GFGV46581O", segment=app.segments.OPTIONS)

print('GFGV46581O ',precios)


app.logout()
