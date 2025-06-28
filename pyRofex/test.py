import pyRofex

market_data_recibida = []
reporte_de_ordenes = []

#pyRofex.initialize(user="20263866623",password="hbAned0609*",account="47352",environment=pyRofex.Environment.REMARKET)

pyRofex._set_environment_parameter("url", "https://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)
pyRofex._set_environment_parameter("ws", "wss://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)

def obtenerSaldoCuenta(cuenta=None):
  resumenCuenta = pyRofex.get_account_position(account=cuenta)
  return resumenCuenta

print(obtenerSaldoCuenta('47352'))

{'status': 'OK', 'positions': [{'instrument': {'symbolReference': 'GFGC65452G', 'settlType': 1}, 'symbol': 'MERV - XMEV - GFGC65452G - 24hs', 'buySize': 2.0, 'buyPrice': 442.488, 'sellSize': 2.0, 'sellPrice': 459.713, 'totalDailyDiff': 3445.0, 'totalDiff': 3013.3, 'tradingSymbol': 'MERV - XMEV - GFGC65452G - 24hs', 'originalBuyPrice': 442.488, 'originalSellPrice': 457.5545, 'originalBuySize': 0, 'originalSellSize': 200.0}, {'instrument': {'symbolReference': 'AL30', 'settlType': 0}, 'symbol': 'MERV - XMEV - AL30 - 24hs', 'buySize': 917.0, 'buyPrice': 817.3, 'sellSize': 25.0, 'sellPrice': 814.7, 'totalDailyDiff': -4168.2, 'totalDiff': -1218.27, 'tradingSymbol': 'MERV - XMEV - AL30 - CI', 'originalBuyPrice': 814.08306522, 'originalSellPrice': 814.7, 'originalBuySize': 917.0, 'originalSellSize': 0}, {'instrument': {'symbolReference': 'GFGC63452G', 'settlType': 1}, 'symbol': 'MERV - XMEV - GFGC63452G - 24hs', 'buySize': 3.0, 'buyPrice': 532.069, 'sellSize': 2.0, 'sellPrice': 485.0, 'totalDailyDiff': -11610.7, 'totalDiff': -13366.81, 'tradingSymbol': 'MERV - XMEV - GFGC63452G - 24hs', 'originalBuyPrice': 537.9227, 'originalSellPrice': 485.0, 'originalBuySize': 200.0, 'originalSellSize': 0}, {'instrument': {'symbolReference': 'AL30', 'settlType': 2}, 'symbol': 'MERV - XMEV - AL30 - 24hs', 'buySize': 83.0, 'buyPrice': 817.3, 'sellSize': 0.0, 'sellPrice': 0.0, 'totalDailyDiff': -381.8, 'totalDiff': -521.84, 'tradingSymbol': 'MERV - XMEV - AL30 - 48hs', 'originalBuyPrice': 818.9872167, 'originalSellPrice': 0.0, 'originalBuySize': 83.0, 'originalSellSize': 0}]}