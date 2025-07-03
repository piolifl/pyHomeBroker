import pyRofex
import os
import environ

env = environ.Env()
environ.Env.read_env()

market_data_recibida = []
reporte_de_ordenes = []

#pyRofex.initialize(user="20263866623",password="hbAned0609*",account="47352",environment=pyRofex.Environment.REMARKET)


pyRofex._set_environment_parameter("url", "https://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)
pyRofex._set_environment_parameter("ws", "wss://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)
pyRofex.initialize(
    user=str(os.environ.get('user')), 
    password=str(os.environ.get('password')), 
    account=str(os.environ.get('account')), 
    environment=pyRofex.Environment.LIVE)


def obtenerSaldoMatriz(cuenta=None):
    try:
        resumenCuenta = pyRofex.get_account_report(account=cuenta)
        print('Disponible Matriz para Gtias: ',resumenCuenta)
    except: print('Error obterner disponible Gtias ')


#obtenerSaldoMatriz(str(os.environ.get('account')))

valor = {'status': 'OK', 'accountData': {'accountName': '47352', 'collateral': 0, 'margin': 0, 'availableToCollateral': 19754262.04, 
        'detailedAccountReports': {'0': {'currencyBalance': {'detailedCurrencyBalance': 
        {'USD MtR': {'consumed': 0.0, 'available': 0.0}, 
        'ARS': {'consumed': 0.0, 'available': 303164.71}, 
        'USD UY': {'consumed': 0.0, 'available': 0.0}, 
        'USD DB': {'consumed': 0.0, 'available': 0.0}, 
        'U$S': {'consumed': 0.0, 'available': 0.0}, 
        'USD G': {'consumed': 0.0, 'available': 0.0}, 
        'USD R': {'consumed': 0.0, 'available': 0.0}, 
        'USD C': {'consumed': 0.0, 'available': 0.05}, 
        'USD D': {'consumed': 0.0, 'available': 0.55}}}, 
        'availableToOperate': {'cash': {'totalCash': 303903.0, 
        'detailedCash': 
        {'USD MtR': 0.0, 'ARS': 303164.71, 
         'USD UY': 0.0, 'USD DB': 0.0, 
         'U$S': 0.0, 'USD G': 0.0, 
         'USD R': 0.0, 'USD C': 0.05, 
         'USD D': 0.55}}, 
         'movements': 0.0, 'total': 303903.0, 'pendingMovements': 0.0}, 
         'settlementDate': 1751425200000}, 
         '1': {'currencyBalance': {'detailedCurrencyBalance': 
                                   {'USD MtR': {'consumed': 0.0, 'available': 0.0}, 
                                    'ARS': {'consumed': 26975646.6, 'available': 48327518.11}, 
                                    'USD UY': {'consumed': 0.0, 'available': 0.0}, 
                                    'USD DB': {'consumed': 0.0, 'available': 0.0}, 
                                    'U$S': {'consumed': 0.0, 'available': 0.0}, 
                                    'USD G': {'consumed': 0.0, 'available': 0.0}, 
                                    'USD R': {'consumed': 0.0, 'available': 0.0}, 
                                    'USD C': {'consumed': 0.0, 'available': 0.05}, 
                                    'USD D': {'consumed': 0.0, 'available': 0.55}}}, 
      'availableToOperate': {'cash': {'totalCash': 0.0, 
                                      'detailedCash': {'USD MtR': 0.0, 'ARS': 0.0, 
                                                       'USD UY': 0.0, 'USD DB': 0.0, 
                                                       'U$S': 0.0, 'USD G': 0.0, 
                                                       'USD R': 0.0, 'USD C': 0.0, 
                                                       'USD D': 0.0}}, 
                                                       'movements': -26975646.6, 
                                                       'credit': 75000000.0, 
                                                       'total': 48328256.4, 
                                                       'pendingMovements': 0.0}, 
                                                       'settlementDate': 1751511600000}, 
            '2': {'currencyBalance': {'detailedCurrencyBalance': 
                                      {'USD MtR': {'consumed': 0.0, 'available': 0.0}, 
                                       'ARS': {'consumed': 0.0, 'available': 48327518.11}, 
                                       'USD UY': {'consumed': 0.0, 'available': 0.0}, 
                                       'USD DB': {'consumed': 0.0, 'available': 0.0}, 
                                       'U$S': {'consumed': 0.0, 'available': 0.0}, 
                                       'USD G': {'consumed': 0.0, 'available': 0.0}, 
                                       'USD R': {'consumed': 0.0, 'available': 0.0}, 
                                       'USD C': {'consumed': 0.0, 'available': 0.05}, 
                                       'USD D': {'consumed': 0.0, 'available': 0.55}}}, 
              'availableToOperate': {'cash': {'totalCash': 0.0, 'detailedCash': {'USD MtR': 0.0, 'ARS': 0.0, 'USD UY': 0.0, 'USD DB': 0.0, 'U$S': 0.0, 'USD G': 0.0, 'USD R': 0.0, 'USD C': 0.0, 'USD D': 0.0}}, 'movements': 0.0, 'total': 48328256.4, 'pendingMovements': 0.0}, 'settlementDate': 1751598000000}}, 'hasError': False, 'lastCalculation': 1751470455253, 'portfolio': 19450359.04, 'ordersMargin': 0.0, 'currentCash': 303903.0, 'dailyDiff': 0.0, 'uncoveredMargin': 0.0}}

#print(valor["accountData"]['detailedAccountReports']['1']['availableToOperate']['credit'])


#print(valor["accountData"]['detailedAccountReports']['1']['currencyBalance']['detailedCurrencyBalance']['ARS']['available'])













{'status': 'OK', 'positions': [{'instrument': {'symbolReference': 'GFGC65452G', 'settlType': 1}, 'symbol': 'MERV - XMEV - GFGC65452G - 24hs', 'buySize': 2.0, 'buyPrice': 442.488, 'sellSize': 2.0, 'sellPrice': 459.713, 'totalDailyDiff': 3445.0, 'totalDiff': 3013.3, 'tradingSymbol': 'MERV - XMEV - GFGC65452G - 24hs', 'originalBuyPrice': 442.488, 'originalSellPrice': 457.5545, 'originalBuySize': 0, 'originalSellSize': 200.0}, {'instrument': {'symbolReference': 'AL30', 'settlType': 0}, 'symbol': 'MERV - XMEV - AL30 - 24hs', 'buySize': 917.0, 'buyPrice': 817.3, 'sellSize': 25.0, 'sellPrice': 814.7, 'totalDailyDiff': -4168.2, 'totalDiff': -1218.27, 'tradingSymbol': 'MERV - XMEV - AL30 - CI', 'originalBuyPrice': 814.08306522, 'originalSellPrice': 814.7, 'originalBuySize': 917.0, 'originalSellSize': 0}, {'instrument': {'symbolReference': 'GFGC63452G', 'settlType': 1}, 'symbol': 'MERV - XMEV - GFGC63452G - 24hs', 'buySize': 3.0, 'buyPrice': 532.069, 'sellSize': 2.0, 'sellPrice': 485.0, 'totalDailyDiff': -11610.7, 'totalDiff': -13366.81, 'tradingSymbol': 'MERV - XMEV - GFGC63452G - 24hs', 'originalBuyPrice': 537.9227, 'originalSellPrice': 485.0, 'originalBuySize': 200.0, 'originalSellSize': 0}, {'instrument': {'symbolReference': 'AL30', 'settlType': 2}, 'symbol': 'MERV - XMEV - AL30 - 24hs', 'buySize': 83.0, 'buyPrice': 817.3, 'sellSize': 0.0, 'sellPrice': 0.0, 'totalDailyDiff': -381.8, 'totalDiff': -521.84, 'tradingSymbol': 'MERV - XMEV - AL30 - 48hs', 'originalBuyPrice': 818.9872167, 'originalSellPrice': 0.0, 'originalBuySize': 83.0, 'originalSellSize': 0}]}