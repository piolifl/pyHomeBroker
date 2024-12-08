import datetime
import time
import pyRofex

pyRofex._set_environment_parameter("url", "https://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)
pyRofex._set_environment_parameter("ws", "wss://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)
pyRofex.initialize(user="20263866623", password="Bordame01!", account="47352", environment=pyRofex.Environment.LIVE)
    

#2-Defines the handlers that will process the messages and exceptions.
def market_data_handler(message):
    print("Market Data Message Received: {0}".format(message))


def error_handler(message):
    print("Error Message Received: {0}".format(message))


def exception_handler(e):
    print("Exception Occurred: {0}".format(e.msg))


# 3-Initialize Websocket Connection with the handlers
pyRofex.init_websocket_connection(market_data_handler=market_data_handler,
                                  error_handler=error_handler,
                                  exception_handler=exception_handler)


# 4-Subscribes to receive market data messages
instruments = ["MERV - XMEV - AL30 - CI", "MERV - XMEV - AL30 - 24hs"]  # Instruments list to subscribe

entries = [pyRofex.MarketDataEntry.BIDS,
           pyRofex.MarketDataEntry.OFFERS,
           pyRofex.MarketDataEntry.LAST]

pyRofex.market_data_subscription(tickers=instruments,
                                 entries=entries)

# Subscribes to an Invalid Instrument (Error Message Handler should be call)
pyRofex.market_data_subscription(tickers=["InvalidInstrument"],
                                 entries=entries)


# Wait 5 sec then close the connection
time.sleep(2)
pyRofex.close_websocket_connection()

'''
Market Data Message Received: {'type': 'Md', 'timestamp': 1733169290243, 'instrumentId': {'marketId': 'ROFX', 'symbol': 'MERV - XMEV - AL30 - 24hs'}, 'marketData': {'OF': [{'price': 76830, 'size': 64364}], 'BI': [{'price': 76820, 'size': 2569}], 'LA': {'price': 76830, 'size': 1, 'date': 1733169288000}}}
Market Data Message Received: {'type': 'Md', 'timestamp': 1733169290243, 'instrumentId': {'marketId': 'ROFX', 'symbol': 'MERV - XMEV - AL30 - CI'}, 'marketData': {'OF': [{'price': 76800, 'size': 19576}], 'BI': [{'price': 76790, 'size': 2094}], 'LA': {'price': 76790, 'size': 120283, 'date': 1733167808000}}}
Error Message Received: {'status': 'ERROR', 'message': '{"type":"smd","level":1,"depth":1,"entries":["BI","OF","LA"],"products":[{"symbol":"InvalidInstrument","marketId":"ROFX"}]}', 'description': "Product InvalidInstrument:ROFX don't exist"}
Market Data Message Received: {'type': 'Md', 'timestamp': 1733169290962, 'instrumentId': {'marketId': 'ROFX', 'symbol': 'MERV - XMEV - AL30 - 24hs'}, 'marketData': {'OF': [{'price': 76830, 'size': 64364}], 'BI': [{'price': 76820, 'size': 2569}], 'LA': {'price': 76830, 'size': 2491, 'date': 1733169290000}}}
Market Data Message Received: {'type': 'Md', 'timestamp': 1733169291462, 'instrumentId': {'marketId': 'ROFX', 'symbol': 'MERV - XMEV - AL30 - 24hs'}, 'marketData': {'OF': [{'price': 76830, 'size': 64364}], 'BI': [{'price': 76820, 'size': 2569}], 'LA': {'price': 76830, 'size': 1389, 'date': 1733169290000}}}
Market Data Message Received: {'type': 'Md', 'timestamp': 1733169291961, 'instrumentId': {'marketId': 'ROFX', 'symbol': 'MERV - XMEV - AL30 - 24hs'}, 'marketData': {'OF': [{'price': 76830, 'size': 60484}], 'BI': [{'price': 76820, 'size': 2569}], 'LA': {'price': 76830, 'size': 1389, 'date': 1733169290000}}}
Market Data Message Received: {'type': 'Md', 'timestamp': 1733169293098, 'instrumentId': {'marketId': 'ROFX', 'symbol': 'MERV - XMEV - AL30 - 24hs'}, 'marketData': {'OF': [{'price': 76830, 'size': 60484}], 'BI': [{'price': 76820, 'size': 2569}], 'LA': {'price': 76830, 'size': 2857, 'date': 1733169293000}}}
Market Data Message Received: {'type': 'Md', 'timestamp': 1733169293597, 'instrumentId': {'marketId': 'ROFX', 'symbol': 'MERV - XMEV - AL30 - 24hs'}, 'marketData': {'OF': [{'price': 76830, 'size': 60484}], 'BI': [{'price': 76820, 'size': 2569}], 'LA': {'price': 76830, 'size': 2363, 'date': 1733169293000}}}
Market Data Message Received: {'type': 'Md', 'timestamp': 1733169294097, 'instrumentId': {'marketId': 'ROFX', 'symbol': 'MERV - XMEV - AL30 - 24hs'}, 'marketData': {'OF': [{'price': 76830, 'size': 56363}], 'BI': [{'price': 76820, 'size': 2569}], 'LA': {'price': 76830, 'size': 2363, 'date': 1733169293000}}}
'''
