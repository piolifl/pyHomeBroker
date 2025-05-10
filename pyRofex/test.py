import yfinance as yf

from curl_cffi import requests
session = requests.Session(impersonate="chrome")
ticker = yf.Ticker('GGAL', session=session)

#valorAdr = yf.download(['^GGAL'],period='1d',interval='1d',auto_adjust=False)['Close'].values



print(ticker)