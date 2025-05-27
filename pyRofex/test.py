import yfinance as yf


valorAdr = yf.download(['GGAL'],period='1d',interval='1d',auto_adjust=False)['Close'].values

max = yf.download(['GGAL'],period='1d',interval='1d',auto_adjust=False)['High'].values

min = yf.download(['GGAL'],period='1d',interval='1d',auto_adjust=False)['Low'].values

print(max,valorAdr,min)
