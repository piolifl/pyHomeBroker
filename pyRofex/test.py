import yfinance as yf

tickers = ["GGAL"]

df = yf.download(tickers, interval='5m', period='1d', prepost=False, progress=False).iloc[-1:].Close

print(df)



