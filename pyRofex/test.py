import yfinance as yf



valorAdr = yf.download(['GGAL'],period='1d',interval='1d',auto_adjust=False)['Close'].values
    #valorAdr = yf.download(['GGAL','YPF'],period='1d',interval='1d',auto_adjust=False)['Close'].values

print(valorAdr[0][0])