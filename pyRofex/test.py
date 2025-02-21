import yfinance as yf

galiciaADR= yf.download('GGAL',period='1d',interval='1d')['Close'].values

print(galiciaADR)