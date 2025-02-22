import yfinance as yf

def adr():
    galiciaADR= yf.download(['GGAL','YPF'],period='1d',interval='1d',auto_adjust=False)['Close'].values
    return 'GGAL'+ galiciaADR[0][0] , galiciaADR[0][1]


adr()