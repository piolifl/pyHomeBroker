import yfinance as yf


#yf.download('GGAL',period='1d',interval='1d')['Close'].value

galiciaADR= yf.download('GGAL',period='1d',interval='1d',prepost=True)['Close'].values
Low = yf.download('GGAL',period='1d',interval='1d',prepost=True)['Low'].values
High = yf.download('GGAL',period='1d',interval='1d',prepost=True)['Close'].values
    

print(galiciaADR)   
print(Low) 
print(High) 

    
'''   
    galiciaADRMin= yf.download('GGAL',period='1d',interval='1d')['Low'].values
    galiciaADRMax= yf.download('GGAL',period='1d',interval='1d')['High'].values
    realbr= yf.download('USDBRL=X', period='1d', interval='1d')['Close'].values
    spy= yf.download('SPY',period='1d',interval='1d')['Close'].values
    shtTest.range('Q2').value = galiciaADR
    shtTest.range('Q4').value = galiciaADRMin
    shtTest.range('Q5').value = galiciaADRMax
    shtTest.range('Q3').value = realbr
    shtTest.range('Q6').value = spy

'''