import yfinance as yf
import xlwings as xw 

wb = xw.Book('.\\epgb_pyHB.xlsb')
shtTest = wb.sheets('HomeBroker')

#yf.download('GGAL',period='1d',interval='1d')['Close'].value

galiciaADR= yf.download('GGAL',period='1d',interval='1d').values
    

print(galiciaADR)    

shtTest.range('AA90').value = galiciaADR[0]  
    
    
    
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