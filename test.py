   
import xlwings as xw                    
import pandas as pd                     
from datetime import date, timedelta
import time

wb = xw.Book('D:\pyHomeBroker\epgb_pyHB.xlsx')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')

def rulo():
    celda,pesos,dolar = 46,100,10000
    tikers = {
        'cclCI':['tiker',dolar],'ccl48':['tiker',dolar],
        'mepCI':['tiker',dolar],'mep48':['tiker',dolar],
        'arsCI':['tiker',pesos],'ars48':['tiker',pesos]
        }
    for valor in shtTest.range('A46:A141').value:
        ars = shtTest.range('AA'+str(celda)).value
        ccl = shtTest.range('Z'+str(celda)).value
        mep = shtTest.range('Z'+str(celda)).value
        if ars != None and ccl != None and mep != None and ars > 0 and ccl > 0 and mep > 0:
            if (valor[7:8] == 's' or valor[8:9] == 's'):
                if ars > tikers['arsCI'][1]: tikers['arsCI'] = [valor[:4]+' - spot',int(ars)]
                if valor[3:4] == 'C' or valor[4:5] == 'C': 
                    if ccl < tikers['cclCI'][1]: tikers['cclCI'] = [valor,int(ccl)]
                if valor[3:4] == 'D' or valor[4:5] == 'D':
                    if mep < tikers['mepCI'][1]: tikers['mepCI'] = [valor,int(mep)]
            if (valor[7:9]=='48' or valor[8:10]=='48'):
                if ars > tikers['ars48'][1]: tikers['ars48'] = [valor[:4]+' - spot',int(ars)]
                if valor[3:4] == 'C' or valor[4:5] == 'C': 
                    if ccl < tikers['ccl48'][1]: tikers['ccl48'] = [valor,int(ccl)]
                if valor[3:4] == 'D' or valor[4:5] == 'D': 
                    if mep < tikers['mep48'][1]: tikers['mep48'] = [valor,int(mep)]
        celda +=1

    # Carga de tikers en planilla excel
    shtTest.range('A2').value = tikers['mepCI'][0]                  
    shtTest.range('A3').value = tikers['cclCI'][0][:4]+'D - spot'   
    shtTest.range('A4').value = tikers['cclCI'][0]                   
    shtTest.range('A5').value = tikers['mepCI'][0][:4]+'C - spot'
    shtTest.range('A6').value = tikers['mepCI'][0]                  
    shtTest.range('A7').value = tikers['arsCI'][0][:4]+'D - spot'   
    shtTest.range('A8').value = tikers['arsCI'][0]                   
    shtTest.range('A9').value = tikers['mepCI'][0][:4]+' - spot'
    shtTest.range('A10').value = tikers['cclCI'][0]                  
    shtTest.range('A11').value = tikers['arsCI'][0][:4]+'C - spot'   
    shtTest.range('A12').value = tikers['arsCI'][0]                   
    shtTest.range('A13').value = tikers['cclCI'][0][:4]+' - spot'

    shtTest.range('A14').value = tikers['mep48'][0]                  
    shtTest.range('A15').value = tikers['ccl48'][0][:4]+'D - 48hs'   
    shtTest.range('A16').value = tikers['ccl48'][0]                   
    shtTest.range('A17').value = tikers['mep48'][0][:4]+'C - 48hs'
    shtTest.range('A18').value = tikers['mep48'][0]                  
    shtTest.range('A19').value = tikers['ars48'][0][:4]+'D - 48hs'   
    shtTest.range('A20').value = tikers['ars48'][0]                   
    shtTest.range('A21').value = tikers['mep48'][0][:4]+' - 48hs'
    shtTest.range('A22').value = tikers['ccl48'][0]                  
    shtTest.range('A23').value = tikers['ars48'][0][:4]+'C - 48hs'   
    shtTest.range('A24').value = tikers['ars48'][0]                   
    shtTest.range('A25').value = tikers['ccl48'][0][:4]+' - 48hs'
    
rulo()







