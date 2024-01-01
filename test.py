   
import xlwings as xw                    
import pandas as pd                     
from datetime import date, timedelta
import time

wb = xw.Book('D:\pyHomeBroker\epgb_pyHB.xlsx')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')

def nameArs(name):
    if name == 'SE4D': name = 'S18E4'
    elif name == 'SE4C': name = 'S18E4'
    elif name == 'MRCA': name = 'MRCAO'
    elif name == 'CLSI': name = 'CLSIO'
    elif name == 'BA7D': name = 'BA37D'
    elif name == 'AL30C': name = 'AL30'
    elif name == 'AL30D': name = 'AL30'
    elif name == 'AE38C': name = 'AE38'
    elif name == 'AE38D': name = 'AE38'
    elif name == 'AL29C': name = 'AL29'
    elif name == 'AL29D': name = 'AL29'
    elif name == 'AL35C': name = 'AL35'
    elif name == 'AL35D': name = 'AL35'
    elif name == 'AL41C': name = 'AL41'
    elif name == 'AL41D': name = 'AL41'
    elif name == 'GD29C': name = 'GD29'
    elif name == 'GD29D': name = 'GD29'
    elif name == 'GD35C': name = 'GD35'
    elif name == 'GD35D': name = 'GD35'
    elif name == 'GD38C': name = 'GD38'
    elif name == 'GD38D': name = 'GD38'
    elif name == 'GD41C': name = 'GD41'
    elif name == 'GD41D': name = 'GD41'
    elif name == 'GD46C': name = 'GD46'
    elif name == 'GD46D': name = 'GD46'
    return name
def nameCcl(name):
    if name == 'S18E4': name = 'SE4C'
    elif name == 'S18E4': name = 'SE4C'
    elif name == 'MRCAO': name = 'MRCAC'
    elif name == 'CLSIO': name = 'CLSIC'
    elif name == 'BA37D': name = 'BA7DC'
    elif name == 'AL30': name = 'AL30C'
    elif name == 'GD30': name = 'GD30C'
    elif name == 'AE38': name = 'AE38C'
    elif name == 'AL29': name = 'AL29C'
    elif name == 'AL35': name = 'AL35C'
    elif name == 'AL41': name = 'AL41C'
    elif name == 'GD29': name = 'GD29C'
    elif name == 'GD35': name = 'GD35C'
    elif name == 'GD38': name = 'GD38C'
    elif name == 'GD41': name = 'GD41C'
    elif name == 'GD46': name = 'GD46C'
    return name
def nameMep(name):
    if name == 'S18E4': name = 'SE4D'
    elif name == 'S18E4': name = 'SE4D'
    elif name == 'MRCAO': name = 'MRCAD'
    elif name == 'CLSIO': name = 'CLSID'
    elif name == 'BA37D': name = 'BA7DD'
    elif name == 'AL30': name = 'AL30D'
    elif name == 'GD30': name = 'GD30D'
    elif name == 'AE38': name = 'AE38D'
    elif name == 'AL29': name = 'AL29D'
    elif name == 'AL35': name = 'AL35D'
    elif name == 'AL41': name = 'AL41D'
    elif name == 'GD29': name = 'GD29D'
    elif name == 'GD35': name = 'GD35D'
    elif name == 'GD38': name = 'GD38D'
    elif name == 'GD41': name = 'GD41D'
    elif name == 'GD46': name = 'GD46D'
    return name

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

        if ars != None and ccl != None and mep != None:
            if (valor[7:8] == 's' or valor[8:9] == 's'):
                if ars > tikers['arsCI'][1]: tikers['arsCI'] = [nameArs(valor[:4])+' - spot',int(ars)]
                if valor[3:4] == 'C' or valor[4:5] == 'C': 
                    if ccl < tikers['cclCI'][1]: tikers['cclCI'] = [nameCcl(valor[:4])+' - spot',int(ccl)]
                if valor[3:4] == 'D' or valor[4:5] == 'D':
                    if mep < tikers['mepCI'][1]: tikers['mepCI'] = [nameMep(valor[:4])+' - spot',int(mep)]

            if (valor[7:9]=='48' or valor[8:10]=='48'):
                if ars > tikers['ars48'][1]: tikers['ars48'] = [nameArs(valor[:4])+' - 48hs',int(ars)]
                if valor[3:4] == 'C' or valor[4:5] == 'C': 
                    if ccl < tikers['ccl48'][1]: tikers['ccl48'] = [nameCcl(valor[:4])+' - 48hs',int(ccl)]
                if valor[3:4] == 'D' or valor[4:5] == 'D': 
                    if mep < tikers['mep48'][1]: tikers['mep48'] = [nameMep(valor[:4])+' - 48hs',int(mep)]
        celda +=1
    print(tikers)
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







