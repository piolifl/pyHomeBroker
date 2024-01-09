   
import xlwings as xw                    
import pandas as pd                     
from datetime import date, timedelta
import time

wb = xw.Book('D:\pyHomeBroker\epgb_pyHB.xlsx')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')

def nameARS(name):
    if  name  == 'XE4D' : name = 'X18E4'
    elif name == 'XE4C' : name = 'X18E4'
    elif name == 'SE4D': name = 'S18E4'
    elif name == 'SE4C': name = 'S18E4'
    elif name == 'MRCA': name = 'MRCAO'
    elif name == 'MRCAD': name = 'MRCAO'
    elif name == 'MRCAC': name = 'MRCAO'
    elif name == 'CLSI': name = 'CLSIO'
    elif name == 'CLSID': name = 'CLSIO'
    elif name == 'CLSIC': name = 'CLSIO'
    elif name == 'BA7D': name = 'BA37D'
    elif name == 'BA7DD': name = 'BA37D'
    elif name == 'BA7DC': name = 'BA37D'
    elif name == 'AL29D': name = 'AL29'
    elif name == 'AL29C': name = 'AL29'
    elif name == 'AL30D': name = 'AL30'
    elif name == 'AL30C': name = 'AL30'
    elif name == 'AE38D': name = 'AE38'
    elif name == 'AE38C': name = 'AE38'
    elif name == 'AL35D': name = 'AL35'
    elif name == 'AL35C': name = 'AL35'
    elif name == 'AL41D': name = 'AL41'
    elif name == 'AL41C': name = 'AL41'
    elif name == 'GD29D': name = 'GD29'
    elif name == 'GD29C': name = 'GD29'
    elif name == 'GD35D': name = 'GD35'
    elif name == 'GD35C': name = 'GD35'
    elif name == 'GD38D': name = 'GD38'
    elif name == 'GD38C': name = 'GD38'
    elif name == 'GD41D': name = 'GD41'
    elif name == 'GD41C': name = 'GD41'
    elif name == 'GD46D': name = 'GD46'
    elif name == 'GD46C': name = 'GD46'
    return name

def nameMEP(name):
    if  name  == 'X18E4':name = 'XE4D'
    elif name  == 'XE4C':name = 'XE4D'
    elif name == 'S18E4': name= 'SE4D'
    elif name == 'SE4D': name= 'SE4C'
    elif name == 'MRCAO': name = 'MRCAD'
    elif name == 'MRCA': name = 'MRCAD'
    elif name == 'CLSI': name= 'CLSID'
    elif name == 'CLSIO': name= 'CLSID'
    elif name == 'BA37D': name= 'BA7DD'
    elif name == 'BA7D': name= 'BA7DD'
    elif name == 'AL29': name = 'AL29D'
    elif name == 'AL30': name = 'AL30D'
    elif name == 'AE38': name = 'AE38D'
    elif name == 'AL35': name = 'AL35D'
    elif name == 'AL41': name = 'AL41D'
    elif name == 'GD29': name = 'GD29D'
    elif name == 'GD35': name = 'GD35D'
    elif name == 'GD38': name = 'GD38D'
    elif name == 'GD41': name = 'GD41D'
    elif name == 'GD46': name = 'GD46D'
    return name

def nameCCL(name):
    if  name  == 'X18E4':name = 'XE4C'
    elif name  == 'XE4D':name = 'XE4C'
    elif name == 'S18E4': name= 'SE4C'
    elif name == 'SE4D': name= 'SE4C'
    elif name == 'MRCAO': name = 'MRCAC'
    elif name == 'MRCA': name = 'MRCAC'
    elif name == 'CLSI': name= 'CLSIC'
    elif name == 'CLSIO': name= 'CLSIC'
    elif name == 'BA37D': name= 'BA7DC'
    elif name == 'AL29': name = 'AL29C'
    elif name == 'AL30': name = 'AL30C'
    elif name == 'AE38': name = 'AE38C'
    elif name == 'AL35': name = 'AL35C'
    elif name == 'AL41': name = 'AL41C'
    elif name == 'GD29': name = 'GD29C'
    elif name == 'GD35': name = 'GD35C'
    elif name == 'GD38': name = 'GD38C'
    elif name == 'GD41': name = 'GD41C'
    elif name == 'GD46': name = 'GD46C'
    return name


def rulo():
    print(time.strftime("%H:%M:%S"),"Buscando mejores precios ...",end=" // ")
    celda,pesos,dolar = 46,100,10000
    tikers = {
        'cclCI':['tiker',dolar],'ccl48':['tiker',dolar],
        'mepCI':['tiker',dolar],'mep48':['tiker',dolar],
        'arsCImep':['tiker',pesos],'ars48mep':['tiker',pesos],
        'arsCIccl':['tiker',pesos],'ars48ccl':['tiker',pesos]
        }

    for valor in shtTest.range('A46:A147').value:
        arsM = shtTest.range('AA'+str(celda)).value
        arsC = shtTest.range('AA'+str(celda)).value
        ccl = shtTest.range('Z'+str(celda)).value
        mep = shtTest.range('Z'+str(celda)).value
        if arsM != None and arsC != None and ccl != None and mep != None:
            if (valor[7:8] == 's' or valor[8:9] == 's'):
                if valor[3:4] == 'C' or valor[4:5] == 'C': 
                    if arsC > tikers['arsCIccl'][1]: tikers['arsCIccl'] = [valor[:4],arsC]
                    if ccl < tikers['cclCI'][1]: tikers['cclCI'] = [valor,ccl]
                if valor[3:4] == 'D' or valor[4:5] == 'D':
                    if arsM > tikers['arsCImep'][1]: tikers['arsCImep'] = [valor[:4],arsM]
                    if mep < tikers['mepCI'][1]: tikers['mepCI'] = [valor,mep]
            if (valor[7:9]=='48' or valor[8:10]=='48'):
                if valor[3:4] == 'C' or valor[4:5] == 'C': 
                    if arsC > tikers['ars48ccl'][1]: tikers['ars48ccl'] = [valor[:4],arsC]
                    if ccl < tikers['ccl48'][1]: tikers['ccl48'] = [valor,ccl]
                if valor[3:4] == 'D' or valor[4:5] == 'D': 
                    if arsM > tikers['ars48mep'][1]: tikers['ars48mep'] = [valor[:4],arsM]
                    if mep < tikers['mep48'][1]: tikers['mep48'] = [valor,mep]
        celda +=1
    
    # Carga de tikers en planilla excel
    shtTest.range('A22').value = tikers['mepCI'][0]     
    shtTest.range('Y22').value = tikers['mepCI'][1]
    shtTest.range('Z22').value = nameARS(tikers['mepCI'][0][:4])+' - spot'
    shtTest.range('AA22').value = nameCCL(tikers['mepCI'][0][:4])+' - spot'
    shtTest.range('A23').value = tikers['mep48'][0]
    shtTest.range('Y23').value = tikers['mep48'][1]
    shtTest.range('Z23').value = nameARS(tikers['mep48'][0][:4])+' - 48hs'
    shtTest.range('AA23').value = nameCCL(tikers['mep48'][0][:4])+' - 48hs'
    shtTest.range('A24').value = tikers['cclCI'][0]
    shtTest.range('Y24').value = tikers['cclCI'][1]
    shtTest.range('Z24').value = nameARS(tikers['cclCI'][0][:4])+' - spot'
    shtTest.range('AA24').value = nameMEP(tikers['cclCI'][0][:4])+' - spot'
    shtTest.range('A25').value = tikers['ccl48'][0]
    shtTest.range('Y25').value = tikers['ccl48'][1]
    shtTest.range('Z25').value = nameARS(tikers['ccl48'][0][:4])+' - 48hs'
    shtTest.range('AA25').value = nameMEP(tikers['ccl48'][0][:4])+' - 48hs'

    shtTest.range('A26').value = tikers['arsCImep'][0]+' - spot'
    shtTest.range('Y26').value = tikers['arsCImep'][1]
    shtTest.range('Z26').value = nameMEP(tikers['arsCImep'][0])+' - spot'
    shtTest.range('AA26').value = nameCCL(tikers['arsCImep'][0])+' - spot'
    shtTest.range('A27').value = tikers['ars48mep'][0]+' - 48hs'
    shtTest.range('Y27').value = tikers['ars48mep'][1]
    shtTest.range('Z27').value = nameMEP(tikers['ars48mep'][0])+' - 48hs'
    shtTest.range('AA27').value = nameCCL(tikers['ars48mep'][0])+' - 48hs'
    shtTest.range('A28').value = tikers['arsCIccl'][0]+' - spot'
    shtTest.range('Y28').value = tikers['arsCIccl'][1]
    shtTest.range('Z28').value = nameCCL(tikers['arsCIccl'][0])+' - spot'
    shtTest.range('AA28').value = nameMEP(tikers['arsCIccl'][0])+' - spot'
    shtTest.range('A29').value = tikers['ars48ccl'][0]+' - 48hs'
    shtTest.range('Y29').value = tikers['ars48ccl'][1]
    shtTest.range('Z29').value = nameCCL(tikers['ars48ccl'][0])+' - 48hs'
    shtTest.range('AA29').value = nameMEP(tikers['ars48ccl'][0])+' - 48hs'
    print(time.strftime("%H:%M:%S"),'Done!')
    print(tikers)
rulo()







