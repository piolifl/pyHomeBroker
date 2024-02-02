   
import xlwings as xw                    
import pandas as pd                     
from datetime import date, timedelta
import time

wb = xw.Book('D:\pyHomeBroker\epgb_pyHB.xlsx')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')
shtTest.range('Q2:X29').value  = 0


celda,pesos,dolar = 46,1000,0
tikers = {
        'cclCI':['',dolar],'ccl48':['',dolar],
        'mepCI':['',dolar],'mep48':['',dolar],
        'arsCIccl':['',pesos],'ars48ccl':['',pesos],
        'arsCImep':['',pesos],'ars48mep':['',pesos]
        }

def namesArs(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA37D'+plazo
    elif nombre[:2] == 'BP': return 'BPO27'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') and (nombre[3:4] == 'D' or nombre[3:4] == 'C'):
        if (nombre[1:2] == 'F' or nombre[1:2] == 'Y'): return nombre[:1]+'20'+nombre[1:3]+plazo
        else: return nombre[:1]+'18'+nombre[1:3]+plazo
    elif (nombre[:2] == 'AL' or nombre[:2] == 'GD' or nombre[:2] == 'AE') and (nombre[4:5] == 'D' or nombre[4:5] == 'C'):
        return nombre[:4]+plazo
    else: return nombre[:4]+'O'+plazo

def namesCcl(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA7DC'+plazo
    elif nombre[:2] == 'BP': return 'BP27C'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') :
        if nombre[3:4] == 'D': return nombre[:3]+'C'+plazo
        else: return nombre[:1]+nombre[3:5]+'C'+plazo
    elif (nombre[:2] == 'AL' or nombre[:2] == 'GD' or nombre[:2] == 'AE') and (nombre[4:5] == 'D' or nombre[4:5] == ' '):
        return nombre[:4]+'C'+plazo
    else: return nombre[:4]+'C'+plazo

def namesMep(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA7DD'+plazo
    elif nombre[:2] == 'BP': return 'BP27D'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') :
        if nombre[3:4] == 'C': return nombre[:3]+'D'+plazo
        else: return nombre[:1]+nombre[3:5]+'D'+plazo
    elif (nombre[:2] == 'AL' or nombre[:2] == 'GD' or nombre[:2] == 'AE') and (nombre[4:5] == 'D' or nombre[4:5] == ' '):
        return nombre[:4]+'D'+plazo
    else: return nombre[:4]+'D'+plazo

def cargoPlanilla(dicc):
    if dicc['mepCI'][1] != 10000:
        shtTest.range('A22').value = dicc['mepCI'][0]    
        shtTest.range('Y22').value = dicc['mepCI'][1]
        shtTest.range('Z22').value = namesArs(dicc['mepCI'][0],' - spot')
        shtTest.range('AA22').value =namesCcl(dicc['mepCI'][0],' - spot')
    if dicc['mep48'][1] != 10000:    
        shtTest.range('A23').value = dicc['mep48'][0]
        shtTest.range('Y23').value = dicc['mep48'][1]
        shtTest.range('Z23').value = namesArs(dicc['mep48'][0],' - 48hs')
        shtTest.range('AA23').value =namesCcl(dicc['mep48'][0],' - 48hs')
    if dicc['cclCI'][1] != 10000:
        shtTest.range('A24').value = dicc['cclCI'][0]
        shtTest.range('Y24').value = dicc['cclCI'][1]
        shtTest.range('Z24').value = namesArs(dicc['cclCI'][0],' - spot')
        shtTest.range('AA24').value =namesMep(dicc['cclCI'][0],' - spot')
    if dicc['ccl48'][1] != 10000:
        shtTest.range('A25').value = dicc['ccl48'][0]
        shtTest.range('Y25').value = dicc['ccl48'][1]
        shtTest.range('Z25').value = namesArs(dicc['ccl48'][0],' - 48hs')
        shtTest.range('AA25').value =namesMep(dicc['ccl48'][0],' - 48hs')

    if dicc['arsCImep'][1] != 100:
        shtTest.range('A26').value = dicc['arsCImep'][0]
        shtTest.range('Y26').value = dicc['arsCImep'][1]
        shtTest.range('Z26').value = namesMep(dicc['arsCImep'][0],' - spot')
        shtTest.range('AA26').value =namesCcl(dicc['arsCImep'][0],' - spot')
    if dicc['ars48mep'][1] != 100:
        shtTest.range('A27').value = dicc['ars48mep'][0]
        shtTest.range('Y27').value = dicc['ars48mep'][1]
        shtTest.range('Z27').value = namesMep(dicc['ars48mep'][0],' - 48hs')
        shtTest.range('AA27').value =namesCcl(dicc['ars48mep'][0],' - 48hs')
    if dicc['arsCIccl'][1] != 100:
        shtTest.range('A28').value = dicc['arsCIccl'][0]
        shtTest.range('Y28').value = dicc['arsCIccl'][1]
        shtTest.range('Z28').value = namesCcl(dicc['arsCIccl'][0],' - spot')
        shtTest.range('AA28').value =namesMep(dicc['arsCIccl'][0],' - spot')
    if dicc['ars48ccl'][1] != 100:
        shtTest.range('A29').value = dicc['ars48ccl'][0]
        shtTest.range('Y29').value = dicc['ars48ccl'][1]
        shtTest.range('Z29').value = namesCcl(dicc['ars48ccl'][0],' - 48hs')
        shtTest.range('AA29').value =namesMep(dicc['ars48ccl'][0],' - 48hs') 

def cargoXplazo(dicc):
    if time.strftime("%H:%M:%S") > '16:26:00':
        shtTest.range('A10').value = dicc['mep48'][0] # mep
        shtTest.range('A11').value = namesMep(dicc['ars48mep'][0],' - 48hs') #  mep
        shtTest.range('A12').value = dicc['ars48mep'][0] #  ars
        shtTest.range('A13').value = namesArs(dicc['mep48'][0],' - 48hs') # ars
        
        shtTest.range('A14').value = dicc['mep48'][0] # mep

        shtTest.range('A15').value = namesMep(dicc['ccl48'][0],' - 48hs')
        shtTest.range('A16').value = dicc['ccl48'][0] # ccl

        shtTest.range('A17').value = namesCcl(dicc['mep48'][0],' - 48hs')
    else:
        shtTest.range('A10').value = dicc['mepCI'][0]
        shtTest.range('A11').value = namesMep(dicc['arsCImep'][0],' - spot')
        shtTest.range('A12').value = dicc['arsCImep'][0]
        shtTest.range('A13').value = namesArs(dicc['mepCI'][0],' - spot')

        shtTest.range('A14').value = dicc['mepCI'][0] # mep

        shtTest.range('A15').value = namesMep(dicc['cclCI'][0],' - spot')
        shtTest.range('A16').value = dicc['cclCI'][0] # ccl

        shtTest.range('A17').value = namesCcl(dicc['mepCI'][0],' - spot')

shtTest.range('A10:A17').value = ''
shtTest.range('A22:A29').value = ''
shtTest.range('Y22:AA29').value = ''

for valor in shtTest.range('A46:A146').value:
        arsM = shtTest.range('AA'+str(celda)).value
        if arsM == None: arsM = 1000
        arsC = arsM
        ccl = shtTest.range('Z'+str(celda)).value
        if ccl == None: ccl = 0
        mep = ccl
        if (valor[7:8] == 's' or valor[8:9] == 's'):
            if valor[3:4] == 'C' or valor[4:5] == 'C': 
                if arsC > tikers['arsCIccl'][1]: tikers['arsCIccl'] = [namesArs(valor[:5],' - spot'),arsC]
                if ccl > tikers['cclCI'][1]: tikers['cclCI'] = [valor,ccl]

            if valor[3:4] == 'D' or valor[4:5] == 'D':
                if arsM > tikers['arsCImep'][1]: tikers['arsCImep'] = [namesArs(valor[:5],' - spot'),arsM]
                if mep > tikers['mepCI'][1]: tikers['mepCI'] = [valor,mep]

        if (valor[7:9]=='48' or valor[8:10]=='48'):
            if valor[3:4] == 'C' or valor[4:5] == 'C': 
                if arsC > tikers['ars48ccl'][1]: tikers['ars48ccl'] = [namesArs(valor[:5],' - 48hs'),arsC]
                if ccl > tikers['ccl48'][1]: tikers['ccl48'] = [valor,ccl]
            if valor[3:4] == 'D' or valor[4:5] == 'D': 
                if arsM > tikers['ars48mep'][1]: tikers['ars48mep'] = [namesArs(valor[:5],' - 48hs'),arsM]
                if mep > tikers['mep48'][1]: tikers['mep48'] = [valor,mep]
        celda +=1

print(tikers)

cargoXplazo(tikers)
cargoPlanilla(tikers)


#[ ]><   \n




