   
import xlwings as xw                    
import pandas as pd                     
from datetime import date, timedelta
import time

wb = xw.Book('D:\pyHomeBroker\epgb_pyHB.xlsx')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')

def cclArs(celdaInicial,dolar,pesos):
    for valor in shtTest.range('A46:A141').value:
        ccl = shtTest.range('Z'+str(celdaInicial)).value
        ars = shtTest.range('AA'+str(celdaInicial)).value
        if valor[:2] != 'BA' and (valor[3:4] == 'C' or valor[4:5] == 'C') and ccl != None and ars != None:
            if valor[8:9] == 's' : 
                if ccl < dolar:
                    dolar = ccl
                    shtTest.range('A10').value = valor
                    shtTest.range('A13').value = valor[:4]+' - spot'
                    shtTest.range('A20').value = valor
                    shtTest.range('A19').value = valor[:4]+'D - spot'
                if ars > pesos:
                    pesos = ars
                    shtTest.range('A11').value = valor
                    shtTest.range('A12').value = valor[:4]+' - spot'
            if valor[8:10] == '48':
                if ccl < dolar:
                    dolar = ccl
                    shtTest.range('A14').value = valor
                    shtTest.range('A17').value = valor[:4]+' - 48hs'
                    shtTest.range('A24').value = valor
                    shtTest.range('A23').value = valor[:4]+'D - 48hs'
                if ars > pesos:
                    pesos = ars
                    shtTest.range('A15').value = valor
                    shtTest.range('A16').value = valor[:4]+' - 48hs'
        celdaInicial +=1


def mepArs(celdaInicial,dolar,pesos):
    for valor in shtTest.range('A46:A141').value:
        mep = shtTest.range('Z'+str(celdaInicial)).value
        ars = shtTest.range('AA'+str(celdaInicial)).value
        if valor[:2] != 'BA' and (valor[3:4] == 'D' or valor[4:5] == 'D') and mep != None and ars != None :
            if valor[8:9] == 's': # operacion en corto
                if mep < dolar:
                    dolar = mep
                    shtTest.range('A2').value = valor
                    shtTest.range('A5').value = valor[:4]+' - spot'
                    shtTest.range('A18').value = valor
                    shtTest.range('A21').value = valor[:4]+'C - spot'
                if ars > pesos:
                    pesos = ars
                    shtTest.range('A3').value = valor
                    shtTest.range('A4').value = valor[:4]+' - spot'
            if valor[8:10] == '48': # operacin en largo
                if mep < dolar:
                    dolar = mep
                    shtTest.range('A6').value = valor
                    shtTest.range('A9').value = valor[:4]+' - 48hs'
                    shtTest.range('A22').value = valor
                    shtTest.range('A25').value = valor[:4]+'C - 48hs'
                if ars > pesos:
                    pesos = ars
                    shtTest.range('A7').value = valor
                    shtTest.range('A8').value = valor[:4]+' - 48hs'
        celdaInicial +=1



mepArs(46,10000,1)
cclArs(46,10000,1)


