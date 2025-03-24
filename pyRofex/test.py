import xlwings as xw

wb = xw.Book('D:\\pyHomeBroker\\epgb.xlsb')
shtTickers = wb.sheets('pyRofex')
shtData = wb.sheets('MATRIZ OMS')

arsMEP = ['',0, '',0]
arsCCL = ['',0, '',0]

celda = 64

nombre = shtData.range('A'+str(int(celda+1))).value
ticker = str(nombre).split()

arsCIbid = shtData.range('C'+str(int(celda+1))).value
arsCIask = shtData.range('D'+str(int(celda+1))).value
cclCIbid = shtData.range('C'+str(int(celda+3))).value
cclCIask = shtData.range('D'+str(int(celda+3))).value
mepCIbid = shtData.range('C'+str(int(celda+5))).value
mepCIask = shtData.range('D'+str(int(celda+5))).value

mep = round(arsCIask/mepCIbid)
ccl = round(arsCIask/cclCIbid)
arsMEP[0] = nombre
arsMEP[1] = mep
arsCCL[0] = nombre
arsCCL[1] = ccl

nombre = shtData.range('A'+str(int(celda+2))).value
ticker = str(nombre).split()

ars24bid = shtData.range('C'+str(int(celda+2))).value
ars24ask = shtData.range('D'+str(int(celda+2))).value
ccl24bid = shtData.range('C'+str(int(celda+4))).value
ccl24ask = shtData.range('D'+str(int(celda+4))).value
mep24bid = shtData.range('C'+str(int(celda+6))).value
mep24ask = shtData.range('D'+str(int(celda+6))).value

mep = round(ars24ask/mep24bid)
ccl = round(ars24ask/ccl24bid)
arsMEP[2] = nombre
arsMEP[3] = mep
arsCCL[2] = nombre
arsCCL[3] = ccl

def excelRulo(celda):

    for valor in shtData.range('A65:A148').value:
        if not valor: continue
        ticker = str(valor).split()

        if ticker[2] == 'CI':
            moneda = ticker[0][-2:]

            try: 
                if moneda == '7D' or moneda == 'M5' or moneda == 'J5' or int(moneda) >= 0 : # COMPRA MEP y CCL mas barato con ARS

                    arsCIask = shtData.range('D'+str(int(celda+1))).value
                    mepCIbid = shtData.range('C'+str(int(celda+5))).value
                    cclCIbid = shtData.range('C'+str(int(celda+3))).value

                    mep = round(arsCIask/mepCIbid)
                    ccl = round(arsCIask/cclCIbid)

                    if mep < arsMEP[2]:  
                        arsMEP[0] = valor
                        arsMEP[1] = mep
                        
                    if ccl < arsCCL[2]:  
                        arsCCL[0] = valor
                        arsCCL[1] = ccl
                        
            except: pass
        celda += 1

excelRulo(64)
print(arsMEP,arsCCL)












        