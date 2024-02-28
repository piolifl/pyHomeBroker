from pyhomebroker import HomeBroker     
import xlwings as xw                    
import pandas as pd                     
from datetime import date, timedelta
import time

wb = xw.Book('..\epgb_pyHB.xlsx')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')
shtTest.range('Q1').value = 'O'
shtTest.range('R1').value = 'B'
shtTest.range('T1').value = 0.025
shtTest.range('U1:V1').value = 0
shtTest.range('S1').value ='N'
shtTest.range('W1').value = 1

#-------------------------------------------------------------------------------------------------------
print(time.strftime("%H:%M:%S"),"Inicia TEST" )
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
        shtTest.range('A18').value = dicc['mepCI'][0]    
        shtTest.range('Y18').value = dicc['mepCI'][1]
        shtTest.range('Z18').value = namesArs(dicc['mepCI'][0],' - spot')
        shtTest.range('AA18').value =namesCcl(dicc['mepCI'][0],' - spot')
    if dicc['mep48'][1] != 10000:    
        shtTest.range('A19').value = dicc['mep48'][0]
        shtTest.range('Y19').value = dicc['mep48'][1]
        shtTest.range('Z19').value = namesArs(dicc['mep48'][0],' - 48hs')
        shtTest.range('AA19').value =namesCcl(dicc['mep48'][0],' - 48hs')
    if dicc['cclCI'][1] != 10000:
        shtTest.range('A20').value = dicc['cclCI'][0]
        shtTest.range('Y20').value = dicc['cclCI'][1]
        shtTest.range('Z20').value = namesArs(dicc['cclCI'][0],' - spot')
        shtTest.range('AA20').value =namesMep(dicc['cclCI'][0],' - spot')
    if dicc['ccl48'][1] != 10000:
        shtTest.range('A21').value = dicc['ccl48'][0]
        shtTest.range('Y21').value = dicc['ccl48'][1]
        shtTest.range('Z21').value = namesArs(dicc['ccl48'][0],' - 48hs')
        shtTest.range('AA21').value =namesMep(dicc['ccl48'][0],' - 48hs')

    if dicc['arsCImep'][1] != 100:
        shtTest.range('A22').value = dicc['arsCImep'][0]
        shtTest.range('Y22').value = dicc['arsCImep'][1]
        shtTest.range('Z22').value = namesMep(dicc['arsCImep'][0],' - spot')
        shtTest.range('AA22').value =namesCcl(dicc['arsCImep'][0],' - spot')
    if dicc['ars48mep'][1] != 100:
        shtTest.range('A23').value = dicc['ars48mep'][0]
        shtTest.range('Y23').value = dicc['ars48mep'][1]
        shtTest.range('Z23').value = namesMep(dicc['ars48mep'][0],' - 48hs')
        shtTest.range('AA23').value =namesCcl(dicc['ars48mep'][0],' - 48hs')
    if dicc['arsCIccl'][1] != 100:
        shtTest.range('A24').value = dicc['arsCIccl'][0]
        shtTest.range('Y24').value = dicc['arsCIccl'][1]
        shtTest.range('Z24').value = namesCcl(dicc['arsCIccl'][0],' - spot')
        shtTest.range('AA24').value =namesMep(dicc['arsCIccl'][0],' - spot')
    if dicc['ars48ccl'][1] != 100:
        shtTest.range('A25').value = dicc['ars48ccl'][0]
        shtTest.range('Y25').value = dicc['ars48ccl'][1]
        shtTest.range('Z25').value = namesCcl(dicc['ars48ccl'][0],' - 48hs')
        shtTest.range('AA25').value =namesMep(dicc['ars48ccl'][0],' - 48hs') 

def limpio():
    shtTest.range('A10:A25').value = ''
    shtTest.range('Y18:AA25').value = ''

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

def ilRulo():
    shtTest.range('A1').value = 'symbol'
    limpio()
    celda,pesos,dolar = 46,1000,0
    tikers = {'cclCI':['',dolar],'ccl48':['',dolar],'mepCI':['',dolar],'mep48':['',dolar],'arsCIccl':['',pesos],'ars48ccl':['',pesos],'arsCImep':['',pesos],'ars48mep':['',pesos]}
    for valor in shtTest.range('A46:A153').value:
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
    cargoXplazo(tikers)
    cargoPlanilla(tikers)
    
def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    global orderC, orderV
    symbol = str(shtTest.range(str(symbol)).value).split()
    mas = float(shtTest.range('U1').value)
    por = int(shtTest.range('W1').value)
    precio = float(shtTest.range(str(price)).value) + mas
    precioV = precio - (mas * 2)
    if shtTest.range('V'+str(int(celda+1))).value == 0: 
        shtTest.range('W'+str(int(celda+1))+':'+'X'+str(int(celda+1))).value = 0
    if tipo.lower() == 'buy': 
        if len(symbol) < 2:
            ##orderC = hb.orders.send_buy_order(symbol[0],'24hs', float(precio),int(size))
            print(f'Buy {symbol[0]} // + {int(size)} // a {precio}')
            shtTest.range('V'+str(int(celda+1))).value += int(size)
            shtTest.range('W'+str(int(celda+1))).value += int(size) * float(precio)*100
        else:
            ##orderC = hb.orders.send_buy_order(symbol[0],symbol[2],float(precio),int(size*por))
            print(f'Buy {symbol[0]} {symbol[2]} // + {int(size*por)} // a {precio/100}')
            shtTest.range('V'+str(int(celda+1))).value += int(size*por)
            shtTest.range('W'+str(int(celda+1))).value += int(size*por) * float(precio)/100
    else: 
        if len(symbol) < 2:
            ##orderV = hb.orders.send_sell_order(symbol[0],'24hs', float(precioV),int(size))
            print(f'Sell {symbol[0]} // - {int(size)} // a {precioV}')
            shtTest.range('V'+str(int(celda+1))).value -= int(size)
            shtTest.range('W'+str(int(celda+1))).value -= int(size) * float(precioV)*100
        else:
            ##orderV = hb.orders.send_sell_order(symbol[0],symbol[2],float(precioV),int(size*por))
            print(f'Sell {symbol[0]} {symbol[2]} // - {int(size*por)} // a {precioV/100}')
            shtTest.range('V'+str(int(celda+1))).value -= int(size*por)
            shtTest.range('W'+str(int(celda+1))).value -= int(size*por) * float(precioV)/100
    shtTest.range('Q'+str(int(valor[0]+1))+':'+'T'+str(int(valor[0]+1))).value = ''
    if int(shtTest.range('V'+str(int(celda+1))).value) == 0: 
        shtTest.range('X'+str(int(celda+1))).value = float(shtTest.range('W'+str(int(celda+1))).value)/-1
    else: shtTest.range('X'+str(int(celda+1))).value = shtTest.range('W'+str(int(celda+1))).value / shtTest.range('V'+str(int(celda+1))).value

####################################### TRAILINGStop STOP ################################################
def trailingStop(nombre=str,cantidad=int,nroCelda=int):
    try:
        nombre = str(shtTest.range(str(nombre)).value).split()
        bid = float(shtTest.range('C'+str(int(nroCelda+1))).value)
        bid_size = int(shtTest.range('B'+str(int(nroCelda+1))).value)
        last = float(shtTest.range('F'+str(int(nroCelda+1))).value)
        costo = float(shtTest.range('X'+str(int(nroCelda+1))).value) 
        ganancia = float(shtTest.range('T1').value)
        if cantidad > bid_size : cantidad = bid_size
        if len(nombre) < 2: #TRAILING sobre opciones financieras
            if bid * 100 > costo * (1 + ganancia): # Precio sube activo trailing y sube % ganancia               
                if str(shtTest.range('W'+str(int(nroCelda+1))).value)=='TRAILING': shtTest.range('T1').value*=1+0.25
                else: shtTest.range('W'+str(int(nroCelda+1))).value = 'TRAILING'
                shtTest.range('X'+str(int(nroCelda+1))).value = bid * 100
            if last * 100 < costo * (1 - ganancia): # Precio baja activo stop y envia orden venta
                if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP' and bid > last * (1 - ganancia):
                    print(time.strftime("%H:%M:%S"),'trailingSTOP',end=' ')
                    enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),cantidad,nroCelda)
                    shtTest.range('T1').value = 0.25
                else: shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP'  
        else: #TRAILING sobre bonos / letras / ons
            if bid / 100 > costo * (1 + (ganancia/25)): # Precio sube activo trailing y sube % ganancia               
                if str(shtTest.range('W'+str(int(nroCelda+1))).value)=='TRAILING': shtTest.range('T1').value*= 1+0.25
                else: shtTest.range('W'+str(int(nroCelda+1))).value = 'TRAILING'
                shtTest.range('X'+str(int(nroCelda+1))).value = bid / 100
            if last / 100 < costo * (1 - ganancia/25): # Precio baja activo stop y envia orden venta
                if str(shtTest.range('W'+str(int(nroCelda+1))).value)=='STOP' and (bid/100)>(last/100)*(1-(ganancia/25)):
                    print(time.strftime("%H:%M:%S"),'trailingSTOP',end=' ')
                    enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),cantidad,nroCelda)
                    shtTest.range('T1').value = 0.25
                else: shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP' 
    except: pass
#########################################################################################################

while True:

    if str(shtTest.range('A1').value) != 'symbol': ilRulo()
    if str(shtTest.range('R1').value).lower() =='b': time.sleep(1)
    else: time.sleep(3)

    for valor in shtTest.range('P26:V59').value:
        if str(shtTest.range('S1').value).lower()!='n' and int(valor[6])>0: # Activa TRAILING STOP ________
            trailingStop('A'+str((int(valor[0])+1)),valor[6],valor[0])
        try: # CANCELAR todas las ordenes _________________________________________________________________
            if str(valor[5]).lower() == 'c': 
                ##hb.orders.cancel_order(int(os.environ.get('account_id')),orderC)
                shtTest.range('U'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = 0
                print("Orden compra fue cancelada")
            if str(valor[5]).lower() == 'v': 
                ##hb.orders.cancel_order(int(os.environ.get('account_id')),orderV)
                shtTest.range('U'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = 0
                print("Orden venta fue cancelada")
            if str(valor[5]).lower() == 'x': 
                ##hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
                shtTest.range('U'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = 0
                print("Todas las ordenes activas canceladas")
        except: 
            shtTest.range('U'+str(int(valor[0]+1))).value = 0
            print('Error, al cancelar orden.')

        if valor[1]: # COMPRAR precio BID ___________________________________________________________
            try: enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0])
            except: pass
        elif valor[2]: # COMPRAR precio ASK _________________________________________________________
            try: enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0])
            except: pass
        elif valor[3]: # VENDER precio BID __________________________________________________________
            try: enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0])
            except: pass
        elif valor[4]: # VENDER precio ASK __________________________________________________________
            try: enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0])
            except: pass
        
        elif valor[5] == '-' or valor[5] == '+': # buy//sell usando puntas ________________________________
            try: cantidad = int(shtTest.range('Y'+str(int(valor[0]+1))).value)
            except: cantidad = int(shtTest.range('W1').value)
            if valor[5] == '-':
                enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
            else: enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidad,valor[0])
            shtTest.range('U'+str(int(valor[0]+1))).value = 0
#[ ]><   \n