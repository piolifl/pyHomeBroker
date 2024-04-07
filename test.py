from pyhomebroker import HomeBroker     
import xlwings as xw                    
import pandas as pd                     
from datetime import date, timedelta
import time
import winsound


wb = xw.Book('D:\\pyHomeBroker\\epgb_pyHB.xlsx')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')
shtTest.range('Q1').value = 'PRC'
shtTest.range('R1').value ='TRAIL'
shtTest.range('S1').value ='STOP'
shtTest.range('T1').value = 0.001
shtTest.range('U1').value = 1
shtTest.range('V1').value = 0
shtTest.range('W1').value = 1


#-------------------------------------------------------------------------------------------------------
print(time.strftime("%H:%M:%S"),"Inicando TESTER ...")

def namesArs(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA37D'+plazo
    elif nombre[:2] == 'BP': return 'BPOA7'+plazo
    elif nombre[:2] == 'KO': return 'KO'+plazo
    elif nombre[:2] == 'GOGL': return 'GOOGL'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') and (nombre[3:4] == 'D' or nombre[3:4] == 'C'):
        if (nombre[1:2] == 'F' or nombre[1:2] == 'Y'): return nombre[:1]+'20'+nombre[1:3]+plazo
        else: return nombre[:1]+'18'+nombre[1:3]+plazo
    elif (nombre[:2] == 'MR' or nombre[:2] == 'CL') and (nombre[4:5] == 'D' or nombre[4:5] == 'C'):
        return nombre[:4]+'O'+plazo 
    else: return nombre[:4]+plazo

def namesCcl(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA7DC'+plazo
    elif nombre[:2] == 'BP': return 'BPA7C'+plazo
    elif nombre[:2] == 'KO': return 'KOC'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') :
        if nombre[3:4] == 'D': return nombre[:3]+'C'+plazo
        else: return nombre[:1]+nombre[3:5]+'C'+plazo
    else: return nombre[:4]+'C'+plazo

def namesMep(nombre,plazo): 
    if nombre[:2] == 'BA': return 'BA7DD'+plazo
    elif nombre[:2] == 'BP': return 'BPA7D'+plazo
    elif nombre[:2] == 'KO': return 'KOD'+plazo
    elif (nombre[:1] == 'X' or nombre[:1] == 'S') :
        if nombre[3:4] == 'C': return nombre[:3]+'D'+plazo
        else: return nombre[:1]+nombre[3:5]+'D'+plazo
    else: return nombre[:4]+'D'+plazo

def cargoXplazo(dicc):
    mejorMep = dicc['mepCI'][0]
    if mejorMep == 'AL30D - spot': shtTest.range('A2').value = ''
    else: shtTest.range('A2').value = mejorMep
    shtTest.range('A3').value = 'AL30D - spot'
    shtTest.range('A4').value = 'AL30 - spot'
    shtTest.range('A5').value = namesArs(dicc['mepCI'][0],' - spot')

    mejorMep = dicc['mep48'][0]
    if mejorMep == 'AL30D - 48hs': shtTest.range('A6').value = ''
    else: shtTest.range('A6').value = mejorMep
    shtTest.range('A7').value = 'AL30D - 48hs'
    shtTest.range('A8').value = 'AL30 - 48hs'
    shtTest.range('A9').value = namesArs(dicc['mep48'][0],' - 48hs')
    
    shtTest.range('A10').value = dicc['mepCI'][0]
    shtTest.range('A11').value = namesMep(dicc['arsCImep'][0],' - spot')
    shtTest.range('A12').value = dicc['arsCImep'][0]
    shtTest.range('A13').value = namesArs(dicc['mepCI'][0],' - spot')
    shtTest.range('A14').value = dicc['mep48'][0]
    shtTest.range('A15').value = namesMep(dicc['ars48mep'][0],' - 48hs')
    shtTest.range('A16').value = dicc['ars48mep'][0]
    shtTest.range('A17').value = namesArs(dicc['mep48'][0],' - 48hs')

    shtTest.range('A18').value = dicc['mepCI'][0]
    shtTest.range('A19').value = namesMep(dicc['cclCI'][0],' - spot')
    shtTest.range('A20').value = dicc['cclCI'][0]
    shtTest.range('A21').value = namesCcl(dicc['mepCI'][0],' - spot')
    shtTest.range('A22').value = dicc['mep48'][0]
    shtTest.range('A23').value = namesMep(dicc['ccl48'][0],' - 48hs')
    shtTest.range('A24').value = dicc['ccl48'][0]
    shtTest.range('A25').value = namesCcl(dicc['mep48'][0],' - 48hs')

    shtTest.range('A1').value = 'symbol'
    winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)

def ilRulo():
    celda,pesos,dolar = 64,1000,0.01
    tikers = {'cclCI':['',dolar],'ccl48':['',dolar],'mepCI':['',dolar],'mep48':['',dolar],'arsCIccl':['',pesos],'ars48ccl':['',pesos],'arsCImep':['',pesos],'ars48mep':['',pesos]}
    
    for valor in shtTest.range('A64:A201').value:
        if not valor: continue
        name = str(valor).split()
        if str(name[2]).lower() == 'spot':
            
            if str(name[0][-1:]).upper()=='C':
                arsC = shtTest.range('AA'+str(celda)).value
                if not arsC: arsC = 1000
                if arsC > tikers['arsCIccl'][1]: tikers['arsCIccl'] = [namesArs(name[0],' - spot'),arsC]
                ccl = shtTest.range('Z'+str(celda)).value
                if not ccl: ccl = 0.01
                if ccl > tikers['cclCI'][1]: tikers['cclCI'] = [valor,ccl]

            if str(name[0][-1:]).upper()=='D': 
                arsM = shtTest.range('AA'+str(celda)).value
                if not arsM: arsM = 1000
                if arsM > tikers['arsCImep'][1]: tikers['arsCImep'] = [namesArs(name[0],' - spot'),arsM]
                mep = shtTest.range('Z'+str(celda)).value
                if not mep: mep = 0.01
                if mep > tikers['mepCI'][1]: tikers['mepCI'] = [valor,mep]

        if str(name[2]) == '48hs':
            if str(name[0][-1:]).upper()=='C':
                arsC = shtTest.range('AA'+str(celda)).value
                if not arsC: arsC = 1000
                if arsC > tikers['ars48ccl'][1]: tikers['ars48ccl'] = [namesArs(name[0],' - 48hs'),arsC]
                ccl = shtTest.range('Z'+str(celda)).value
                if not ccl: ccl = 0.01
                if ccl > tikers['ccl48'][1]: tikers['ccl48'] = [valor,ccl]

            if str(name[0][-1:]).upper()=='D': 
                arsM = shtTest.range('AA'+str(celda)).value
                if not arsM: arsM = 1000
                if arsM > tikers['ars48mep'][1]: tikers['ars48mep'] = [namesArs(name[0],' - 48hs'),arsM]
                mep = shtTest.range('Z'+str(celda)).value
                if not mep: mep = 0.01
                if mep > tikers['mep48'][1]: tikers['mep48'] = [valor,mep]
        celda +=1
    cargoXplazo(tikers)

############################################ ENVIAR ORDENES ################################################    
def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    global orderC, orderV
    symbol = str(shtTest.range(str(symbol)).value).split()
    precio = shtTest.range(str(price)).value
    menosRecompra = float(shtTest.range('U1').value)
    if not shtTest.range('V'+str(int(celda+1))).value: shtTest.range('V'+str(int(celda+1))+':'+'X'+str(int(celda+1))).value = 0
    if tipo.lower() == 'buy': 
        try: 
            if len(symbol) < 2:
                if str(shtTest.range('R1').value) == 'REC': 
                    if not menosRecompra: 
                        precio -= 1
                        shtTest.range('U1').value = 1
                    else:  precio -= menosRecompra / 10
                    shtTest.range('R1').value = ''
                    print(f'{time.strftime("%H:%M:%S")} RECOMPRA ',end=' || ')
                #orderC = hb.orders.send_buy_order(symbol[0],'24hs', float(precio), int(size))
                try: shtTest.range('V'+str(int(celda+1))).value += int(size)
                except: shtTest.range('V'+str(int(celda+1))).value = int(size)
                try: shtTest.range('W'+str(int(celda+1))).value += int(size) * precio*100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio*100
                print(f'Buy  {symbol[0]} // cantidad: + {int(size)} // precio: {precio}')
            else:
                if str(shtTest.range('R1').value) == 'REC': 
                    if not menosRecompra: 
                        precio -= 100
                        shtTest.range('U1').value = 1
                    else:  precio -= menosRecompra * 10
                    #shtTest.range('Q'+str(int(celda+1))).value = cantidad +1 
                    shtTest.range('R1').value = ''
                    print(f'{time.strftime("%H:%M:%S")} RECOMPRA ',end=' || ')
                #orderC = hb.orders.send_buy_order(symbol[0],symbol[2], float(precio), int(size))
                try: shtTest.range('V'+str(int(celda+1))).value += int(size)
                except: shtTest.range('V'+str(int(celda+1))).value = int(size)
                try: shtTest.range('W'+str(int(celda+1))).value += int(size) * precio/100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio/100
                print(f'Buy  {symbol[0]} {symbol[2]} // cantidad: + {int(size)} // precio {round(precio/100,2)}')
        except: 
            winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
            shtTest.range('Q'+str(int(celda+1))+':'+'U'+str(int(celda+1))).value = ''
            print(time.strftime("%H:%M:%S"), 'Error en COMPRA.')
    else: 
        try:
            if len(symbol) < 2:
                #orderV = hb.orders.send_sell_order(symbol[0],'24hs', float(precio), int(size))
                try: shtTest.range('V'+str(int(celda+1))).value -= int(size)
                except: shtTest.range('V'+str(int(celda+1))).value = int(size)
                try: shtTest.range('W'+str(int(celda+1))).value -= int(size) * precio*100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio*100
                print(f'Sell {symbol[0]} // cantidad: - {int(size)} // precio: {precio}')
            else:
                #orderV = hb.orders.send_sell_order(symbol[0],symbol[2], float(precio), int(size))
                try: shtTest.range('V'+str(int(celda+1))).value -= int(size)
                except: shtTest.range('V'+str(int(celda+1))).value = int(size)
                try: shtTest.range('W'+str(int(celda+1))).value -= int(size) * precio/100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio/100
                print(f'Sell {symbol[0]} {symbol[2]} // cantidad: - {int(size)} // precio: {round(precio/100,2)}')
        except:
            winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
            shtTest.range('Q'+str(int(celda+1))+':'+'U'+str(int(celda+1))).value = ''
            print(time.strftime("%H:%M:%S"), 'Error en VENTA.')

    try: shtTest.range('X'+str(int(celda+1))).value=shtTest.range('W'+str(int(celda+1))).value / shtTest.range('V'+str(int(celda+1))).value
    except: pass
    shtTest.range('Q'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
############################################ TRAILING STOP ################################################
def trailingStop(nombre=str,cantidad=int,nroCelda=int):
    try:
        nombre = str(shtTest.range(str(nombre)).value).split()
        bid = float(shtTest.range('C'+str(int(nroCelda+1))).value)
        bid_size = int(shtTest.range('B'+str(int(nroCelda+1))).value)
        stock = int(shtTest.range('V'+str(int(nroCelda+1))).value)
        last = float(shtTest.range('F'+str(int(nroCelda+1))).value)
        costo = float(shtTest.range('X'+str(int(nroCelda+1))).value) 
        try: ganancia = float(shtTest.range('T1').value)
        except:
            shtTest.range('T1').value = 0.001
            ganancia = float(shtTest.range('T1').value)
        if cantidad > stock : cantidad = stock
        if cantidad > bid_size : cantidad = bid_size
        if len(nombre) < 2: #TRAILING sobre opciones financieras
            if bid * 100 > costo * (1 + (ganancia*25)): # Precio sube activo trailing y sube % ganancia 
                shtTest.range('W'+str(int(nroCelda+1))).value = 'TRAILING'
                shtTest.range('X'+str(int(nroCelda+1))).value = bid * 100
            
            if not shtTest.range('S1').value:
                if last * 100 < costo * (1 - (ganancia*10)): # Precio baja activo stop y envia orden venta
                    if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP' and bid>last*(1-(ganancia*10)):
                        print(f'{time.strftime("%H:%M:%S")} STOP     ',end=' || ')
                        shtTest.range('R1').value = 'REC'
                        shtTest.range('W'+str(int(nroCelda+1))).value = ''
                        shtTest.range('X'+str(int(nroCelda+1))).value = bid * 100
                        enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),cantidad,nroCelda)
                    else: shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP'  
        else: #TRAILING sobre bonos / letras / ons
            if bid / 100 > costo * (1 + ganancia): # Precio sube activo trailing y sube % ganancia               
                shtTest.range('W'+str(int(nroCelda+1))).value = 'TRAILING'
                shtTest.range('X'+str(int(nroCelda+1))).value = round(bid / 100,5)
            
            if not shtTest.range('S1').value:
                if last / 100 < costo * (1 - ganancia): # Precio baja activo stop y envia orden venta
                    if str(shtTest.range('W'+str(int(nroCelda+1))).value)=='STOP' and (bid/100)>(last/100)*(1-ganancia):
                        print(f'{time.strftime("%H:%M:%S")} STOP     ',end=' || ')
                        shtTest.range('R1').value = 'REC'
                        shtTest.range('W'+str(int(nroCelda+1))).value = ''
                        shtTest.range('X'+str(int(nroCelda+1))).value = round(bid / 100,5)
                        enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),cantidad,nroCelda)
                    else: shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP' 
    except: pass
########################################### CARGA BUCLE EN EXCEL ##########################################
while True:
    for valor in shtTest.range('P2:V59').value:
        if valor[1]: # COMPRAR precio BID _________________________________________________________________
            try:   enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0])
            except: 
                shtTest.range('Q'+str(int(valor[0]+1))).value = ''
                winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
        if valor[2]: # COMPRAR precio ASK _______________________________________________________________
            try:  enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0])
            except: 
                shtTest.range('R'+str(int(valor[0]+1))).value = ''
                winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
        if valor[3]: # VENDER precio BID ________________________________________________________________
            try:  enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0])
            except: 
                shtTest.range('S'+str(int(valor[0]+1))).value = ''
                winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
        if valor[4]: # VENDER precio ASK ________________________________________________________________
            try:  enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0])
            except: 
                shtTest.range('T'+str(int(valor[0]+1))).value = ''
                winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
        if valor[5]:
            try: # CANCELAR todas las ordenes _____________________________________________________________
                if str(valor[5]).lower() == 'c': 
                    #hb.orders.cancel_order(int(os.environ.get('account_id')),orderC)
                    shtTest.range('U'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = ''
                    print("Orden compra fue cancelada")
                elif str(valor[5]).lower() == 'v': 
                    #hb.orders.cancel_order(int(os.environ.get('account_id')),orderV)
                    shtTest.range('U'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = ''
                    print("Orden venta fue cancelada")
                elif str(valor[5]).lower() == 'x': 
                    #hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
                    shtTest.range('U'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = ''
                    print("Todas las ordenes activas canceladas")
            except: 
                winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                shtTest.range('U'+str(int(valor[0]+1))).value = ''
                print('Error, al cancelar orden.')

            if valor[5] == '-' or valor[5] == '+': # buy//sell usando puntas ______________________________
                try: cantidad = int(shtTest.range('Y'+str(int(valor[0]+1))).value)
                except: cantidad = 1
                if valor[5] == '-':enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidad,valor[0])
                else: enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
                shtTest.range('U'+str(int(valor[0]+1))).value = ''

        if not shtTest.range('R1').value: # Activa TRAILING  __________________________________________
            if time.strftime("%H:%M:%S") > '16:24:50': shtTest.range('R1').value = 'TRAIL'
            try: stock = int(valor[6])
            except: stock = 0
            if stock > 0:
                if not shtTest.range('Y'+str(int(valor[0]+1))).value: cantidad = 1
                else: cantidad = int(shtTest.range('Y'+str(int(valor[0]+1))).value)
                trailingStop('A'+str((int(valor[0])+1)),cantidad,int(valor[0]))

        if str(shtTest.range('R1').value).upper() == 'REC': # Activa RECOMPRA AUTOMATICA _____________
            try: 
                enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
            except: 
                winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                print(time.strftime("%H:%M:%S"), 'Error RECOMPRA Automatica. Intenta compra en punta BID')
                shtTest.range('U'+str(int(valor[0]+1))).value = '+'
            
    time.sleep(2)
    if str(shtTest.range('A1').value) != 'symbol': ilRulo()
#[ ]><   \n