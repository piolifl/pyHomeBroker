from pyhomebroker import HomeBroker     
import xlwings as xw                    
import pandas as pd                     
from datetime import date, timedelta
import time
import winsound
import os
import environ

env = environ.Env()
environ.Env.read_env()
wb = xw.Book('D:\\pyHomeBroker\\epgb_pyHB.xlsx')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')

#-------------------------------------------------------------------------------------------------------

hb = HomeBroker(int(os.environ.get('broker')))

hb.auth.login(dni=str(os.environ.get('dni')),
              user=str(os.environ.get('user')),
              password=str(os.environ.get('password')),
              raise_exception=True)

#-------------------------------------------------------------------------------------------------------
print(time.strftime("%H:%M:%S"),f"Logueo SOLO ORDENES: {os.environ.get('name')} cuenta: {int(os.environ.get('account_id'))}")

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
                        shtTest.range('U1').value = 10
                    else:  precio -= menosRecompra / 10
                    shtTest.range('R1').value = ''
                    print(f'{time.strftime("%H:%M:%S")} RECOMPRA ',end=' || ')
                orderC = hb.orders.send_buy_order(symbol[0],'24hs', float(precio), int(size))
                print(f'Buy  {symbol[0]} // cantidad: + {int(size)} // precio: {precio}')
                try: shtTest.range('V'+str(int(celda+1))).value += int(size)
                except: shtTest.range('V'+str(int(celda+1))).value = int(size)
                try: shtTest.range('W'+str(int(celda+1))).value += int(size) * precio*100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio*100
            else:
                if str(shtTest.range('R1').value) == 'REC': 
                    if not menosRecompra: 
                        precio -= 100
                        shtTest.range('U1').value = 10
                    else:  precio -= menosRecompra * 10
                    #shtTest.range('Q'+str(int(celda+1))).value = cantidad +1 
                    shtTest.range('R1').value = ''
                    print(f'{time.strftime("%H:%M:%S")} RECOMPRA ',end=' || ')
                orderC = hb.orders.send_buy_order(symbol[0],symbol[2], float(precio), int(size))
                print(f'Buy  {symbol[0]} {symbol[2]} // cantidad: + {int(size)} // precio {round(precio/100,2)}')
                try: shtTest.range('V'+str(int(celda+1))).value += int(size)
                except: shtTest.range('V'+str(int(celda+1))).value = int(size)
                try: shtTest.range('W'+str(int(celda+1))).value += int(size) * precio/100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio/100
        except: 
            winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
            shtTest.range('Q'+str(int(celda+1))+':'+'U'+str(int(celda+1))).value = ''
            print('Error al enviar Compra.')
    else: 
        try:
            if len(symbol) < 2:
                print(f'Sell {symbol[0]} // cantidad: - {int(size)} // precio: {precio}')
                orderV = hb.orders.send_sell_order(symbol[0],'24hs', float(precio), int(size))
                try: shtTest.range('V'+str(int(celda+1))).value -= int(size)
                except: shtTest.range('V'+str(int(celda+1))).value = int(size)
                try: shtTest.range('W'+str(int(celda+1))).value -= int(size) * precio*100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio*100
            else:
                print(f'Sell {symbol[0]} {symbol[2]} // cantidad: - {int(size)} // precio: {round(precio/100,2)}')
                orderV = hb.orders.send_sell_order(symbol[0],symbol[2], float(precio), int(size))
                try: shtTest.range('V'+str(int(celda+1))).value -= int(size)
                except: shtTest.range('V'+str(int(celda+1))).value = int(size)
                try: shtTest.range('W'+str(int(celda+1))).value -= int(size) * precio/100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio/100
        except:
            winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
            shtTest.range('Q'+str(int(celda+1))+':'+'U'+str(int(celda+1))).value = ''
            print('Error al enviar Venta.')

    shtTest.range('X'+str(int(celda+1))).value=shtTest.range('W'+str(int(celda+1))).value / shtTest.range('V'+str(int(celda+1))).value
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
    for valor in shtTest.range('P22:V25').value:
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
                    hb.orders.cancel_order(int(os.environ.get('account_id')),orderC)
                    shtTest.range('U'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = ''
                    print("Orden compra fue cancelada")
                elif str(valor[5]).lower() == 'v': 
                    hb.orders.cancel_order(int(os.environ.get('account_id')),orderV)
                    shtTest.range('U'+str(int(valor[0]+1))+':'+'X'+str(int(valor[0]+1))).value = ''
                    print("Orden venta fue cancelada")
                elif str(valor[5]).lower() == 'x': 
                    hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
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
            try: stock = int(valor[6])
            except: stock = 0
            if stock > 0:
                if not shtTest.range('Y'+str(int(valor[0]+1))).value: cantidad = 1
                else: cantidad = int(shtTest.range('Y'+str(int(valor[0]+1))).value)
                trailingStop('A'+str((int(valor[0])+1)),cantidad,int(valor[0]))

        if str(shtTest.range('R1').value).upper() == 'REC': # Activa RECOMPRA AUTOMATICA _____________
            try: 
                enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
                shtTest.range('Q'+str(int(valor[0]+1))).value = 1
            except: 
                winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                print('Error RECOMPRA Automatica.')
            
    time.sleep(2)
    if time.strftime("%H:%M:%S") > '17:03:00': break 
      
print(time.strftime("%H:%M:%S"), 'Mercado cerrado.')
  
#[ ]><   \n
