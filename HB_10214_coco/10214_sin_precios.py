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

#getGrupos()

#-------------------------------------------------------------------------------------------------------
print(time.strftime("%H:%M:%S"),f"SOLO ORDENES en: {os.environ.get('name')} cuenta: {int(os.environ.get('account_id'))}")

############################################ ENVIAR ORDENES ################################################    
def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    global orderC, orderV
    symbol = str(shtTest.range(str(symbol)).value).split()
    por = int(shtTest.range('W1').value)
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

            else:
                if str(shtTest.range('R1').value) == 'REC': 
                    if not menosRecompra: 
                        precio -= 100
                        shtTest.range('U1').value = 10
                    else:  precio -= menosRecompra * 10
                    shtTest.range('R1').value = ''
                    print(f'{time.strftime("%H:%M:%S")} RECOMPRA ',end=' || ')
                orderC = hb.orders.send_buy_order(symbol[0],symbol[2], float(precio), int(size*por))
                print(f'Buy  {symbol[0]} {symbol[2]} // cantidad: + {int(size*por)} // precio {round(precio/100,2)}')
                
        except: 
            winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
            print('Error al enviar Compra.')
    else: 
        try:
            if len(symbol) < 2:
                orderV = hb.orders.send_sell_order(symbol[0],'24hs', float(precio), int(size))
                print(f'Sell {symbol[0]} // cantidad: - {int(size)} // precio: {precio}')
                
            else:
                orderV = hb.orders.send_sell_order(symbol[0],symbol[2], float(precio), int(size*por))
                print(f'Sell {symbol[0]} {symbol[2]} // cantidad: - {int(size*por)} // precio: {round(precio/100,2)}')
                
        except:
            winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
            print('Error al enviar Venta.')

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

        if str(shtTest.range('R1').value).upper() == 'REC': # Activa RECOMPRA AUTOMATICA _____________
            try:   enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
            except: 
                winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                shtTest.range('R1').value = ''
                print('Error RECOMPRA Automatica.')
            shtTest.range('Q'+str(int(valor[0]+1))).value = 1

    if time.strftime("%H:%M:%S") > '17:03:00': break 

print(time.strftime("%H:%M:%S"), 'Mercado cerrado.')
  
#[ ]><   \n
