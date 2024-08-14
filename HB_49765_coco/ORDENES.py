from pyhomebroker import HomeBroker     
import xlwings as xw                    
import pandas as pd                     
from datetime import date, timedelta
import time
import winsound
import os
import environ
import requests

env = environ.Env()
environ.Env.read_env()
#wb = xw.Book('..\\epgb_pyHB.xlsx')
wb = xw.Book('..\\epgb_pyHB.xlsb')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')

rangoDesde = '26'
rangoHasta = '89'



def login():
    hb.auth.login(dni=str(os.environ.get('dni')), 
    user=str(os.environ.get('user')),  
    password=str(os.environ.get('password')),
    raise_exception=True)

hb = HomeBroker(int(os.environ.get('broker')))
login()


def getPortfolio(hb, comitente):
    try:
        shtTest.range('U'+str(rangoDesde)+':'+'U'+str(rangoHasta)).value = ''
        payload = {'comitente': str(comitente),
        'consolida': '0',
        'proceso': '22',
        'fechaDesde': None,
        'fechaHasta': None,
        'tipo': None,
        'especie': None,
        'comitenteMana': None}
        
        if os.environ.get('name') == 'COCOS.CAPITAL':
            portfolio = requests.post("https://cocoscap.com/Consultas/GetConsulta", cookies=hb.auth.cookies, json=payload).json()
        else: 
            portfolio = requests.post("https://clientes.bcch.org.ar/Consultas/GetConsulta", cookies=hb.auth.cookies, json=payload).json()
        subtotal = [ (i['DETA'],i['IMPO']) for i in portfolio["Result"]["Totales"]["Detalle"] ]
        print(subtotal)
        subtotal = [ i['Subtotal'] for i in portfolio["Result"]["Activos"][0:] ]
        for i in subtotal[0:]:
            if i[0]['NERE'] != 'Pesos':  
                subtotal = [ ( x['NERE'],x['CAN0'],x['CANT']) for x in i[0:] if x['CANT'] != None]
                for x in subtotal:
                    for valor in shtTest.range('A'+str(rangoDesde)+':'+'P'+str(rangoHasta)).value:
                        if not valor[0]: continue
                        ticker = str(valor[0]).split()
                        if x[0] == ticker[0]: 
                            shtTest.range('U'+str(int(valor[15]+1))).value = x[2]
                            if not shtTest.range('V'+str(int(valor[15]+1))).value:
                                if len(ticker) < 2: 
                                    shtTest.range('X'+str(int(valor[15]+1))).value = x[1]
                                else:
                                    try: shtTest.range('X'+str(int(valor[15]+1))).value = x[1] /100
                                    except: shtTest.range('X'+str(int(valor[15]+1))).value = x[1]
    except: pass

#--------------------------------------------------------------------------------------------------------------------------------
print(time.strftime("%H:%M:%S"),f"solo ORDENES: {os.environ.get('name')} cuenta: {int(os.environ.get('account_id'))}")
#--------------------------------------------------------------------------------------------------------------------------------

def cancelaCompra(celda):
    try:
        orderC = shtTest.range('AB'+str(int(celda+1))).value
        if not orderC: orderC = 0
        hb.orders.cancel_order(int(os.environ.get('account_id')),int(orderC))
        shtTest.range('V'+str(int(celda+1))).value -= shtTest.range('AC'+str(int(celda+1))).value
        shtTest.range('AB'+str(int(celda+1))+':'+'AD'+str(int(celda+1))).value = ''
        print(f" /// Cancela Compra nro: {int(orderC)} ",time.strftime("%H:%M:%S"))
    except: print(time.strftime("%H:%M:%S"),'______ ERROR al cancelar orden.')

def cancelarVenta(celda):
    try:
        orderV = shtTest.range('AE'+str(int(celda+1))).value
        if not orderV: orderV = 0
        hb.orders.cancel_order(int(os.environ.get('account_id')),int(orderV))
        shtTest.range('V'+str(int(celda+1))).value += shtTest.range('AF'+str(int(celda+1))).value
        shtTest.range('AE'+str(int(celda+1))+':'+'AG'+str(int(celda+1))).value = ''
        print(f" /// Cancela Venta nro : {int(orderV)} ",time.strftime("%H:%M:%S"))
    except: print(time.strftime("%H:%M:%S"),'______ ERROR al cancelar orden.')

def cancelarTodo(desde,hasta):
    try:
        hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
        shtTest.range('AB'+str(desde)+':'+'AH'+str(hasta)).value = ''
        print(" /// Todas las ordenes activas canceladas ",time.strftime("%H:%M:%S"))
    except: print(time.strftime("%H:%M:%S"),'______ ERROR al cancelar orden.')

###############################################################  ENVIAR ORDENES ################################################    
def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    global orderC, orderV
    orderC, orderV = 0,0
    symbol = str(shtTest.range(str(symbol)).value).split()
    precio = shtTest.range(str(price)).value
    if not shtTest.range('V'+str(int(celda+1))).value: shtTest.range('W'+str(int(celda+1))+':'+'X'+str(int(celda+1))).value = 0
    if tipo.lower() == 'buy': 
        try: 
            if len(symbol) < 2:
                orderC = hb.orders.send_buy_order(symbol[0],'24hs', float(precio), int(size))
                shtTest.range('AD'+str(int(celda+1))).value = float(precio)
                try: shtTest.range('W'+str(int(celda+1))).value += int(size) * precio*100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio*100
                print(f'______ BUY  opcion {symbol[0]} 24hs // precio: {precio} // + {int(size)} // orden: {orderC}') 
            else:
                orderC = hb.orders.send_buy_order(symbol[0],symbol[2], float(precio), int(size))
                shtTest.range('AD'+str(int(celda+1))).value = float(precio/100)
                try: shtTest.range('W'+str(int(celda+1))).value += int(size) * precio/100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio/100
                print(f'______ BUY  {symbol[0]} {symbol[2]} // precio: {round(precio/100,4)} // + {int(size)} // orden: {orderC}')
            # ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            shtTest.range('Q'+str(int(celda+1))+':'+'R'+str(int(celda+1))).value = ''
            try: shtTest.range('V'+str(int(celda+1))).value += int(size)
            except: shtTest.range('V'+str(int(celda+1))).value = int(size)
            shtTest.range('AB'+str(int(celda+1))).value = orderC
            shtTest.range('AC'+str(int(celda+1))).value = int(size)
            shtTest.range('AH'+str(int(celda+1))).value = str(time.strftime("%H:%M:%S"))
        except: 
            shtTest.range('Q'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            print(f'______ ERROR en COMPRA. {symbol[0]} // precio: {precio} // + {int(size)}')
    else: 
        try:
            if len(symbol) < 2:
                orderV = hb.orders.send_sell_order(symbol[0],'24hs', float(precio), int(size))
                shtTest.range('AG'+str(int(celda+1))).value = float(precio)
                try: shtTest.range('W'+str(int(celda+1))).value -= int(size) * precio*100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio*100
                print(f'______ SELL opcion {symbol[0]} 24hs // precio: {precio} // - {int(size)} // orden: {orderV}')
            else:
                orderV = hb.orders.send_sell_order(symbol[0],symbol[2], float(precio), int(size))
                shtTest.range('AG'+str(int(celda+1))).value = float(precio/100)
                try: shtTest.range('W'+str(int(celda+1))).value -= int(size) * precio/100
                except: shtTest.range('W'+str(int(celda+1))).value = int(size) * precio/100
                print(f'______ SELL {symbol[0]} {symbol[2]} // precio: {round(precio/100,4)} // - {int(size)} // orden: {orderV}')
            # ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            shtTest.range('S'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            try: shtTest.range('V'+str(int(celda+1))).value -= int(size)
            except: shtTest.range('V'+str(int(celda+1))).value = int(size)/-1
            shtTest.range('AE'+str(int(celda+1))).value = orderV
            shtTest.range('AF'+str(int(celda+1))).value = int(size)
            shtTest.range('AH'+str(int(celda+1))).value = str(time.strftime("%H:%M:%S"))
        except:
            shtTest.range('Q'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            print(f'______ ERROR en VENTA. {symbol[0]} // precio: {precio} // {int(size)/-1}')
    try: 
        tieneW = shtTest.range('W'+str(int(celda+1))).value
        tieneV = shtTest.range('V'+str(int(celda+1))).value
        if tieneW != 'TRAILING' or tieneW != 'STOP' or tieneW != '': 
            shtTest.range('X'+str(int(celda+1))).value = tieneW / tieneV
        else: 
            shtTest.range('W'+str(int(celda+1))).value = ''
            shtTest.range('X'+str(int(celda+1))).value = 0
    except: 
        shtTest.range('W'+str(int(celda+1))).value = ''
        shtTest.range('X'+str(int(celda+1))).value = 0
############################################################### TRAILING STOP #################################################
def trailingStop(nombre=str,cantidad=int,nroCelda=int):
    try:
        nombre = str(shtTest.range(str(nombre)).value).split()
        bid = shtTest.range('C'+str(int(nroCelda+1))).value
        stock = shtTest.range('V'+str(int(nroCelda+1))).value
        last = shtTest.range('F'+str(int(nroCelda+1))).value
        costo = shtTest.range('X'+str(int(nroCelda+1))).value 
        ganancia = shtTest.range('Z1').value
        if not ganancia: ganancia = 0.0005
        if cantidad > stock : cantidad = int(stock)

        if len(nombre) < 2: # Ingresa si son OPCIONES ///////////////////////////////////////////////////////////////////////////
            if bid * 100 > costo * (1 + (ganancia*10)):
                if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'TRAILING': pass
                else: shtTest.range('W'+str(int(nroCelda+1))).value = 'TRAILING'
                shtTest.range('X'+str(int(nroCelda+1))).value = bid * 100
            if not shtTest.range('X1').value:
                if last * 100 < costo * (1 - (ganancia*75)): # Precio baja activo stop y envia orden venta
                    if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP' and bid > last * (1-(ganancia*15)):
                        shtTest.range('W'+str(int(nroCelda+1))).value = ''
                        shtTest.range('X'+str(int(nroCelda+1))).value = 0
                        if shtTest.range('Y'+str(int(nroCelda+1))).value : 
                            enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),cantidad,nroCelda)
                    else:
                        if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                        else:
                            shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP'
                            winsound.PlaySound("SystemHand", winsound.SND_ALIAS)      

        else: # Ingresa si son BONOS / LETRAS / ON / CEDEARS ////////////////////////////////////////////////////////////////////
            if time.strftime("%H:%M:%S") > '16:24:50' and str(nombre[2]).lower() == 'spot': 
                shtTest.range('W'+str(int(nroCelda+1))).value = "CLOSED"
                pass
            if time.strftime("%H:%M:%S") > '16:56:50' and str(nombre[2]).lower() == '24hs': 
                shtTest.range('W'+str(int(nroCelda+1))).value = "CLOSED"
                pass
            else:
                # Rutina, si el precio BID sube modifica precio promedio de compra //////////////////////////////////////////////
                if bid / 100 > costo * (1 + ganancia):             
                    if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'TRAILING': pass
                    else: shtTest.range('W'+str(int(nroCelda+1))).value = 'TRAILING'
                    shtTest.range('X'+str(int(nroCelda+1))).value = round(bid / 100,5)

                # Si X1 esta vacio, habilita estrategias de ventas  ////////////////////////////////////////////////////////////
                if not shtTest.range('X1').value:
                    #  Precio LAST baja, inica estrategia salida vendiendo stock spot en 24hs
                    if last / 100 < costo * (1 - ganancia):
                        if str(nombre[2]).lower() == 'spot':
                            bid2 = shtTest.range('C'+str(int(nroCelda+2))).value
                            last2 = shtTest.range('F'+str(int(nroCelda+2))).value
                            if str(shtTest.range('W'+str(int(nroCelda+1))).value)=='STOP' and (bid2/100)>(last2/100)*(1-ganancia):
                                shtTest.range('W'+str(int(nroCelda+1))).value = ''
                                shtTest.range('X'+str(int(nroCelda+1))).value = 0
                                try: shtTest.range('V'+str(int(nroCelda+1))).value -= cantidad
                                except: shtTest.range('V'+str(int(nroCelda+1))).value = cantidad/-1
                                print(f'{time.strftime("%H:%M:%S")} STOP vendo    ',end=' || ')
                                if shtTest.range('Y'+str(int(nroCelda+1))).value : 
                                    enviarOrden('sell','A'+str((int(nroCelda)+2)),'C'+str((int(nroCelda)+2)),cantidad,nroCelda+1)
                            else: 
                                if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                                else:
                                    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                                    shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP'
                        else:
                            if str(shtTest.range('W'+str(int(nroCelda+1))).value)=='STOP' and (bid/100)>(last/100)*(1-ganancia):
                                shtTest.range('W'+str(int(nroCelda+1))).value = ''
                                shtTest.range('X'+str(int(nroCelda+1))).value = 0
                                print(f'{time.strftime("%H:%M:%S")} STOP vendo    ',end=' || ')
                                if shtTest.range('Y'+str(int(nroCelda+1))).value : 
                                    enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),cantidad,nroCelda)
                            else: 
                                if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                                else:
                                    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                                    shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP'
    except: pass
############################################################## BUSCA OPERACIONES ###############################################
def buscoOperaciones(inicio,fin):
    for valor in shtTest.range('P'+str(inicio)+':'+'V'+str(fin)).value:

        cantidad = shtTest.range('Y'+str(int(valor[0]+1))).value
        if cantidad == None: cantidad = 1

        if not shtTest.range('W1').value: # Activa TRAILING  ///////////////////////////////////////////////////////////////////
            if not valor[6]: pass
            else:
                try: 
                    if valor[6] > 0: trailingStop('A'+str((int(valor[0])+1)),cantidad,int(valor[0]))
                except: pass

        if valor[1]: # # Columna Q en el excel /////////////////////////////////////////////////////////////////////////////////
            if str(valor[1]).lower() == 'c': cancelaCompra(valor[0])
            elif str(valor[1]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[1] == '+': enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
            elif str(valor[1]).upper() == 'P': getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0]) # Compra Bid
                except: shtTest.range('Q'+str(int(valor[0]+1))).value = ''
            shtTest.range('Q'+str(int(valor[0]+1))).value = ''

        if valor[2]: #  Columna R en el excel //////////////////////////////////////////////////////////////////////////////////
            if str(valor[2]).lower() == 'c': cancelaCompra(valor[0])
            elif str(valor[2]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[2] == '+': enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidad,valor[0])
            elif str(valor[2]).upper() == 'P': getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0]) # Compra Ask
                except: shtTest.range('R'+str(int(valor[0]+1))).value = ''
            shtTest.range('R'+str(int(valor[0]+1))).value = ''

        if valor[3]: # Columna S en el excel ///////////////////////////////////////////////////////////////////////////////////
            if str(valor[3]).lower() == 'v': cancelarVenta(valor[0])
            elif str(valor[3]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[3] == '-': enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidad,valor[0])
            elif str(valor[3]).upper() == 'P': getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0]) # Vendo Bid
                except: shtTest.range('S'+str(int(valor[0]+1))).value = ''
            shtTest.range('S'+str(int(valor[0]+1))).value = ''

        if valor[4]: # Columna T en el excel //////////////////////////////////////////////////////////////////////////////////
            if str(valor[4]).lower() == 'v': cancelarVenta(valor[0])
            elif str(valor[4]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[4] == '-': enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidad,valor[0])
            elif str(valor[4]).upper() == 'P': getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0]) # Vendo Ask
                except: shtTest.range('T'+str(int(valor[0]+1))).value = ''
            shtTest.range('T'+str(int(valor[0]+1))).value = ''
############################################################ CARGA BUCLE EN EXCEL ##############################################
while True:

    if time.strftime("%H:%M:%S") > '17:01:00': 
        if time.strftime("%H:%M:%S") > '17:10:00': pass
        else:
            try: getPortfolio(hb, os.environ.get('account_id'))
            except: pass
            break
        
    if not shtTest.range('Y1').value: buscoOperaciones(2,25)
    else: buscoOperaciones(rangoDesde,rangoHasta)

    time.sleep(2)

    
try: 
    hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
    hb.online.disconnect()
except: pass
print(time.strftime("%H:%M:%S"), 'Mercado cerrado. ')
shtTest.range('Q1').value = 'BONOS'
shtTest.range('S1').value = 'OPCIONES'
shtTest.range('W1').value = 'TRAILING'
shtTest.range('X1').value = 'STOP'
shtTest.range('Y1').value = 'ROLLER'

#[ ]><   \n


'''
{'Success': True, 'Error': {'Codigo': 0, 'Descripcion': None}, 
'Result': {'Totales': 
{'TotalPosicion': '110221.19', 
'Detalle': [
        {'DETA': 'Tenencia Opciones', 'IMPO': '-18203.9', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}, 
        {'DETA': 'Cuenta Corriente $', 'IMPO': '128425.09', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}]}, 
'Activos': [{'GTOS': '0', 'IMPO': '128425.09', 'ESPE': 'Subtotal Cuenta Corriente', 'TIPO': '11', 'Hora': '', 
'Subtotal': [
{'IMPO': '128425.09', 'ESPE': '', 
'APERTURA': [
{'DETA': 'Vencido', 'IMPO': '7185.22', 'GTIA': None, 'ACUM': '7185.22'}, 
{'DETA': '24 Hs. 14/08/24', 'IMPO': None, 'GTIA': None, 'ACUM': '7185.22'}, 
{'DETA': '48 Hs. 15/08/24', 'IMPO': None, 'GTIA': None, 'ACUM': '7185.22'}, 
{'DETA': '72 Hs. 16/08/24', 'IMPO': None, 'GTIA': None, 'ACUM': '7185.22'}, 
{'DETA': '+ de 72 Hs.', 'IMPO': None, 'GTIA': None, 'ACUM': '7185.22'}, 
{'DETA': 'Gtia.Opciones', 'IMPO': '121239.87', 'GTIA': None, 'ACUM': '128425.09'}], 'Detalle': [
{'DETA': 'Disponible', 'IMPO': '7185.22', 'CANT': None, 'PCIO': '1'}, 
{'DETA': 'Garantia', 'IMPO': '121239.87', 'CANT': None, 'PCIO': '1'}], 'TESP': '0', 'NERE': 'Pesos', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '11', 'Hora': 'Pesos', 'AMPL': '', 'DIVI': '100', 'TICK': 'Pesos', 'CANT': None, 'PCIO': '1', 'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}], 'CANT': None, 'TCAM': '1', 'CAN2': '116.5157897'}, {'GTOS': '-3696.89316', 'IMPO': '-18203.9', 'ESPE': 'Subtotal Opciones', 'TIPO': '10', 'Hora': '', 'Subtotal': [{'IMPO': '4', 'ESPE': '8308B', 'TESP': '4', 'NERE': 'GFGV26973G', 'GTOS': '-9.01', 'DETA': '', 'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 2,697.300 AGOSTO', 'DIVI': '100', 'TICK': 'GFGV26973G', 'CANT': '1', 'PCIO': '.04', 'CAN3': '-69.2544197', 'CAN2': '.0036291', 'CAN0': '.1301'}, {'IMPO': '198', 'ESPE': '8279B', 'TESP': '4', 'NERE': 'GFGV29581G', 'GTOS': '-2411.94294', 'DETA': '', 'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 2,958.100 AGOSTO', 'DIVI': '100', 'TICK': 'GFGV29581G', 'CANT': '22', 'PCIO': '.09', 'CAN3': '-92.4136272', 'CAN2': '.1796388', 'CAN0': '1.1863377'}, {'IMPO': '1080', 'ESPE': '8364B', 'TESP': '4', 'NERE': 'GFGV29581O', 'GTOS': '431.21', 'DETA': '', 'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 2,958.100 OCTUBRE', 'DIVI': '100', 'TICK': 'GFGV29581O', 'CANT': '4', 'PCIO': '2.7', 'CAN3': '66.4637248', 'CAN2': '.9798479', 'CAN0': '1.621975'}, {'IMPO': '3890.1', 'ESPE': '8359B', 'TESP': '4', 'NERE': 'GFGV32581O', 'GTOS': '45.45', 'DETA': '', 'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 3,258.100 OCTUBRE', 'DIVI': '100', 'TICK': 'GFGV32581O', 'CANT': '3', 'PCIO': '12.967', 'CAN3': '1.1821622', 'CAN2': '3.5293576', 'CAN0': '12.8155'}, {'IMPO': '167.4', 'ESPE': '8285B', 'TESP': '4', 'NERE': 'GFGV34081G', 'GTOS': '-52.81266', 'DETA': '', 'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 3,408.100 AGOSTO', 'DIVI': '100', 'TICK': 'GFGV34081G', 'CANT': '2', 'PCIO': '.837', 'CAN3': '-23.9825721', 'CAN2': '.1518764', 'CAN0': '1.1010633'}, {'IMPO': '-16841.2', 'ESPE': '8351B', 'TESP': '4', 'NERE': 'GFGV35581O', 'GTOS': '-1659.59', 'DETA': '', 'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 3,558.100 OCTUBRE', 'DIVI': '100', 'TICK': 'GFGV35581O', 'CANT': '-4', 'PCIO': '42.103', 'CAN3': '10.931581', 'CAN2': '-15.2794576', 'CAN0': '37.954025'}, {'IMPO': '-6702.2', 'ESPE': '8239B', 'TESP': '4', 'NERE': 'GFGV37081G', 'GTOS': '-40.19756', 'DETA': '', 'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 3,708.100 AGOSTO', 'DIVI': '100', 'TICK': 'GFGV37081G', 'CANT': '-23', 'PCIO': '2.914', 'CAN3': '.6033855', 'CAN2': '-6.0806819', 'CAN0': '2.8965228'}], 'CANT': None, 'TCAM': '1', 'CAN2': '-16.5157897'}]}}


[('Tenencia Opciones', '-18203.9'), ('Cuenta Corriente $', '128425.09')]

'''