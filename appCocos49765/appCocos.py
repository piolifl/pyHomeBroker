import pycocos
import xlwings as xw   
import time

wb = xw.Book('.\\epgb_pyHB.xlsb')
shtTest = wb.sheets('HomeBroker')

shtTest.range('W1').value = 'TRAILING'
shtTest.range('X1').value = 'STOP'
shtTest.range('Z1').value = 0.001
rangoDesde = '26'
rangoHasta = '89'
hoyEs = 'Monday' # time.strftime("%A")

def diaLaboral():
    if hoyEs == 'Saturday' or hoyEs == 'Sunday':
        return 'Fin de semana'

if diaLaboral():
    print('Es FIN DE SEMANA, sin logueo en comitente. Planilla: activa || Buscando operaciones: activo')
    esFinde = True
else: 
    esFinde = False
    app = pycocos.Cocos("miryam.borda79@gmail.com","Bordame1+")
    print(time.strftime("%H:%M:%S"),"Logueo correcto en COCOS 49765")

def cancelaCompra(celda):
    orderC = shtTest.range('AB'+str(int(celda+1))).value
    if not orderC or orderC == None or orderC == 'None' or orderC == '': orderC = 0
    if esFinde == False: 
        try: 
            #hb.orders.cancel_order(int(os.environ.get('account_id')),int(orderC))
            print(f"/// Cancelada Compra : {int(orderC)} ",end='\t')
        except: pass
    try: shtTest.range('V'+str(int(celda+1))).value -= shtTest.range('AC'+str(int(celda+1))).value
    except: pass
    shtTest.range('AB'+str(int(celda+1))+':'+'AD'+str(int(celda+1))).value = ''
        
def cancelarVenta(celda):
    orderV = shtTest.range('AE'+str(int(celda+1))).value
    if not orderV or orderV == None or orderV == 'None' or orderV == '': orderV = 0
    if esFinde == False: 
        try:
            #hb.orders.cancel_order(int(os.environ.get('account_id')),int(orderV))
            print(f"/// Cancelada Venta  : {int(orderV)} " ,end='\t')
        except: pass
    try: shtTest.range('V'+str(int(celda+1))).value += shtTest.range('AF'+str(int(celda+1))).value
    except: pass
    shtTest.range('AE'+str(int(celda+1))+':'+'AG'+str(int(celda+1))).value = ''

def cancelarTodo(desde,hasta):
    if esFinde == False:
        try:  
            #hb.orders.cancel_all_orders(int(os.environ.get('account_id')))
            print("/// Todas las ordenes activas canceladas ")
        except: pass
    shtTest.range('AB'+str(desde)+':'+'AH'+str(hasta)).value = ''

def cantidadAuto(nroCelda):
    cantidad = shtTest.range('Y'+str(int(nroCelda))).value
    if not cantidad or cantidad == None or cantidad == 'None': 
        cantidad = 0
    return abs(int(cantidad))

def soloContinua():
    pass

def generaNombre(ticker=str,plazo=str,moneda=str):
    # Generar long ticker T0=ci, T1=24hs PESOS, USD, CCL, OPTIONS. Solo aclarar para opciones: segment = app.segments.OPTIONS

    if (plazo).lower() == 'spot': settlement = app.settlements.T0
    else: settlement = app.settlements.T0

    if (moneda).lower() == 'usd': currency = app.currencies.USD
    elif (moneda).lower() == 'cable': currency = app.currencies.CABLE
    else: currency = app.currencies.PESOS

    long_ticker = app.long_ticker(ticker, settlement, currency)

    return long_ticker



def enviarOrden(tipo=str,symbol=str, price=float, size=int, celda=int):
    global orderC, orderV
    orderC, orderV = hoyEs,hoyEs
    symbol = str(shtTest.range(str(symbol)).value).split()
    precio = shtTest.range(str(price)).value
    if tipo.lower() == 'buy': 
        try: 
            if len(symbol) < 2:
                if esFinde == False: 
                    pass # orderC = hb.orders.send_buy_order(symbol[0],'24hs', float(precio), abs(int(size)))
                    
                shtTest.range('AD'+str(int(celda+1))).value = float(precio)
                shtTest.range('X'+str(int(celda+1))).value = precio
                print(f'        ______/ BUY  opcion + {int(size)} {symbol[0]} // precio: {precio} // {orderC}') 
            else:
                if esFinde == False: 
                    pass #orderC = hb.orders.send_buy_order(symbol[0],symbol[2], float(precio), abs(int(size)))
                shtTest.range('AD'+str(int(celda+1))).value = float(precio/100)
                shtTest.range('X'+str(int(celda+1))).value = precio / 100
                print(f'        ______/ BUY + {int(size)} {symbol[0]} {symbol[2]} // precio: {round(precio/100,4)} // {orderC}')
        except: 
            shtTest.range('Q'+str(int(celda+1))+':'+'R'+str(int(celda+1))).value = ''
            print(f'        ______/ ERROR en COMPRA. {symbol[0]} // precio: {precio} // + {int(size)}')

        shtTest.range('Q'+str(int(celda+1))+':'+'R'+str(int(celda+1))).value = ''
        try: shtTest.range('V'+str(int(celda+1))).value += abs(int(size))
        except: shtTest.range('V'+str(int(celda+1))).value = abs(int(size))
        shtTest.range('AB'+str(int(celda+1))).value = orderC
        shtTest.range('AC'+str(int(celda+1))).value = abs(int(size))
    
    else: # VENTA
        try:
            if len(symbol) < 2:
                if esFinde == False: 
                    pass #orderV = hb.orders.send_sell_order(symbol[0],'24hs', float(precio), abs(int(size)))
                shtTest.range('AG'+str(int(celda+1))).value = float(precio)
                shtTest.range('X'+str(int(celda+1))).value = precio
                print(f'______/ SELL opcion - {int(size)} {symbol[0]} // precio: {precio} // {orderV}')
            else:
                if esFinde == False: 
                    pass #orderV = hb.orders.send_sell_order(symbol[0],symbol[2], float(precio), abs(int(size)))
                shtTest.range('AG'+str(int(celda+1))).value = float(precio/100)
                shtTest.range('X'+str(int(celda+1))).value = precio /100
                print(f'______/ SELL - {int(size)} {symbol[0]} {symbol[2]} // precio: {round(precio/100,4)} // {orderV}')
        except:
            shtTest.range('S'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
            print(f'______/ ERROR en VENTA. {symbol[0]} // precio: {precio} // {int(size)}')

        shtTest.range('S'+str(int(celda+1))+':'+'T'+str(int(celda+1))).value = ''
        try: shtTest.range('V'+str(int(celda+1))).value -= abs(int(size))
        except: shtTest.range('V'+str(int(celda+1))).value = int(size) / -1
        shtTest.range('AE'+str(int(celda+1))).value = orderV
        shtTest.range('AF'+str(int(celda+1))).value = abs(int(size))

def trailingStop(nombre=str,cantidad=int,nroCelda=int,vendido=str):
    try:
        costo = shtTest.range('X'+str(int(nroCelda+1))).value 
        if not costo or costo == None or costo == 'None': soloContinua()
        nombre = str(shtTest.range(str(nombre)).value).split()
        bid = shtTest.range('C'+str(int(nroCelda+1))).value
        ask = shtTest.range('D'+str(int(nroCelda+1))).value
        last = shtTest.range('F'+str(int(nroCelda+1))).value
        if not last or last == None or last == 'None': soloContinua()

        ganancia = shtTest.range('Z1').value
        if not ganancia: ganancia = 0.0005

        if len(nombre) < 2: # Ingresa si son OPCIONES ///////////////////////////////////////////////////////////////////////////
            if vendido == 'no':
                if bid > abs(costo) * (1 + (ganancia*75)):
                    if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'BUYTRAIL': pass
                    else: shtTest.range('W'+str(int(nroCelda+1))).value = 'BUYTRAIL'
                    shtTest.range('X'+str(int(nroCelda+1))).value = bid
                if not shtTest.range('X1').value:
                    if last < abs(costo) * (1 - (ganancia*75)): 
                        if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP':
                            if bid > last * (1-(ganancia*45)):
                                if shtTest.range('Y'+str(int(nroCelda+1))).value : 
                                    try: shtTest.range('U'+str(int(nroCelda+1))).value -= abs(cantidad)
                                    except: pass
                                    
                                    enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),abs(cantidad),nroCelda)
                        else:
                            if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                            else: shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP'
            else:
                if ask < abs(costo) * (1 - (ganancia*75)):
                    if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'SELLTRAIL': pass
                    else: shtTest.range('W'+str(int(nroCelda+1))).value = 'SELLTRAIL'
                    shtTest.range('X'+str(int(nroCelda+1))).value = ask
                if not shtTest.range('X1').value:  

                    if last > abs(costo) * (1 + (ganancia*75)): 
                        if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP': 

                            if ask < last * (1-(ganancia*15)):
                                if shtTest.range('Y'+str(int(nroCelda+1))).value : 
                                    try: shtTest.range('U'+str(int(nroCelda+1))).value += abs(cantidad)
                                    except: pass
                                    
                                    enviarOrden('buy','A'+str((int(nroCelda)+1)),'D'+str((int(nroCelda)+1)),abs(cantidad),nroCelda)
                        else:
                            if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                            else: shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP'

        else: # Ingresa si son BONOS / LETRAS / ON / CEDEARS ////////////////////////////////////////////////////////////////////
            if time.strftime("%H:%M:%S") > '16:24:50' and str(nombre[2]).lower() == 'spot': 
                if time.strftime("%H:%M:%S") > '17:05:00': pass
                else: 
                    shtTest.range('W'+str(int(nroCelda+1))).value = "CLOSED"
                    pass
            if time.strftime("%H:%M:%S") > '16:56:50' and str(nombre[2]).lower() == '24hs': 
                if time.strftime("%H:%M:%S") > '17:05:00': pass
                else: 
                    shtTest.range('W'+str(int(nroCelda+1))).value = "CLOSED"
                    pass
            else:
                # Rutina, si el precio BID sube modifica precio promedio de compra //////////////////////////////////////////////
                if bid / 100 > abs(costo) * (1 + ganancia):             
                    if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'BUYTRAIL': pass
                    else: shtTest.range('W'+str(int(nroCelda+1))).value = 'BUYTRAIL'
                    shtTest.range('X'+str(int(nroCelda+1))).value = round(bid / 100,5)

                if not shtTest.range('X1').value:
                    if last / 100 < abs(costo) * (1 - ganancia):
                        if str(shtTest.range('W'+str(int(nroCelda+1))).value)=='STOP' and (bid/100)>(last/100)*(1-ganancia):
                            print(f'{time.strftime("%H:%M:%S")} STOP vendo    ',end=' || ')
                            if shtTest.range('Y'+str(int(nroCelda+1))).value : 
                                shtTest.range('U'+str(int(nroCelda+1))).value -= abs(cantidad)
                                enviarOrden('sell','A'+str((int(nroCelda)+1)),'C'+str((int(nroCelda)+1)),abs(cantidad),nroCelda)
                        else: 
                            if str(shtTest.range('W'+str(int(nroCelda+1))).value) == 'STOP': pass
                            else: shtTest.range('W'+str(int(nroCelda+1))).value = 'STOP'
    except: pass

def buscoOperaciones(inicio,fin):
    for valor in shtTest.range('P'+str(inicio)+':'+'U'+str(fin)).value:
        try:
            if not shtTest.range('W1').value: # Permite TRAILING  ///////////////////////////////////////////////////////////////
                if not valor[5]:  pass
                else: 
                    if valor[5] < 0: vendido = 'si'
                    else: vendido = 'no'
                    trailingStop('A'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),int(valor[0]),vendido)
        except: pass

        if valor[1]: # # Columna Q en el excel /////////////////////////////////////////////////////////////////////////////////
            if str(valor[1]).lower() == 'c': cancelaCompra(valor[0])
            elif str(valor[1]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[1] == '+': 
                enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
            elif str(valor[1]).upper() == 'P': 
                if esFinde == False: 
                    pass #getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: 
                    if shtTest.range('AB'+str(int(valor[0]+1))).value: cancelaCompra(valor[0]) # CANCELA oreden compra anterior
                    enviarOrden('buy','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[1],valor[0]) # Compra Bid
                except: shtTest.range('Q'+str(int(valor[0]+1))).value = ''
            shtTest.range('Q'+str(int(valor[0]+1))).value = ''

        if valor[2]: #  Columna R en el excel //////////////////////////////////////////////////////////////////////////////////
            if str(valor[2]).lower() == 'c': cancelaCompra(valor[0])
            elif str(valor[2]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[2] == '+': 
                enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
            elif str(valor[2]).upper() == 'P': 
                if esFinde == False: 
                    pass #getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: 
                    if shtTest.range('AB'+str(int(valor[0]+1))).value: cancelaCompra(valor[0])
                    enviarOrden('buy','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[2],valor[0]) # Compra Ask
                except: shtTest.range('R'+str(int(valor[0]+1))).value = ''
            shtTest.range('R'+str(int(valor[0]+1))).value = ''

        if valor[3]: # Columna S en el excel ///////////////////////////////////////////////////////////////////////////////////
            if str(valor[3]).lower() == 'v': cancelarVenta(valor[0])
            elif str(valor[3]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[3] == '-': 
                enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
            elif str(valor[3]).upper() == 'P': 
                if esFinde == False: 
                    pass #getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: 
                    if shtTest.range('AE'+str(int(valor[0]+1))).value: cancelarVenta(valor[0])
                    enviarOrden('sell','A'+str((int(valor[0])+1)),'C'+str((int(valor[0])+1)),valor[3],valor[0]) # Vendo Bid
                except: shtTest.range('S'+str(int(valor[0]+1))).value = ''
            shtTest.range('S'+str(int(valor[0]+1))).value = ''

        if valor[4]: # Columna T en el excel //////////////////////////////////////////////////////////////////////////////////
            if str(valor[4]).lower() == 'v': cancelarVenta(valor[0])
            elif str(valor[4]).lower() == 'x': cancelarTodo(inicio,fin)
            elif valor[4] == '-': 
                enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),cantidadAuto(valor[0]+1),valor[0])
            elif str(valor[4]).upper() == 'P': 
                if esFinde == False: 
                    pass #getPortfolio(hb, os.environ.get('account_id'))
            else: 
                try: 
                    if shtTest.range('AE'+str(int(valor[0]+1))).value: cancelarVenta(valor[0]) # CANCELA oreden venta anterior
                    enviarOrden('sell','A'+str((int(valor[0])+1)),'D'+str((int(valor[0])+1)),valor[4],valor[0]) # Vendo Ask
                except: shtTest.range('T'+str(int(valor[0]+1))).value = ''
            shtTest.range('T'+str(int(valor[0]+1))).value = ''

while True:

    if time.strftime("%H:%M:%S") > '17:01:00': 
        if time.strftime("%H:%M:%S") > '17:05:00': pass
        else: break
        
    buscoOperaciones(rangoDesde,rangoHasta)
    
    time.sleep(2)
    if not shtTest.range('M1').value: 
        shtTest.range('M1').value = 'volume'
        app.logout() # Cerrar sesión












