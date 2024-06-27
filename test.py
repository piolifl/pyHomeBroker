from pyhomebroker import HomeBroker
import datetime
import pandas as pd
import requests

codigo_broker = 265 # cocos capital

dni_cuenta = 26386662 # tu dni
user_cuenta = 'piolifl' # tu nombre de usuario
user_password = 'Piolifl01' # tu contraseña
comitente = '10214' # tu comitente

## homebroker 265 es cocos capital
hb = HomeBroker(codigo_broker)

## log in: aca usar las credenciales propias
hb.auth.login(dni=dni_cuenta, user=user_cuenta, password=user_password, raise_exception=True)
hb.online.connect()

def get_current_portfolio(hb, comitente):
    
    '''Esta funcion hace un request contra /Consultas/GetConsultas al proceso 22. Esto te devuelve tu comitente'''
    
    payload = {'comitente': str(comitente),
     'consolida': '0',
     'proceso': '22',
     'fechaDesde': None,
     'fechaHasta': None,
     'tipo': None,
     'especie': None,
     'comitenteMana': None}
    
    portfolio = requests.post("https://cocoscap.com/Consultas/GetConsulta", cookies=hb.auth.cookies, json=payload).json()


    return portfolio

'''
{'Success': True, 'Error': 
{'Codigo': 0, 'Descripcion': None}, 'Result': 
{'Totales': 
{'TotalPosicion': '320493.79', 'Detalle': [
{'DETA': 'Tenencia Disponible', 'IMPO': '1471.1', 'TIPO': '1', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}, 
{'DETA': 'Tenencia Opciones', 'IMPO': '33920', 'TIPO': '1', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}, 
{'DETA': 'Cuenta Corriente $', 'IMPO': '285102.69', 'TIPO': '1', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}]}, 
'Activos': [
{'GTOS': '0', 'IMPO': '285102.69', 'ESPE': 'Subtotal Cuenta Corriente', 'TIPO': '11', 'Hora': '', 
'Subtotal': [
{'IMPO': '285102.69', 'ESPE': '', 
'APERTURA': [
{'DETA': 'Vencido', 'IMPO': '64715.81', 'GTIA': None, 'ACUM': '64715.81'}, 
{'DETA': '24 Hs. 27/06/24', 'IMPO': '220386.88', 'GTIA': None, 'ACUM': '285102.69'}, 
{'DETA': '48 Hs. 28/06/24', 'IMPO': None, 'GTIA': None, 'ACUM': '285102.69'}, 
{'DETA': '72 Hs. 01/07/24', 'IMPO': None, 'GTIA': None, 'ACUM': '285102.69'}, 
{'DETA': '+ de 72 Hs.', 'IMPO': None, 'GTIA': None, 'ACUM': '285102.69'}, 
{'DETA': 'Gtia.Opciones', 'IMPO': None, 'GTIA': None, 'ACUM': '285102.69'}], 
'Detalle': [
{'DETA': 'Disponible', 'IMPO': '64715.81', 'CANT': None, 'PCIO': '1'}, 
{'DETA': 'A Liq', 'IMPO': '220386.88', 'CANT': None, 'PCIO': '1'}], 
'TESP': '0', 'NERE': 'Pesos', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '11', 
'Hora': 'Pesos', 'AMPL': '', 'DIVI': '100', 'TICK': 'Pesos', 'CANT': None, 'PCIO': '1', 'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}], '
CANT': None, 'TCAM': '1', 'CAN2': '88.9573211'}, {'GTOS': '-911.53', 'IMPO': '33920', 'ESPE': 'Subtotal Opciones', 'TIPO': '10', 
'Hora': '', 'Subtotal': [{'IMPO': '33920', 'ESPE': '8234B', 'TESP': '4', 'NERE': 'GFGC49009G', 'GTOS': '-911.53', 'DETA': '', 
'TIPO': '10', 'Hora': '15:05:55', 'AMPL': 'GFG(C) 4,900.900 AGOSTO', 'DIVI': '100', 'TICK': 'GFGC49009G', 'CANT': '2', 
'PCIO': '169.6', 'CAN3': '-2.616968', 'CAN2': '10.5836684', 'CAN0': '174.15765'}], 'CANT': None, 'TCAM': '1', 
'CAN2': '10.5836684'}, {'GTOS': '389.0202141', 'IMPO': '1471.1', 'ESPE': 'Subtotal Titulos Publicos DOLAR MEP', 'TIPO': '1', 
'Hora': '', 'Subtotal': [{'IMPO': '724', 'ESPE': '05921', 'TESP': '1', 'NERE': 'AL30', 'GTOS': '381.9948333', 'DETA': '', 
'TIPO': '1', 'Hora': 'ANTERIOR', 'AMPL': 'BONO REP. ARGENTINA USD STEP UP 2030', 'DIVI': '.01', 'TICK': 'AL30', 'CANT': '1', 
'PCIO': '72400', 'CAN3': '111.6927083', 'CAN2': '.2259014', 'CAN0': '34200.51667'}, {'IMPO': '747.1', 'ESPE': '81086', 
'TESP': '1', 'NERE': 'GD30', 'GTOS': '7.0253808', 'DETA': '', 'TIPO': '1', 'Hora': 'ANTERIOR', 
'AMPL': 'Bonos Globales de la Rep. Arg. 2030', 'DIVI': '.01', 'TICK': 'GD30', 'CANT': '1', 'PCIO': '74710', 'CAN3': '.9492801', 
'CAN2': '.233109', 'CAN0': '74007.46192'}], 'CANT': None, 'TCAM': '1', 'CAN2': '.4590105'}]}}
[('Tenencia Disponible', '1471.1'), ('Tenencia Opciones', '33920'), ('Cuenta Corriente $', '285102.69')]
'''


  
#[ ]><   \n
#print("\nimprimir en linea nueva")