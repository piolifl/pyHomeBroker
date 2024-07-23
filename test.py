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
{'Success': True, 'Error': {'Codigo': 0, 'Descripcion': None}, 
'Result': {'Totales': {'TotalPosicion': '21402.72', 
'Detalle': [
{'DETA': 'Tenencia Opciones', 'IMPO': '-900', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}, 
{'DETA': 'Cuenta Corriente $', 'IMPO': '22302.72', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}]}, 

'Activos': [{'GTOS': '0', 'IMPO': '22302.72', 'ESPE': 'Subtotal Cuenta Corriente', 'TIPO': '11', 'Hora': '', 
'Subtotal': [{'IMPO': '22302.72', 'ESPE': '', 
'APERTURA': [
{'DETA': 'Vencido', 'IMPO': '20747.81', 'GTIA': None, 'ACUM': '20747.81'}, 
{'DETA': '24 Hs. 23/07/24', 'IMPO': '1554.91', 'GTIA': None, 'ACUM': '22302.72'}, 
{'DETA': '48 Hs. 24/07/24', 'IMPO': None, 'GTIA': None, 'ACUM': '22302.72'}, 
{'DETA': '72 Hs. 25/07/24', 'IMPO': None, 'GTIA': None, 'ACUM': '22302.72'}, 
{'DETA': '+ de 72 Hs.', 'IMPO': None, 'GTIA': None, 'ACUM': '22302.72'}, 
{'DETA': 'Gtia.Opciones', 'IMPO': None, 'GTIA': None, 'ACUM': '22302.72'}], 

'Detalle': [
{'DETA': 'Disponible', 'IMPO': '20747.81', 'CANT': None, 'PCIO': '1'}, 
{'DETA': 'A Liq', 'IMPO': '1554.91', 'CANT': None, 'PCIO': '1'}], 
'TESP': '0', 'NERE': 'Pesos', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '11', 'Hora': 'Pesos', 'AMPL': '', 
'DIVI': '100', 'TICK': 'Pesos', 'CANT': None, 'PCIO': '1', 'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}], 
'CANT': None, 'TCAM': '1', 'CAN2': '104.205073'}, 
{'GTOS': '0', 'IMPO': '-900', 'ESPE': 'Subtotal Opciones', 'TIPO': '10', 'Hora': '', 
'Subtotal': [{'IMPO': '3100', 'ESPE': '8308B', 'TESP': '4', 'NERE': 'GFGV26973G', 'GTOS': '0', 'DETA': 'A Liq', 'TIPO': '10', 
'Hora': '15:22:12', 'AMPL': 'GFG(V) 2,697.300 AGOSTO', 'DIVI': '100', 'TICK': 'GFGV26973G', 'CANT': '10', 'PCIO': '3.1', 
'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}, {'IMPO': '-4000', 'ESPE': '8268B', 'TESP': '4', 'NERE': 'GFGV28081G', 'GTOS': '0', 
'DETA': 'A Liq', 'TIPO': '10', 'Hora': '15:39:47', 'AMPL': 'GFG(V) 2,808.100 AGOSTO', 'DIVI': '100', 'TICK': 'GFGV28081G', 
'CANT': '-10', 'PCIO': '4', 'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}], 'CANT': None, 'TCAM': '1', 'CAN2': '-4.205073'}]}}




'''


  
#[ ]><   \n
#print("\nimprimir en linea nueva")