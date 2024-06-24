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




  
#[ ]><   \n
#print("\nimprimir en linea nueva")