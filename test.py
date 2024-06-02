from pyhomebroker import HomeBroker
import datetime
import pandas as pd
import requests

"""En esta notebook vamos a mostrar como armar un portfolio de manera programática (usando algun criterio) y comprarlo. Además, mostramos código para rebalancear. La estrategia que elegimos es una forma de 1/N. Es una estretegia sencilla que suele usarse como baseline. (Esto no es una recomendacion de compra ni nada por el estilo, solamente tiene fines didácticos y busca mostrar como operar usando python).
El código está hecho para ser mas claro que performante: puede optimizarse mucho, pero esta es la forma mas clara que encontré de explicarlo.
"""

#Primero me conecto a cocos capital usando pyhomebroker.


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
    portfolio = portfolio["Result"]["Activos"][1]["Subtotal"]
    
    ## esto devuelve el ticker, el precio y la cantidad que tenes
    portfolio = [( x["NERE"], float(x["PCIO"]), float(x["CANT"]) ) for x in portfolio]
    return portfolio




  
#[ ]><   \n
#print("\nimprimir en linea nueva")