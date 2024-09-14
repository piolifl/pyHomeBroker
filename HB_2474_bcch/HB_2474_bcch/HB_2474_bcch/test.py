from pyhomebroker import HomeBroker  
import os
import environ
import requests

env = environ.Env()
environ.Env.read_env()

def getPortfolio(hb, comitente):
    try:
        payload = {'comitente': str(comitente),
        'consolida': '0',
        'proceso': '22',
        'fechaDesde': None,
        'fechaHasta': None,
        'tipo': None,
        'especie': None,
        'comitenteMana': None}

        portfolio = requests.post("https://clientes.bcch.org.ar/Consultas/GetConsulta", cookies=hb.auth.cookies, json=payload).json()

        print(portfolio)

    except: pass

def login():
    hb.auth.login(dni=str(os.environ.get('dni')), 
    user=str(os.environ.get('user')),  
    password=str(os.environ.get('password')),
    raise_exception=True)

hb = HomeBroker(int(os.environ.get('broker')))


login()

print('Logueado en BCCH')

getPortfolio(hb, os.environ.get('account_id'))

'''
 'Subtotal': [
 {'IMPO': '16947.88807', 'ESPE': '', 
 'APERTURA': [
 {'DETA': 'Vencido', 'IMPO': '16947.88807', 'GTIA': None, 'ACUM': '16947.88807'}, 
 {'DETA': '24 Hs. 09/09/24', 'IMPO': None, 'GTIA': None, 'ACUM': '16947.88807'}, 
 {'DETA': '48 Hs. 10/09/24', 'IMPO': None, 'GTIA': None, 'ACUM': '16947.88807'}, 
 {'DETA': '72 Hs. 11/09/24', 'IMPO': None, 'GTIA': None, 'ACUM': '16947.88807'}, 
 {'DETA': '+ de 72 Hs.', 'IMPO': None, 'GTIA': None, 'ACUM': '16947.88807'}, 
 {'DETA': 'Gtia.Opciones', 'IMPO': None, 'GTIA': None, 'ACUM': '16947.88807'}], 
 
 'TESP': '0', 'NERE': 'Pesos', 'GTOS': '0', 'DETA': '', 'TIPO': '11', 'Hora': 'Pesos', 'AMPL': '', 'DIVI': '100', 'TICK': 'Pesos', 'CANT': None, 'PCIO': '1', 'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}], 
 'CANT': None, 'TCAM': '1', 'CAN2': '107.470614'}, 
 
 {'GTOS': '-42.1902', 'IMPO': '-1178.1', 'ESPE': 'Subtotal Opciones', 'TIPO': '10', 'Hora': '', 
 'Subtotal': [
 {'IMPO': '104.8', 'ESPE': '8100F', 'TESP': '4', 
 'NERE': 'GFGV29581O', 'GTOS': '-.6302', 'DETA': '', 'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 2,958.100 OCTUBRE', 'DIVI': '100', 'TICK': 'GFGV29581O', 'CANT': '1', 'PCIO': '1.048', 'CAN3': '-.5977414', 'CAN2': '.6645619', 'CAN0': '1.054302'}, 
 
 {'IMPO': '-1282.9', 'ESPE': '8054F', 'TESP': '4', 
 'NERE': 'GFGV38581O', 'GTOS': '-41.56', 'DETA': '', 'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 3,858.100 OCTUBRE', 'DIVI': '100', 'TICK': 'GFGV38581O', 'CANT': '-1', 'PCIO': '12.829', 'CAN3': '3.3479949', 'CAN2': '-8.1351759', 'CAN0': '12.4134'}], 
 
 'CANT': None, 'TCAM': '1', 'CAN2': '-7.470614'}]}}
'''