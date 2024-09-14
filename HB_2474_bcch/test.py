from pyhomebroker import HomeBroker  
import os
import environ
import requests
import time

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

print(time.strftime("%A"))
if time.strftime("%A") == 'Saturday': print('Hoy es sabado')

login()

print('Logueado en BCCH')

getPortfolio(hb, os.environ.get('account_id'))

{
    'Success': True, 'Error': {'Codigo': 0, 'Descripcion': None}, 
    'Result': {'Totales': {'TotalPosicion': '73098.49', 'Detalle': [
                                {'DETA': 'Tenencia Opciones', 'IMPO': '-3418.3', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}, 
                                {'DETA': 'Cuenta Corriente $', 'IMPO': '76516.79', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}]}, 
                'Activos': [
                    {'GTOS': '0', 'IMPO': '76516.79', 'ESPE': 'Subtotal Cuenta Corriente', 'TIPO': '11', 'Hora': '', 
                     'Subtotal': [
                         {'IMPO': '76516.79', 'ESPE': '', 'APERTURA': [
                             {'DETA': 'Vencido', 'IMPO': '-21000', 'GTIA': None, 'ACUM': '-21000'}, 
                             {'DETA': '24 Hs. 16/09/24', 'IMPO': '79516.79', 'GTIA': None, 'ACUM': '58516.79'}, 
                             {'DETA': '48 Hs. 17/09/24', 'IMPO': None, 'GTIA': None, 'ACUM': '58516.79'}, 
                             {'DETA': '72 Hs. 18/09/24', 'IMPO': None, 'GTIA': None, 'ACUM': '58516.79'}, 
                             {'DETA': '+ de 72 Hs.', 'IMPO': None, 'GTIA': None, 'ACUM': '58516.79'}, 
                             {'DETA': 'Gtia.Opciones', 'IMPO': '18000', 'GTIA': None, 'ACUM': '76516.79'}], 
                        'Detalle': [
                            {'DETA': 'Disponible', 'IMPO': '-21000', 'CANT': None, 'PCIO': '1'}, 
                            {'DETA': 'A Liq', 'IMPO': '79516.79', 'CANT': None, 'PCIO': '1'}, 
                            {'DETA': 'Garantia', 'IMPO': '18000', 'CANT': None, 'PCIO': '1'}], 
                            'TESP': '0', 'NERE': 'Pesos', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '11', 'Hora': 'Pesos', 'AMPL': '', 'DIVI': '1', 
                            'TICK': 'Pesos', 'CANT': None, 'PCIO': '1', 'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}], 
                    'CANT': None, 'TCAM': '1', 'CAN2': '104.6762936'}, 
                    {'GTOS': '20.54462', 'IMPO': '-3418.3', 'ESPE': 'Subtotal Opciones', 'TIPO': '10', 'Hora': '', 
                     'Subtotal': [
                        {'IMPO': '3770', 'ESPE': '8100F', 'TESP': '4', 'NERE': 'GFGV29581O', 'GTOS': '220.282', 'DETA': '', 'TIPO': '10', 'Hora': 'ANTERIOR', 
                          'AMPL': 'GFG(V) 2,958.100 OCTUBRE', 'DIVI': '100', 'TICK': 'GFGV29581O', 'CANT': '100', 'PCIO': '.377', 'CAN3': '6.2056197', 'CAN2': '5.1574253', 'CAN0': '.3549718'}, 
                        {'IMPO': None, 'ESPE': '8056F', 
                         'Detalle': [
                             {'DETA': 'Disponible', 'IMPO': '7246.2', 'CANT': '78', 'PCIO': '.929'}, 
                             {'DETA': 'A Liq.', 'IMPO': '-7246.2', 'CANT': '-78', 'PCIO': '.929'}], 
                        'TESP': '4', 'NERE': 'GFGV32581O', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '10', 'Hora': 'ANTERIOR', 
                        'AMPL': 'GFG(V) 3,258.100 OCTUBRE', 'DIVI': '100', 'TICK': 'GFGV32581O', 'CANT': None, 'PCIO': '.929', 'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}, 
                        {'IMPO': '-6150.9', 'ESPE': '8048F', 'TESP': '4', 'NERE': 'GFGV35581O', 'GTOS': '-275.83988', 'DETA': 'A Liq', 'TIPO': '10', 'Hora': 'ANTERIOR', 
                         'AMPL': 'GFG(V) 3,558.100 OCTUBRE', 'DIVI': '100', 'TICK': 'GFGV35581O', 'CANT': '-29', 'PCIO': '2.121', 'CAN3': '4.6950988', 'CAN2': '-8.4145377', 'CAN0': '2.0258828'}, 
                         {'IMPO': '-1037.4', 'ESPE': '8054F', 
                          'Detalle': [
                              {'DETA': 'Disponible', 'IMPO': '-10374', 'CANT': '-30', 'PCIO': '3.458'}, 
                              {'DETA': 'A Liq.', 'IMPO': '9336.6', 'CANT': '27', 'PCIO': '3.458'}], 
                        'TESP': '4', 'NERE': 'GFGV38581O', 'GTOS': '76.1025', 'DETA': 'Total', 'TIPO': '10', 'Hora': 'ANTERIOR', 
                        'AMPL': 'GFG(V) 3,858.100 OCTUBRE', 'DIVI': '100', 'TICK': 'GFGV38581O', 'CANT': '-3', 'PCIO': '3.458', 'CAN3': '-6.8345154', 'CAN2': '-1.4191812', 'CAN0': '3.711675'}], 
                        
                        'CANT': None, 'TCAM': '1', 'CAN2': '-4.6762936'}]}}
