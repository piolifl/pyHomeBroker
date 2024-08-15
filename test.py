from pyhomebroker import HomeBroker
import datetime
import pandas as pd
import requests

codigo_broker = 265 # cocos capital

dni_cuenta = 26968339 # tu dni
user_cuenta = 'bordame' # tu nombre de usuario
user_password = 'Bordame02' # tu contraseña
comitente = '49765' # tu comitente


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

    for i in portfolio['Result']['Activos'][0]['Subtotal'][0]['APERTURA']:
        if i['IMPO'] != None: 
            print(i['DETA'],i['IMPO'])



get_current_portfolio(hb, comitente)

'''
{
'Success': True, 
'Error': 
    {'Codigo': 0, 'Descripcion': None}, 
'Result': 
    {'Totales': 
        {
        'TotalPosicion': '158254.06', 
        'Detalle': 
                [
                {'DETA': 'Tenencia Opciones', 'IMPO': '-8713.9', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}, 
                {'DETA': 'Cuenta Corriente $', 'IMPO': '166967.96', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}
                ]
        }, 
    'Activos': 
        [
        {'GTOS': '0', 'IMPO': '166967.96', 'ESPE': 'Subtotal Cuenta Corriente', 'TIPO': '11', 'Hora': '', 
        'Subtotal': 
                [
                {'IMPO': '166967.96', 'ESPE': '', 
                'APERTURA': 
                    [
                    {'DETA': 'Vencido', 'IMPO': '243.62', 'GTIA': None, 'ACUM': '243.62'}, 
                    {'DETA': '24 Hs. 16/08/24', 'IMPO': None, 'GTIA': None, 'ACUM': '243.62'}, 
                    {'DETA': '48 Hs. 19/08/24', 'IMPO': None, 'GTIA': None, 'ACUM': '243.62'}, 
                    {'DETA': '72 Hs. 20/08/24', 'IMPO': None, 'GTIA': None, 'ACUM': '243.62'}, 
                    {'DETA': '+ de 72 Hs.', 'IMPO': None, 'GTIA': None, 'ACUM': '243.62'}, 
                    {'DETA': 'Gtia.Opciones', 'IMPO': '166724.34', 'GTIA': None, 'ACUM': '166967.96'}
                    ], 
                'Detalle': 
                    [
                    {'DETA': 'Disponible', 'IMPO': '243.62', 'CANT': None, 'PCIO': '1'}, 
                    {'DETA': 'Garantia', 'IMPO': '166724.34', 'CANT': None, 'PCIO': '1'}
                    ], 
                'TESP': '0', 'NERE': 'Pesos', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '11', 'Hora': 'Pesos', 'AMPL': '', 
                'DIVI': '100', 'TICK': 'Pesos', 'CANT': None, 'PCIO': '1', 'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}
                ], 'CANT': None, 'TCAM': '1', 'CAN2': '105.5062726'
                }, 

        {'GTOS': '48997.89001', 'IMPO': '-8713.9', 'ESPE': 'Subtotal Opciones', 'TIPO': '10', 'Hora': '', 
        'Subtotal': 
                [
                {'IMPO': '83', 'ESPE': '8237B', 'TESP': '4', 'NERE': 'GFGV35581G', 'GTOS': '52.96', 'DETA': '', 
                'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 3,558.100 AGOSTO', 'DIVI': '100', 'TICK': 'GFGV35581G', 
                'CANT': '10', 'PCIO': '.083', 'CAN3': '176.298269', 'CAN2': '.0524473', 'CAN0': '.03004'}, 

                {'IMPO': '9.9', 'ESPE': '8239B', 'TESP': '4', 'NERE': 'GFGV37081G', 'GTOS': '2.39001', 'DETA': '', 
                'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 3,708.100 AGOSTO', 'DIVI': '100', 'TICK': 'GFGV37081G', 
                'CANT': '3', 'PCIO': '.033', 'CAN3': '31.8244099', 'CAN2': '.0062558', 'CAN0': '.0250333'}, 
                
                {'IMPO': '1130', 'ESPE': '8233B', 'TESP': '4', 'NERE': 'GFGV38581G', 'GTOS': '-8681.81', 'DETA': '', 
                'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 3,858.100 AGOSTO', 'DIVI': '100', 'TICK': 'GFGV38581G', 
                'CANT': '100', 'PCIO': '.113', 'CAN3': '-88.4832666', 'CAN2': '.7140417', 'CAN0': '.981181'}, 
                
                {'IMPO': '-5896', 'ESPE': '8235B', 'TESP': '4', 'NERE': 'GFGV40581G', 'GTOS': '50498.68', 'DETA': '', 
                'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 4,058.100 AGOSTO', 'DIVI': '100', 'TICK': 'GFGV40581G', 
                'CANT': '-80', 'PCIO': '.737', 'CAN3': '-89.5451131', 'CAN2': '-3.7256548', 'CAN0': '7.049335'}, 
                
                {'IMPO': '-4040.8', 'ESPE': '8247B', 'TESP': '4', 'NERE': 'GFGV42581G', 'GTOS': '7125.67', 'DETA': '', 
                'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 4,258.100 AGOSTO', 'DIVI': '100', 'TICK': 'GFGV42581G', 
                'CANT': '-2', 'PCIO': '20.204', 'CAN3': '-63.813094', 'CAN2': '-2.5533626', 'CAN0': '55.83235'}
                ], 
                
        'CANT': None, 'TCAM': '1', 'CAN2': '-5.5062726'}]}}

'''


  
#[ ]><   \n
#print("\nimprimir en linea nueva")