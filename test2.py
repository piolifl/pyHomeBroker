import os
portfolio = {
    'Success': True, 
    'Error': {'Codigo': 0, 'Descripcion': None}, 
    'Result': {'Totales': 
               {'TotalPosicion': '80527.73', 
                'Detalle': [{'DETA': 'Tenencia Opciones', 'IMPO': '79869.2', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}, 
                            {'DETA': 'Cuenta Corriente $', 'IMPO': '658.53', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}]}, 
                'Activos': [{'GTOS': '0', 'IMPO': '658.53', 'ESPE': 'Subtotal Cuenta Corriente', 'TIPO': '11', 'Hora': '', 
                             'Subtotal': [{'IMPO': '658.53', 'ESPE': '', 
                                           'APERTURA': [{'DETA': 'Vencido', 'IMPO': '75386.8', 'GTIA': None, 'ACUM': '75386.8'}, 
                                                        {'DETA': '24 Hs. 20/05/24', 'IMPO': '-74728.27', 'GTIA': None, 'ACUM': '658.53'}, 
                                                        {'DETA': '48 Hs. 21/05/24', 'IMPO': None, 'GTIA': None, 'ACUM': '658.53'}, 
                                                        {'DETA': '72 Hs. 22/05/24', 'IMPO': None, 'GTIA': None, 'ACUM': '658.53'}, 
                                                        {'DETA': '+ de 72 Hs.', 'IMPO': None, 'GTIA': None, 'ACUM': '658.53'}, 
                                                        {'DETA': 'Gtia.Opciones', 'IMPO': None, 'GTIA': None, 'ACUM': '658.53'}], 
                                            'Detalle': [{'DETA': 'Disponible', 'IMPO': '75386.8', 'CANT': None, 'PCIO': '1'}, 
                                                        {'DETA': 'A Liq', 'IMPO': '-74728.27', 'CANT': None, 'PCIO': '1'}], 
                                            'TESP': '0', 'NERE': 'Pesos', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '11', 'Hora': 'Pesos', 
                                            'AMPL': '', 'DIVI': '100', 'TICK': 'Pesos', 'CANT': None, 'PCIO': '1', 'CAN3': '0', 'CAN2': '0', 
                                            'CAN0': '0'}], 'CANT': None, 'TCAM': '1', 'CAN2': '.817768'}, 
                                            {'GTOS': '1464.34', 'IMPO': '79869.2', 'ESPE': 'Subtotal Opciones', 'TIPO': '10', 'Hora': '', 
                                             'Subtotal': [{'IMPO': '79869.2', 'ESPE': '8118B', 'TESP': '4', 'NERE': 'GFGC41559J', 
                                                           'GTOS': '1464.34', 'DETA': 'A Liq', 'TIPO': '10', 'Hora': 'ANTERIOR', 
                                                           'AMPL': 'GFG(C) 4,155.900 JUNIO', 'DIVI': '100', 'TICK': 'GFGC41559J', 'CANT': '4', 
                                                           'PCIO': '199.673', 'CAN3': '1.8676648', 'CAN2': '99.182232', 'CAN0': '196.01215'}, 
                                                           {'IMPO': None, 'ESPE': '8174B', 
                                                            'Detalle': [{'DETA': 'Disponible', 'IMPO': '7227.1', 'CANT': '1', 'PCIO': '72.271'}, 
                                                                        {'DETA': 'A Liq.', 'IMPO': '-7227.1', 'CANT': '-1', 'PCIO': '72.271'}], 
                                                                        'TESP': '4', 'NERE': 'GFGC45559J', 'GTOS': '0', 'DETA': 'Total', 
                                                                        'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(C) 4,555.900 JUNIO', 
                                                                        'DIVI': '100', 'TICK': 'GFGC45559J', 'CANT': None, 'PCIO': '72.271', 
                                                                        'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}, 
                                                            {'IMPO': None, 'ESPE': '8192B', 
                                                             'Detalle': [{'DETA': 'Disponible', 'IMPO': '3658.6', 'CANT': '2', 'PCIO': '18.293'},
                                                                         {'DETA': 'A Liq.', 'IMPO': '-3658.6', 'CANT': '-2', 'PCIO': '18.293'}], 
                                                                         'TESP': '4', 'NERE': 'GFGC51559J', 'GTOS': '0', 'DETA': 'Total', 
                                                                         'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(C) 5,155.900 JUNIO', 
                                                                         'DIVI': '100', 'TICK': 'GFGC51559J', 'CANT': None, 'PCIO': '18.293', 
                                                                         'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}, 
                                                            {'IMPO': None, 'ESPE': '8194B',
                                                             'Detalle': [{'DETA': 'Disponible', 'IMPO': '1237.2', 'CANT': '1', 'PCIO': '12.372'}, 
                                                                         {'DETA': 'A Liq.', 'IMPO': '-1237.2', 'CANT': '-1', 'PCIO': '12.372'}], 
                                                                         'TESP': '4', 'NERE': 'GFGC53559J', 'GTOS': '0', 'DETA': 'Total', 
                                                                         'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(C) 5,355.900 JUNIO', 
                                                                         'DIVI': '100', 'TICK': 'GFGC53559J', 'CANT': None, 'PCIO': '12.372', 
                                                                         'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}], 
                                                                         'CANT': None, 'TCAM': '1', 'CAN2': '99.182232'}]}}
#portfolio = portfolio["Result"]["Activos"][0:]
#portfolio = [ (x['NERE']) for x in portfolio[1]['Subtotal'][0]]
#portfolio = portfolio["Result"]["Activos"][1]["Subtotal"]
os.system('cls')

print()

def buscando():
    subtotal = [ i['Subtotal'] for i in portfolio["Result"]["Activos"][0:] ]
    for i in subtotal[0:]:

        if i[0]['NERE'] == 'Pesos': 
            try: subtotal = [ (x['DETA'],x['IMPO'],x['ACUM']) for x in i[0]['APERTURA'] if x['IMPO'] != None]
            except: subtotal = [ (x['DETA'],x['IMPO'],x['ACUM']) for x in i[0] if x['IMPO'] != None]
            for x in subtotal: print(x)
        else: 
            subtotal = [ (x['CANT'], x['NERE'],x['CAN0'],' || ',x['PCIO'],x['GTOS']) for x in i[0:] if x['CANT'] != None]
            for x in subtotal: print(x)
            
buscando()







  
#[ ]><   \n
#print("\nimprimir en linea nueva")