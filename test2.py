portfolio = {
    'Success': True, 'Error': {'Codigo': 0, 'Descripcion': None}, 
    'Result': {'Totales': {'TotalPosicion': '87566.99',  'Detalle': [{'DETA': 'Tenencia a Liquidar','IMPO': '1893.61','TIPO': '10','Hora': 'Pesos','CANT': None,'TCAM': '1'}, 
    {'DETA': 'Cuenta Corriente $','IMPO': '85673.38','TIPO': '10','Hora': 'Pesos','CANT': None,'TCAM': '1'}] },
    'Activos': [

            {'GTOS': '0', 'IMPO': '85673.38', 'ESPE': 'Subtotal Cuenta Corriente','TIPO': '11', 'Hora': '', 
             
             'Subtotal': [ { 'IMPO': '85673.38', 'ESPE': '', 
                            'APERTURA':[{'DETA': 'Vencido', 'IMPO': '52297.43', 'GTIA': None,'ACUM': '52297.43'},  
                                        {'DETA': '24 Hs. 06/05/24', 'IMPO': '33375.95', 'GTIA': None, 'ACUM': '85673.38'}, 
                                        {'DETA': '48 Hs. 07/05/24', 'IMPO': None, 'GTIA': None, 'ACUM': '85673.38'}, 
                                        {'DETA': '72 Hs. 08/05/24','IMPO': None, 'GTIA': None, 'ACUM': '85673.38'}, 
                                        {'DETA': '+ de 72 Hs.', 'IMPO': None, 'GTIA': None,'ACUM': '85673.38'}, 
                                        {'DETA': 'Gtia.Opciones', 'IMPO': None, 'GTIA': None, 'ACUM': '85673.38'} ], 
                            'Detalle': [{'DETA': 'Disponible', 'IMPO': '52297.43', 'CANT': None, 'PCIO': '1'},  
                                        {'DETA': 'A Liq', 'IMPO': '33375.95', 'CANT': None, 'PCIO': '1'} ], 
                            'TESP': '0', 'NERE': 'Pesos', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '11', 'Hora': 'Pesos', 
                            'AMPL': '', 'DIVI': '100', 'TICK': 'Pesos', 'CANT': None, 'PCIO': '1','CAN3': '0', 'CAN2': '0', 
                            'CAN0': '0' } ], 'CANT': None, 'TCAM': '1', 'CAN2': '97.8375299' }, 

            {'GTOS': '0','IMPO': '1893.61', 'ESPE': 'Subtotal Letras', 'TIPO': '6', 'Hora': '',  
            'Subtotal': [{'IMPO': '1893.61','ESPE': '09239','TESP': '1','NERE': 'X20Y4','GTOS': '0','DETA': 'A Liq','TIPO': '6',  
                          'Hora': '14:02:33','AMPL': 'LT REP ARG AJ CER A DESC V20/05/24','DIVI': '.01', 'TICK': 'X20Y4',
                          'CANT': '1000',  'PCIO': '189.361','CAN3': '0','CAN2': '0','CAN0': '0'} ], 
            'CANT': None, 'TCAM': '1', 'CAN2': '2.1624701'}, 


    
            {'GTOS': '0', 'IMPO': None, 'ESPE': 'Subtotal Opciones', 'TIPO': '10', 'Hora': '',   
             'Subtotal': [{'IMPO': None, 'ESPE': '8118B', 
                           'Detalle': [ {'DETA': 'Disponible', 'IMPO': '32080', 'CANT': '4', 'PCIO': '80.2'},  
                                       {'DETA': 'A Liq.', 'IMPO': '-32080', 'CANT': '-4', 'PCIO': '80.2'} ], 
                            'TESP': '4', 'NERE': 'GFGC4200JU', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '10', 'Hora': '14:02:21',  'AMPL': 'GFG(C) 4200.000 JUNIO', 'DIVI': '100', 'TICK': 'GFGC4200JU', 'CANT': None, 'PCIO': '80.2',  'CAN3': '-100.7063696', 'CAN2': '0', 'CAN0': '-11353.83'}], 
            'CANT': None, 'TCAM': '1', 'CAN2': '0'}
                ] 
                }
                }

#portfolio = portfolio["Result"]["Activos"][0:]
#portfolio = [ (x['NERE']) for x in portfolio[1]['Subtotal'][0]]
#portfolio = portfolio["Result"]["Activos"][1]["Subtotal"]

'''activos = [(i) for x,i in portfolio['Result'].items() if x == 'Activos']
subtotal = [(x) for x in activos[0][0]]

print(type(subtotal),subtotal)'''

for valor in portfolio["Result"]["Activos"][0:]:
    tipo = valor['ESPE']
    for i in valor['Subtotal']:
        if tipo == 'Subtotal Cuenta Corriente': 
            detalle = [k for j,k in i.items() if j == 'Detalle']
            for h in detalle:
                print(h)
        elif i['IMPO'] != None: print(tipo,i['NERE'],'/ cantidad:',i['CANT'],'/ importe:',i['IMPO'],'/',i['Hora'])


  
#[ ]><   \n
