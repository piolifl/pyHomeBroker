import os
portfolio = {
    'Success': True, 'Error': {'Codigo': 0, 'Descripcion': None}, 
    'Result': {'Totales': {'TotalPosicion': '118710.36', 
                           'Detalle': [
                               {'DETA': 'Tenencia Opciones', 'IMPO': '56183', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 
                                'TCAM': '1'}, 
                               {'DETA': 'Cuenta Corriente $', 'IMPO': '62527.36', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 
                                'TCAM': '1'}]}, 
        'Activos': [
            {'GTOS': '0', 'IMPO': '62527.36', 'ESPE': 'Subtotal Cuenta Corriente', 'TIPO': '11', 'Hora': '', 
             'Subtotal': [
                 {'IMPO': '62527.36', 'ESPE': '', 
                  'APERTURA': [
                      {'DETA': 'Vencido', 'IMPO': '45886.18', 'GTIA': None, 'ACUM': '45886.18'}, 
                      {'DETA': '24 Hs. 22/05/24', 'IMPO': '16641.18', 'GTIA': None, 'ACUM': '62527.36'}, 
                      {'DETA': '48 Hs. 23/05/24', 'IMPO': None, 'GTIA': None, 'ACUM': '62527.36'}, 
                      {'DETA': '72 Hs. 24/05/24', 'IMPO': None, 'GTIA': None, 'ACUM': '62527.36'}, 
                      {'DETA': '+ de 72 Hs.', 'IMPO': None, 'GTIA': None, 'ACUM': '62527.36'}, 
                      {'DETA': 'Gtia.Opciones', 'IMPO': None, 'GTIA': None, 'ACUM': '62527.36'}], 
                'Detalle': [
                    {'DETA': 'Disponible', 'IMPO': '45886.18', 'CANT': None, 'PCIO': '1'}, 
                    {'DETA': 'A Liq', 'IMPO': '16641.18', 'CANT': None, 'PCIO': '1'}], 
                    'TESP': '0', 'NERE': 'Pesos', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '11', 'Hora': 'Pesos', 'AMPL': '', 
                    'DIVI': '100', 'TICK': 'Pesos', 'CANT': None, 'PCIO': '1', 'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}], 
                    'CANT': None, 'TCAM': '1', 'CAN2': '52.6722015'}, 

            {'GTOS': '7118.04665', 'IMPO': '56183', 'ESPE': 'Subtotal Opciones', 'TIPO': '10', 'Hora': '', 
             'Subtotal': [
                 
                 {'IMPO': None, 'ESPE': '8118B', 
                  'Detalle': [{'DETA': 'Disponible', 'IMPO': '56501.2', 'CANT': '2', 'PCIO': '282.506'}, 
                              {'DETA': 'A Liq.', 'IMPO': '-56501.2', 'CANT': '-2', 'PCIO': '282.506'}], 
                    'TESP': '4', 'NERE': 'GFGC40608J', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '10', 'Hora': 'CIERRE', 
                    'AMPL': 'GFG(C) 4,060.800 JUNIO', 'DIVI': '100', 'TICK': 'GFGC40608J', 'CANT': None, 'PCIO': '282.506', 
                    'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}, 

                {'IMPO': '56183', 'ESPE': '8174B', 
                 'Detalle': [{'DETA': 'Disponible', 'IMPO': '22473.2', 'CANT': '2', 'PCIO': '112.366'}, 
                             {'DETA': 'A Liq.', 'IMPO': '33709.8', 'CANT': '3', 'PCIO': '112.366'}], 
                    'TESP': '4', 'NERE': 'GFGC44608J', 'GTOS': '7118.04665', 'DETA': 'Total', 'TIPO': '10', 'Hora': 'CIERRE', 
                    'AMPL': 'GFG(C) 4,460.800 JUNIO', 'DIVI': '100', 'TICK': 'GFGC44608J', 'CANT': '5', 'PCIO': '112.366', 
                    'CAN3': '14.5073951', 
                    'CAN2': '47.3277985', 'CAN0': '98.1299067'}], 'CANT': None, 'TCAM': '1', 'CAN2': '47.3277985'}]}}



os.system('cls')



def buscando():
    largo = 1
    subtotal = [ (i['DETA'],i['IMPO']) for i in portfolio["Result"]["Totales"]["Detalle"] ]
    print(subtotal)

    subtotal = [ i['Subtotal'] for i in portfolio["Result"]["Activos"][0:] ]
    for i in subtotal[0:]:
        if i[0]['NERE'] != 'Pesos':  
            subtotal = [ ( x['NERE'],x['CAN0'],x['CANT'],' || ',x['PCIO'],x['GTOS']) for x in i[0:] if x['CANT'] != None]
            for x in subtotal: print(x)
            for j in subtotal[0]:
                largo += len(j[0:])
    linea = '═' * largo * 2
    print(linea)
       
buscando()


'''
('Vencido', '45886.18', '45886.18')
('24 Hs. 22/05/24', '16641.18', '62527.36')
('GFGC44608J', '98.1299067', '5', ' || ', '112.366', '7118.04665')
6961
'''




  
#[ ]><   \n
#print("\nimprimir en linea nueva")