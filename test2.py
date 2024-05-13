import os
portfolio = {
    'Success': True, 
    'Error': {'Codigo': 0, 'Descripcion': None}, 
    'Result':{'Totales': {'TotalPosicion': '96368.62', 
                    'Detalle': [
                        {'DETA': 'Tenencia Opciones', 'IMPO': '45045.3', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}, 
                        {'DETA': 'Cuenta Corriente $', 'IMPO': '51323.32', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}]}, 
                    'Activos': [{'GTOS': '0', 'IMPO': '51323.32', 'ESPE': 'Subtotal Cuenta Corriente', 'TIPO': '11', 'Hora': '', 
                            'Subtotal': [{'IMPO': '51323.32', 'ESPE': '', 
                                          'APERTURA': [
                                              {'DETA': 'Vencido', 'IMPO': '100617.49', 'GTIA': None, 'ACUM': '100617.49'}, 
                                              {'DETA': '24 Hs. 13/05/24', 'IMPO': '-49294.17', 'GTIA': None, 'ACUM': '51323.32'}, 
                                              {'DETA': '48 Hs. 14/05/24', 'IMPO': None, 'GTIA': None, 'ACUM': '51323.32'}, 
                                              {'DETA': '72 Hs. 15/05/24', 'IMPO': None, 'GTIA': None, 'ACUM': '51323.32'}, 
                                              {'DETA': '+ de 72 Hs.', 'IMPO': None, 'GTIA': None, 'ACUM': '51323.32'}, 
                                              {'DETA': 'Gtia.Opciones', 'IMPO': None, 'GTIA': None, 'ACUM': '51323.32'}], 
                                              'Detalle': [
                                                  {'DETA': 'Disponible', 'IMPO': '100617.49', 'CANT': None, 'PCIO': '1'}, 
                                                  {'DETA': 'A Liq', 'IMPO': '-49294.17', 'CANT': None, 'PCIO': '1'}], 
                                            'TESP': '0', 'NERE': 'Pesos', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '11', 'Hora': 'Pesos', 'AMPL': '', 
                                            'DIVI': '100', 'TICK': 'Pesos', 'CANT': None, 'PCIO': '1', 'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}], 
                                            'CANT': None, 'TCAM': '1', 'CAN2': '53.2572948'}, 
                                            {'GTOS': '-6709.2415', 'IMPO': '45045.3', 'ESPE': 'Subtotal Opciones', 'TIPO': '10', 'Hora': '', 
                                             'Subtotal': [
                                                 {'IMPO': '62424.4', 'ESPE': '8139B', 'TESP': '4', 'NERE': 'GFGC37059J', 
                                                  'GTOS': '-6659.39', 'DETA': 'A Liq', 'TIPO': '10', 'Hora': 'ANTERIOR', 
                                                  'AMPL': 'GFG(C) 3,705.900 JUNIO', 'DIVI': '100', 'TICK': 'GFGC37059J', 
                                                  'CANT': '2', 'PCIO': '312.122', 'CAN3': '-9.639584', 'CAN2': '64.7766877', 
                                                  'CAN0': '345.41895'}, 
                                                 {'IMPO': '-17557', 'ESPE': '8135B', 'TESP': '4', 'NERE': 'GFGC40059J', 
                                                  'GTOS': '21.7', 'DETA': 'A Liq', 'TIPO': '10', 'Hora': 'ANTERIOR', 
                                                  'AMPL': 'GFG(C) 4,005.900 JUNIO', 'DIVI': '100', 'TICK': 'GFGC40059J', 
                                                  'CANT': '-1', 'PCIO': '175.57', 'CAN3': '-.1234449', 'CAN2': '-18.2185861', 
                                                  'CAN0': '175.787'}, 
                                                 {'IMPO': '177.9', 'ESPE': '8143B', 'Detalle': [{'DETA': 'Disponible', 
                                                                                                 'IMPO': '1186', 'CANT': '20', 'PCIO': '.593'}, {'DETA': 'A Liq.', 'IMPO': '-1008.1', 'CANT': '-17', 'PCIO': '.593'}], 
                                                  'TESP': '4', 'NERE': 'GFGV23559J', 'GTOS': '-71.5515', 'DETA': 'Total', 'TIPO': '10', 'Hora': 'ANTERIOR', 'AMPL': 'GFG(V) 2,355.900 JUNIO', 'DIVI': '100', 'TICK': 'GFGV23559J', 'CANT': '3', 'PCIO': '.593', 'CAN3': '-28.6835317', 'CAN2': '.1846037', 'CAN0': '.831505'}], 'CANT': None, 'TCAM': '1', 'CAN2': '46.7427052'}]}}


#portfolio = portfolio["Result"]["Activos"][0:]
#portfolio = [ (x['NERE']) for x in portfolio[1]['Subtotal'][0]]
#portfolio = portfolio["Result"]["Activos"][1]["Subtotal"]
os.system('cls')

print()
subtotal = [ i['Subtotal'] for i in portfolio["Result"]["Activos"][0:] ]
for i in subtotal[0:]:
    if i[0]['NERE'] == 'Pesos': 
        subtotal = [ (x['DETA'],x['IMPO'],x['ACUM']) for x in i[0]['APERTURA'] if x['IMPO'] != None]
    else: subtotal = [ (x['NERE'],x['CAN0'],x['CANT'],x['PCIO'],x['GTOS']) for x in i[0:]]
    print(subtotal)
print()

'''subtotal = [ i[0]['NERE'] for i in subtotal[0:] ]
print(subtotal)
print()'''









  
#[ ]><   \n
#print("\nimprimir en linea nueva")