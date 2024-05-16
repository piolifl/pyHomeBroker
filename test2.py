import os
portfolio = {
    'Success': True, 
    'Error': {'Codigo': 0, 'Descripcion': None}, 
    'Result': {'Totales': 
               {'TotalPosicion': '129590.89', 
                'Detalle': [{'DETA': 'Tenencia Opciones', 'IMPO': '8700', 'TIPO': '1', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}, 
                            {'DETA': 'Tenencia a Liquidar', 'IMPO': '30875', 'TIPO': '1', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}, 
                            {'DETA': 'Cuenta Corriente $', 'IMPO': '90015.89', 'TIPO': '1', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}]}, 
                'Activos': [
                    {'GTOS': '0', 'IMPO': '90015.89', 'ESPE': 'Subtotal Cuenta Corriente', 'TIPO': '11', 'Hora': '', 
                     'Subtotal': [{'IMPO': '90015.89', 'ESPE': '', 
                                   'APERTURA': [
                                       {'DETA': 'Vencido', 'IMPO': '38664.03', 'GTIA': None, 'ACUM': '38664.03'}, 
                                       {'DETA': '24 Hs. 15/05/24', 'IMPO': '51351.86', 'GTIA': None, 'ACUM': '90015.89'}, 
                                       {'DETA': '48 Hs. 16/05/24', 'IMPO': None, 'GTIA': None, 'ACUM': '90015.89'}, 
                                       {'DETA': '72 Hs. 17/05/24', 'IMPO': None, 'GTIA': None, 'ACUM': '90015.89'}, 
                                       {'DETA': '+ de 72 Hs.', 'IMPO': None, 'GTIA': None, 'ACUM': '90015.89'}, 
                                       {'DETA': 'Gtia.Opciones', 'IMPO': None, 'GTIA': None, 'ACUM': '90015.89'}], 
                                       'Detalle': [
                                           {'DETA': 'Disponible', 'IMPO': '38664.03', 'CANT': None, 'PCIO': '1'}, 
                                           {'DETA': 'A Liq', 'IMPO': '51351.86', 'CANT': None, 'PCIO': '1'}], 
                                        'TESP': '0', 'NERE': 'Pesos', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '11', 'Hora': 'Pesos', 'AMPL': '', 
                                        'DIVI': '100', 'TICK': 'Pesos', 'CANT': None, 'PCIO': '1', 'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}], 
                                        'CANT': None, 'TCAM': '1', 'CAN2': '69.4615879'}, 
                    {'GTOS': '0', 'IMPO': '8700', 'ESPE': 'Subtotal Opciones', 'TIPO': '10', 'Hora': '', 
                     'Subtotal': [{'IMPO': '8700', 'ESPE': '8139B', 
                                   'Detalle': [
                                       {'DETA': 'Disponible', 'IMPO': '35000', 'CANT': '1', 'PCIO': '350'}, 
                                       {'DETA': 'A Liq.', 'IMPO': '-35000', 'CANT': '-1', 'PCIO': '350'}], 
                                'TESP': '4', 'NERE': 'GFGC37059J', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '10', 'Hora': '16:54:21', 
                                'AMPL': 'GFG(C) 3,705.900 JUNIO', 'DIVI': '100', 'TICK': 'GFGC37059J', 'CANT': '1', 'PCIO': '350', 
                                'CAN3': '-111.7018894', 'CAN2': '0', 'CAN0': '-2990.97'}, {'IMPO': None, 'ESPE': '8129B', 
                                    'Detalle': [
                                        {'DETA': 'Disponible', 'IMPO': '27000', 'CANT': '1', 'PCIO': '270'}, 
                                        {'DETA': 'A Liq.', 'IMPO': '-27000', 'CANT': '-1', 'PCIO': '270'}], 
                                'TESP': '4', 'NERE': 'GFGC38559J', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '10', 'Hora': '16:54:53', 
                                'AMPL': 'GFG(C) 3,855.900 JUNIO', 'DIVI': '100', 'TICK': 'GFGC38559J', 'CANT': None, 'PCIO': '270', 'CAN3': '-104.8716589', 'CAN2': '0', 'CAN0': '-5542.26'}, 
                                {'IMPO': None, 'ESPE': '8143B', 'Detalle': [{'DETA': 'Disponible', 'IMPO': '79.7', 'CANT': '1', 'PCIO': '.797'}, 
                                                                            {'DETA': 'A Liq.', 'IMPO': '-79.7', 'CANT': '-1', 'PCIO': '.797'}], 
                                        'TESP': '4', 'NERE': 'GFGV23559J', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '10', 
                                        'Hora': '16:54:14', 'AMPL': 'GFG(V) 2,355.900 JUNIO', 'DIVI': '100', 
                                        'TICK': 'GFGV23559J', 'CANT': None, 'PCIO': '.797', 'CAN3': '-106.8709858', 
                                        'CAN2': '0', 'CAN0': '-11.5995'}, 
                                        {'IMPO': '8700', 'ESPE': '8130B', 'TESP': '4','NERE': 'GFGV34059J', 'GTOS': '0', 
                                         'DETA': 'A Liq', 'TIPO': '10', 'Hora': '16:54:42', 'AMPL': 'GFG(V) 3,405.900 JUNIO', 
                                         'DIVI': '100', 'TICK': 'GFGV34059J', 'CANT': '3', 'PCIO': '29', 'CAN3': '0', 'CAN2': '0', 
                                         'CAN0': '0'}], 'CANT': None, 'TCAM': '1', 'CAN2': '6.7134349'}, 

                    {'GTOS': '0', 'IMPO': '30875', 'ESPE': 'Subtotal Titulos Publicos DOLAR MEP', 'TIPO': '1', 'Hora': '', 
                     'Subtotal': [{'IMPO': '30875', 'ESPE': '05921', 'TESP': '1', 'NERE': 'AL30', 'GTOS': '0', 'DETA': 'A Liq', 
                                   'TIPO': '1', 'Hora': '16:54:48', 'AMPL': 'BONO REP. ARGENTINA USD STEP UP 2030', 'DIVI': '.01', 
                                   'TICK': 'AL30', 'CANT': '50', 'PCIO': '61750', 'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}], 
                                   'CANT': None, 'TCAM': '1', 
                     'CAN2': '23.8249772' }
                       ] } }
#portfolio = portfolio["Result"]["Activos"][0:]
#portfolio = [ (x['NERE']) for x in portfolio[1]['Subtotal'][0]]
#portfolio = portfolio["Result"]["Activos"][1]["Subtotal"]
os.system('cls')

print()
subtotal = [ i['Subtotal'] for i in portfolio["Result"]["Activos"][0:] ]

for i in subtotal[0:]:

    if i[0]['NERE'] == 'Pesos': 
        try: subtotal = [ (x['DETA'],x['IMPO']) for x in i[0]['Detalle']]
        except: subtotal = [ (x['DETA'],x['IMPO'],x['ACUM']) for x in i[0] if x['IMPO'] != None]
        print(subtotal)
    else: 
        subtotal = [ (x['NERE'],x['CAN0'],x['CANT'],x['PCIO'],x['IMPO'],x['GTOS']) for x in i[0:] if x['CANT'] != None]
        for x in subtotal: print(x)


    









  
#[ ]><   \n
#print("\nimprimir en linea nueva")