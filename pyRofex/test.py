portfolio = {'Success': True, 
    'Error': {'Codigo': 0, 'Descripcion': None}, 
    'Result': {'Totales': 
               {'TotalPosicion': '1082540.1624925', 
                'Detalle': [
                    {'DETA': 'Tenencia a Liquidar', 'IMPO': '76660', 'TIPO': '1', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}, 
                    {'DETA': 'Cuenta Corriente $', 'IMPO': '443749.58', 'TIPO': '1', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}, 
                    {'DETA': 'Cuenta Corriente USD MEP', 'IMPO': '562130.5824925', 'TIPO': '1', 'Hora': 'USD MEP', 'CANT': '482.13', 'TCAM': '1165.9315589'}]}, 
                'Activos': [
                    {'GTOS': '0', 'IMPO': '1005880.1624925', 'ESPE': 'Subtotal Cuenta Corriente', 'TIPO': '11', 'Hora': '', 
                            'Subtotal': [
                                {'IMPO': '443749.58', 'ESPE': '', 
                                    'APERTURA': [
                                        {'DETA': 'Vencido', 'IMPO': '497506.59', 'GTIA': None, 'ACUM': '497506.59'}, 
                                        {'DETA': '24 Hs. 21/04/25', 'IMPO': '-53757.01', 'GTIA': None, 'ACUM': '443749.58'}, 
                                        {'DETA': '48 Hs. 22/04/25', 'IMPO': None, 'GTIA': None, 'ACUM': '443749.58'}, 
                                        {'DETA': '72 Hs. 23/04/25', 'IMPO': None, 'GTIA': None, 'ACUM': '443749.58'}, 
                                        {'DETA': '+ de 72 Hs.', 'IMPO': None, 'GTIA': None, 'ACUM': '443749.58'}, 
                                        {'DETA': 'Gtia.Opciones', 'IMPO': None, 'GTIA': None, 'ACUM': '443749.58'}], 

                                'Detalle': [
                                    {'DETA': 'Disponible', 'IMPO': '497506.59', 'CANT': None, 'PCIO': '1'}, 
                                    {'DETA': 'A Liq', 'IMPO': '-53757.01', 'CANT': None, 'PCIO': '1'}], 

                                    'TESP': '8', 'NERE': 'Pesos', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '11', 'Hora': 'Pesos', 
                                    'AMPL': '', 'DIVI': '1', 'TICK': 'Pesos', 'CANT': None, 'PCIO': '1', 'CAN3': '0', 'CAN2': '0', 
                                    'CAN0': '0'}, 
                            
                                {'IMPO': '562130.58', 'ESPE': '', 
                                    'APERTURA': [
                                        {'DETA': 'Vencido', 'IMPO': '415.75', 'GTIA': None, 'ACUM': '415.75'}, 
                                        {'DETA': '24 Hs. 21/04/25', 'IMPO': '66.38', 'GTIA': None, 'ACUM': '482.13'}, 
                                        {'DETA': '48 Hs. 22/04/25', 'IMPO': None, 'GTIA': None, 'ACUM': '482.13'}, 
                                        {'DETA': '72 Hs. 23/04/25', 'IMPO': None, 'GTIA': None, 'ACUM': '482.13'}, 
                                        {'DETA': '+ de 72 Hs.', 'IMPO': None, 'GTIA': None, 'ACUM': '482.13'}, 
                                        {'DETA': 'Gtia.Opciones', 'IMPO': None, 'GTIA': None, 'ACUM': '482.13'}], 
                                        
                                'Detalle': [
                                    {'DETA': 'Disponible', 'IMPO': '484736.05', 'CANT': '415.75', 'PCIO': '1165.9315589'}, 
                                    {'DETA': 'A Liq', 'IMPO': '77394.54', 'CANT': '66.38', 'PCIO': '1165.9315589'}], 
                                    
                                    'TESP': '8', 'NERE': 'USD MEP', 'GTOS': '0', 'DETA': 'Total', 'TIPO': '11', 'Hora': 'USD MEP', 
                                    'AMPL': '', 'DIVI': '1', 'TICK': 'USD MEP', 'CANT': '482.13', 'PCIO': '1165.9315589', 
                                    'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}], 
                                    'CANT': None, 'TCAM': '1', 'CAN2': '92.91850754'}, 

                    {'GTOS': '-317.695', 'IMPO': '76660', 'ESPE': 'Subtotal Titulos Publicos','TIPO': '1', 'Hora': '', 
                        'Subtotal': [{'IMPO': '76660', 'ESPE': '05921', 'TESP': '8', 'NERE': 'AL30', 'GTOS': '-317.695', 
                                      'DETA': 'A Liq', 'TIPO': '1', 
                                'Hora': 'ANTERIOR', 'AMPL': 'BONO REP. ARGENTINA USD STEP UP 2030', 'DIVI': '.01', 
                                'TICK': 'AL30', 'CANT': '100', 
                                'PCIO': '76660', 'CAN3': '-.41271046', 'CAN2': '7.08149246', 'CAN0': '76977.695'}], 
                                'CANT': None, 'TCAM': '1', 'CAN2': '7.0814925'}]}}


print(portfolio['Result']['Activos'][0]['Subtotal'][0]['APERTURA'][1]['ACUM'])
print(portfolio['Result']['Activos'][0]['Subtotal'][1]['APERTURA'][1]['ACUM'])
print()


#subtotal = [ i['Subtotal'][0] for i in portfolio["Result"]["Activos"][0:]]
for i in portfolio["Result"]["Activos"][0]['Subtotal']:
    for x in i['APERTURA'][0:]:
        print(x)

'''
None
None
415.75
66.38

'''