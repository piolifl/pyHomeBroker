import os
portfolio = {
    'Success': True, 
    'Error': {'Codigo': 0, 'Descripcion': None}, 'Result': 
                                                        {'Totales': 
                                                                    {'TotalPosicion': '211526.45', 'Detalle': [
                                                                                                                {'DETA': 'Tenencia Opciones', 'IMPO': '96098.9', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}, 
                                                                                                                {'DETA': 'Cuenta Corriente $', 'IMPO': '115427.55', 'TIPO': '10', 'Hora': 'Pesos', 'CANT': None, 'TCAM': '1'}]}, 
                                                                    'Activos': [

                                                                        {'GTOS': '0', 'IMPO': '115427', 'ESPE': 'Subtotal Cuenta Corriente', 'TIPO': '11', 'Hora': '', 
                                                                                'Subtotal': [
                                                                                        {'IMPO': '115427.55', 'ESPE': '', 
                                                                                        'APERTURA': [
                                                                                                {'DETA': 'Vencido', 'IMPO': '115427.55', 'GTIA': None, 'ACUM': '115427.55'}, 
                                                                                                {'DETA': '24 Hs. 08/05/24', 'IMPO': None, 'GTIA': None, 'ACUM': '115427.55'}, 
                                                                                                {'DETA': '48 Hs. 09/05/24', 'IMPO': None, 'GTIA': None, 'ACUM': '115427.55'}, 
                                                                                                {'DETA': '72 Hs. 10/05/24', 'IMPO': None, 'GTIA': None, 'ACUM': '115427.55'}, 
                                                                                                {'DETA': '+ de 72 Hs.', 'IMPO': None, 'GTIA': None, 'ACUM': '115427.55'}, 
                                                                                                {'DETA': 'Gtia.Opciones', 'IMPO': None, 'GTIA': None, 'ACUM': '115427.55'}], 
                                                                                                'TESP': '0', 'NERE': 'Pesos', 'GTOS': '0', 'DETA': '', 'TIPO': '11', 'Hora': 'Pesos', 'AMPL': '', 'DIVI': '100', 
                                                                                                'TICK': 'Pesos', 'CANT': None, 'PCIO': '1', 'CAN3': '0', 'CAN2': '0', 'CAN0': '0'}], 
                                                                        'CANT': None, 'TCAM': '1', 'CAN2': '54.5688494'}, 

                                                                        {'GTOS': '1687.38998', 'IMPO': '96098.9', 'ESPE': 'Subtotal Opciones', 'TIPO': '10', 'Hora': '', 
                                                                                'Subtotal': [
{'IMPO': '49869.5', 'ESPE': '8174B', 'TESP': '4', 'NERE': 'GFGC4600JU', 'GTOS': '309.6', 'DETA': '', 'TIPO': '10', 'Hora': 'ANTERIOR', 
'AMPL': 'GFG(C) 4600.000 JUNIO', 'DIVI': '100', 'TICK': 'GFGC4600JU', 'CANT': '5', 'PCIO': '99.739', 'CAN3': '.6246986', 'CAN2': '23.5760114', 'CAN0': '99.1198'}, 
                                                                                    {'IMPO': '46229.4', 'ESPE': '8183B', 'TESP': '4', 'NERE': 'GFGC4800JU', 'GTOS': '1377.78998', 'DETA': '', 'TIPO': '10', 'Hora': 'ANTERIOR', 
                                                                                        'AMPL': 'GFG(C) 4800.000 JUNIO', 'DIVI': '100', 'TICK': 'GFGC4800JU', 'CANT': '7', 'PCIO': '66.042', 'CAN3': '3.0718852', 'CAN2': '21.8551392', 'CAN0': '64.0737286'}], 
                                                                        'CANT': None, 'TCAM': '1', 'CAN2': '45.4311506'}]}}

#portfolio = portfolio["Result"]["Activos"][0:]
#portfolio = [ (x['NERE']) for x in portfolio[1]['Subtotal'][0]]
#portfolio = portfolio["Result"]["Activos"][1]["Subtotal"]
os.system('cls')

print()
subtotal = [ i['Subtotal'] for i in portfolio["Result"]["Activos"][0:] ]
for i in subtotal[0:]:
    if i[0]['NERE'] == 'Pesos': subtotal = [ (x['NERE'],x['IMPO']) for x in i[0:]]
    else: subtotal = [ (x['NERE'],x['CANT'],x['PCIO'],x['IMPO'],x['Hora']) for x in i[0:]]
    print(subtotal)
print()

'''subtotal = [ i[0]['NERE'] for i in subtotal[0:] ]
print(subtotal)
print()'''









  
#[ ]><   \n
