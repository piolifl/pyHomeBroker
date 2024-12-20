'''import pycocos
import pandas as pd
import xlwings as xw
import time


#wb = xw.Book('D:\\pyHomeBroker\\epgb_appCocos.xlsb')
#shtTickers = wb.sheets('pyRofex')
#shtData = wb.sheets('MATRIZ OMS')

# Inicializar la conexión a Cocos
app = pycocos.Cocos("miryam.borda79@gmail.com","Bordame1+")

# Obtener los datos de los instrumentos
lista_pesos = app.get_instrument_list_snapshot(
    instrument_type=app.instrument_types.BONOS, 
    instrument_subtype=app.instrument_subtypes.USD, 
    settlement=app.settlements.T1, 
    currency=app.currencies.USD, 
    segment=app.segments.DEFAULT)



for i in lista_pesos['items']:
    print(i)
'''
resultado  = {'short_ticker': 'AE38D', 'long_ticker': 'AE38D-0002-C-CT-USD', 'instrument_code': 'AE38', 'ext_code_cv': 5923, 'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2038', 'instrument_short_name': 'Argentina 2038', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'maturity': '2038-01-09', 'expires_at': '2038-01-09', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 1, 'currency': 'USD', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 3399, 'tick_size': 0.01, 'date': '2024-12-19', 'open': 72.7, 'high': 72.7, 'low': 71, 'close': 71.65, 'prev_close': 72.8, 'last': 71.65, 'bid': 70.11, 'ask': 73, 'bids': [{'size': 2000, 'price': 70.11}, {'size': 705, 'price': 70}, {'size': 7158, 'price': 69}, {'size': 50, 'price': 65.25}, {'size': 50, 'price': 62}], 'asks': [{'size': 3381, 'price': 73}, {'size': 1199, 'price': 74}, {'size': 200, 'price': 76.75}, {'size': 4557, 'price': 77.75}, {'size': 18154, 'price': 81.03}], 'turnover': 2208009.86, 'volume': 3084465, 'variation': -0.0157967, 'term': 4768, 'id_tick_size_rule': 'BYMA_FIXED_INCOME'}

for i in resultado['long_ticker']:
    print(i[0:])

#df = pd.DataFrame(lista_pesos)


    #df = pd.DataFrame(precios)

# Convertir los datos a un DataFrame de pandas299171299171


# Exportar los datos a un archivo Excel
#print(df)




resultado = [
    {
        'short_ticker': 'AL30', 
        'long_ticker': 'AL30-0002-C-CT-ARS', 
        'instrument_code': 'AL30', 
        'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2030', 
        'instrument_short_name': 'Argentina 2030', 
        'instrument_type': 'BONOS_PUBLICOS', 
        'instrument_subtype': 'NACIONALES_USD', 
        'logo_file_name': 'ARG.jpg', 
        'id_venue': 'BYMA', 
        'id_session': 'CT', 
        'id_segment': 'C', 
        'settlement_days': 1, 
        'currency': 'ARS', 
        'price_factor': 100, 
        'contract_size': 1, 
        'min_lot_size': 1, 
        'id_security': 3259, 
        'tick_size': 10, 
        'date': '2024-12-13', 
        'open': 77700, 
        'high': 78500, 
        'low': 77100, 
        'close': 78410, 
        'prev_close': 77080, 
        'last': 78410, 
        'bid': 77000, 
        'ask': 80000, 
        'bids': [
            {'size': 1827, 'price': 77000}, 
            {'size': 2597, 'price': 75000}, 
            {'size': 80000, 'price': 53480}, 
            {'size': 1, 'price': 1}], 
        'asks': [
            {'size': 10413, 'price': 80000}, 
            {'size': 234, 'price': 80800}, 
            {'size': 183, 'price': 83000}, 
            {'size': 300, 'price': 84000}, 
            {'size': 200, 'price': 90000}], 
        'turnover': 105857599791.9, 
        'volume': 136267302, 
        'variation': 0.0172548, 
        'term': '24hs', 
        'id_tick_size_rule': 'BYMA_FIXED_INCOME', 
        'is_favorite': True, 
        'newTerm': 1}, 
    {
        'short_ticker': 'AL30D', 
        'long_ticker': 'AL30D-0002-C-CT-USD', 
        'instrument_code': 'AL30', 
        'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2030', 
        'instrument_short_name': 'Argentina 2030', 
        'instrument_type': 'BONOS_PUBLICOS', 
        'instrument_subtype': 'NACIONALES_USD', 
        'logo_file_name': 'ARG.jpg', 
        'id_venue': 'BYMA', 
        'id_session': 'CT', 
        'id_segment': 'C', 
        'settlement_days': 1, 
        'currency': 'USD', 
        'price_factor': 100, 
        'contract_size': 1, 
        'min_lot_size': 1, 
        'id_security': 3286, 
        'tick_size': 0.01, 
        'date': '2024-12-13', 
        'open': 72.51, 
        'high': 73.15, 
        'low': 72.51, 
        'close': 73.07, 
        'prev_close': 72.84, 
        'last': 73.07, 
        'bid': 72, 
        'ask': 73.26, 
        'bids': [
            {'size': 6679, 'price': 72}, 
            {'size': 2600, 'price': 70.7}, 
            {'size': 670, 'price': 61.5}, 
            {'size': 500, 'price': 59.2}, 
            {'size': 1000, 'price': 57}], 
        'asks': [
            {'size': 1, 'price': 73.26}, 
            {'size': 1, 'price': 73.27}, 
            {'size': 1, 'price': 73.28}, 
            {'size': 1, 'price': 73.29}, 
            {'size': 42, 'price': 73.3}], 
        'turnover': 36507161.42, 
        'volume': 50109617, 
        'variation': 0.00315761, 
        'term': '24hs', 
        'id_tick_size_rule': 'BYMA_FIXED_INCOME', 
        'is_favorite': True, 
        'newTerm': 1}, 
    {
        'short_ticker': 'AL30C', 
        'long_ticker': 'AL30C-0002-C-CT-EXT', 
        'instrument_code': 'AL30', 
        'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2030', 
        'instrument_short_name': 'Argentina 2030', 
        'instrument_type': 'BONOS_PUBLICOS', 
        'instrument_subtype': 'NACIONALES_USD', 
        'logo_file_name': 'ARG.jpg', 
        'id_venue': 'BYMA', 
        'id_session': 'CT', 
        'id_segment': 'C', 
        'settlement_days': 1, 
        'currency': 'EXT', 
        'price_factor': 100, 
        'contract_size': 1, 
        'min_lot_size': 1, 
        'id_security': 3278, 
        'tick_size': 0.01, 
        'date': '2024-12-13', 
        'open': 71.66, 
        'high': 72.49, 
        'low': 71.33, 
        'close': 72.09, 
        'prev_close': 71.65, 
        'last': 72.09, 
        'bid': None, 
        'ask': None, 
        'bids': [
            {'size': 0, 'price': 0}], 
        'asks': [
            {'size': 0, 'price': 0}], 
        'turnover': 13457499.22, 
        'volume': 18766533, 
        'variation': 0.00614096, 
        'term': '24hs', 
        'id_tick_size_rule': 'BYMA_FIXED_INCOME', 
        'is_favorite': False, 
        'newTerm': 1}, 
    {
        'short_ticker': 'AL30C', 
        'long_ticker': 'AL30C-0001-C-CT-EXT', 
        'instrument_code': 'AL30', 
        'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2030', 
        'instrument_short_name': 'Argentina 2030', 
        'instrument_type': 'BONOS_PUBLICOS', 
        'instrument_subtype': 'NACIONALES_USD', 
        'logo_file_name': 'ARG.jpg', 
        'id_venue': 'BYMA', 
        'id_session': 'CT', 
        'id_segment': 'C', 
        'settlement_days': 0, 
        'currency': 'EXT', 
        'price_factor': 100, 
        'contract_size': 1, 
        'min_lot_size': 1, 
        'id_security': 3277, 
        'tick_size': 0.01, 
        'date': '2024-12-13', 
        'open': 71.67, 
        'high': 72, 
        'low': 71.31, 
        'close': 71.81, 
        'prev_close': 71.78, 
        'last': 71.81, 
        'bid': None, 
        'ask': 72.32, 
        'bids': [
            {'size': 0, 'price': 0}], 
        'asks': [
            {'size': 13727, 'price': 72.32}], 
        'turnover': 53020689.42, 
        'volume': 73995153, 
        'variation': 0.00041794, 
        'term': 'CI', 
        'id_tick_size_rule': 'BYMA_FIXED_INCOME', 
        'is_favorite': False, 
        'newTerm': 0}, 
    {
        'short_ticker': 'AL30D', 
        'long_ticker': 'AL30D-0001-C-CT-USD', 
        'instrument_code': 'AL30', 
        'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2030', 
        'instrument_short_name': 'Argentina 2030', 
        'instrument_type': 'BONOS_PUBLICOS', 
        'instrument_subtype': 'NACIONALES_USD', 
        'logo_file_name': 'ARG.jpg', 
        'id_venue': 'BYMA', 
        'id_session': 'CT', 
        'id_segment': 'C', 
        'settlement_days': 0, 
        'currency': 'USD', 
        'price_factor': 100, 
        'contract_size': 1, 
        'min_lot_size': 1, 
        'id_security': 3285, 
        'tick_size': 0.01, 
        'date': '2024-12-13', 
        'open': 72.78, 
        'high': 73.16, 
        'low': 72.56, 
        'close': 73.1, 
        'prev_close': 73.04, 
        'last': 73.1, 
        'bid': 69, 
        'ask': 73.23, 
        'bids': [
            {'size': 717, 'price': 69}, 
            {'size': 500, 'price': 59.2}], 
        'asks': [
            {'size': 9608, 'price': 73.23}, 
            {'size': 5000, 'price': 73.4}, 
            {'size': 7016, 'price': 73.5}, 
            {'size': 15196, 'price': 74}, 
            {'size': 1000, 'price': 74.5}], 
        'turnover': 157110249.88, 
        'volume': 215841903, 
        'variation': 0.00082147, 
        'term': 'CI', 
        'id_tick_size_rule': 'BYMA_FIXED_INCOME', 
        'is_favorite': True, 'newTerm': 0}, 
    {
        'short_ticker': 'AL30', 
        'long_ticker': 'AL30-0001-C-CT-ARS', 
        'instrument_code': 'AL30', 
        'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2030', 
        'instrument_short_name': 'Argentina 2030', 
        'instrument_type': 'BONOS_PUBLICOS', 
        'instrument_subtype': 'NACIONALES_USD', 
        'logo_file_name': 'ARG.jpg', 
        'id_venue': 'BYMA', 
        'id_session': 'CT', 
        'id_segment': 'C', 
        'settlement_days': 0, 
        'currency': 'ARS', 
        'price_factor': 100, 
        'contract_size': 1, 
        'min_lot_size': 1, 
        'id_security': 3258, 
        'tick_size': 10, 
        'date': '2024-12-13', 
        'open': 77080, 
        'high': 78300, 
        'low': 76600, 
        'close': 78260, 
        'prev_close': 77140, 
        'last': 78260, 
        'bid': 75800, 
        'ask': None, 
        'bids': [
            {'size': 2686, 'price': 75800}, 
            {'size': 10516, 'price': 75500}, 
            {'size': 200, 'price': 74510}, 
            {'size': 3300, 'price': 70500}, 
            {'size': 50, 'price': 68000}], 
        'asks': [
            {'size': 0, 'price': 0}], 
        'turnover': 197151131356, 
        'volume': 254556616, 
        'variation': 0.01451906, 
        'term': 'CI', 
        'id_tick_size_rule': 'BYMA_FIXED_INCOME', 
        'is_favorite': True, 
        'newTerm': 0}
]


#app.logout()