import pycocos
import pandas as pd
import xlwings as xw
import time

wb = xw.Book('D:\\pyHomeBroker\\epgb_appCocos.xlsb')
#shtTickers = wb.sheets('pyRofex')
shtData = wb.sheets('MATRIZ OMS')

# Inicializar la conexión a Cocos
app = pycocos.Cocos("miryam.borda79@gmail.com","Bordame1+")

# Obtener los datos de los instrumentos
lista_pesos = app.get_instrument_list_snapshot(
    instrument_type=app.instrument_types.BONOS, 
    instrument_subtype=app.instrument_subtypes.USD, 
    settlement=app.settlements.T1, 
    currency=app.currencies.PESOS, 
    segment=app.segments.DEFAULT)

df = pd.DataFrame(lista_pesos['items'])

shtData.range('A2').options(index=True, headers=True).value = df

'''for i in lista_pesos['items']:
    print(i)'''

resulta = {'short_ticker': 'AE38', 'long_ticker': 'AE38-0002-C-CT-ARS', 'instrument_code': 'AE38', 'ext_code_cv': 5923, 'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2038', 'instrument_short_name': 'Argentina 2038', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'maturity': '2038-01-09', 'expires_at': '2038-01-09', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 1, 'currency': 'ARS', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 3375, 'tick_size': 10, 'date': '2024-12-20', 'open': 81500, 'high': 81900, 'low': 80310, 'close': 81560, 'prev_close': 81160, 'last': 81560, 'bid': 76500, 'ask': 90000, 'bids': [{'size': 10606, 'price': 76500}], 'asks': [{'size': 13000, 'price': 90000}, {'size': 55464, 'price': 91000}, {'size': 10732, 'price': 100000}], 'turnover': 12719960199.3, 'volume': 15616723, 'variation': 0.00492854, 'term': 4767, 'id_tick_size_rule': 'BYMA_FIXED_INCOME'}
{'short_ticker': 'AL29', 'long_ticker': 'AL29-0002-C-CT-ARS', 'instrument_code': 'AL29', 'ext_code_cv': 5927, 'instrument_name': 'BONO REP ARGENTINA USD 1% 2029', 'instrument_short_name': 'Argentina 2029 1%', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'maturity': '2029-07-09', 'expires_at': '2029-07-09', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 1, 'currency': 'ARS', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 3562, 'tick_size': 10, 'date': '2024-12-20', 'open': 92000, 'high': 92900, 'low': 91110, 'close': 92800, 'prev_close': 91750, 'last': 92800, 'bid': 89100, 'ask': 96000, 'bids': [{'size': 500, 'price': 89100}, {'size': 1000, 'price': 89020}, {'size': 500, 'price': 86100}, {'size': 10, 'price': 81000}, {'size': 10, 'price': 78500}], 'asks': [{'size': 10, 'price': 96000}, {'size': 10, 'price': 98500}, {'size': 10210, 'price': 100000}], 'turnover': 1406509394.5, 'volume': 1522395, 'variation': 0.01144414, 'term': 1661, 'id_tick_size_rule': 'BYMA_FIXED_INCOME'}
{'short_ticker': 'AL30', 'long_ticker': 'AL30-0002-C-CT-ARS', 'instrument_code': 'AL30', 'ext_code_cv': 5921, 'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2030', 'instrument_short_name': 'Argentina 2030', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'maturity': '2030-07-09', 'expires_at': '2030-07-09', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 1, 'currency': 'ARS', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 3259, 'tick_size': 10, 'date': '2024-12-20', 'open': 83700, 'high': 84660, 'low': 83280, 'close': 84500, 'prev_close': 83900, 'last': 84500, 'bid': 84000, 'ask': 84870, 'bids': [{'size': 2366, 'price': 84000}, {'size': 1201, 'price': 82500}, {'size': 916, 'price': 81500}, {'size': 700, 'price': 80200}, {'size': 268, 'price': 80000}], 'asks': [{'size': 1000, 'price': 84870}, {'size': 100000, 'price': 85000}, {'size': 8150, 'price': 85700}, {'size': 1500, 'price': 85950}, {'size': 100000, 'price': 86800}], 'turnover': 199142957971.9, 'volume': 236924390, 'variation': 0.00715137, 'term': 2026, 'id_tick_size_rule': 'BYMA_FIXED_INCOME'}
{'short_ticker': 'AL35', 'long_ticker': 'AL35-0002-C-CT-ARS', 'instrument_code': 'AL35', 'ext_code_cv': 5922, 'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2035', 'instrument_short_name': 'Argentina 2035', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'maturity': '2035-07-09', 'expires_at': '2035-07-09', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 1, 'currency': 'ARS', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 3319, 'tick_size': 10, 'date': '2024-12-20', 'open': 76610, 'high': 79140, 'low': 75710, 'close': 78450, 'prev_close': 76580, 'last': 78450, 'bid': 71950, 'ask': 81500, 'bids': [{'size': 20, 'price': 71950}, {'size': 20, 'price': 70000}, {'size': 20, 'price': 68050}, {'size': 1, 'price': 5000}], 'asks': [{'size': 17874, 'price': 81500}, {'size': 11963, 'price': 100000}], 'turnover': 3682560709.9, 'volume': 4775879, 'variation': 0.02441891, 'term': 3852, 'id_tick_size_rule': 'BYMA_FIXED_INCOME'}
{'short_ticker': 'AL41', 'long_ticker': 'AL41-0002-C-CT-ARS', 'instrument_code': 'AL41', 'ext_code_cv': 5924, 'instrument_name': 'BONO REP. ARGENTINA USD STEP UP 2041', 'instrument_short_name': 'Argentina 2041', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'maturity': '2041-07-09', 'expires_at': '2041-07-09', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 1, 'currency': 'ARS', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 3431, 'tick_size': 10, 'date': '2024-12-20', 'open': 72950, 'high': 73890, 'low': 72170, 'close': 73890, 'prev_close': 72950, 'last': 73890, 'bid': 72000, 'ask': 76700, 'bids': [{'size': 800, 'price': 72000}], 'asks': [{'size': 84, 'price': 76700}, {'size': 10, 'price': 97700}], 'turnover': 1827427020.4, 'volume': 2507285, 'variation': 0.01288554, 'term': 6044, 'id_tick_size_rule': 'BYMA_FIXED_INCOME'}
{'short_ticker': 'BPY26', 'long_ticker': 'BPY26-0002-C-CT-ARS', 'instrument_code': 'BPY26', 'ext_code_cv': 9247, 'instrument_name': 'BOPREAL S.3 VTO31/05/26', 'instrument_short_name': 'BOPREAL EN USD SERIE 3', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'maturity': '2026-05-31', 'expires_at': '2026-05-31', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 1, 'currency': 'ARS', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 100, 'id_security': 107586856, 'tick_size': 10, 'date': '2024-12-20', 'open': 106800, 'high': 107500, 'low': 103630, 'close': 105990, 'prev_close': 104660, 'last': 105990, 'bid': 102000, 'ask': 110000, 'bids': [{'size': 100, 'price': 102000}, {'size': 2500, 'price': 99020}, {'size': 100, 'price': 13600}], 'asks': [{'size': 100, 'price': 110000}, {'size': 1100, 'price': 110500}], 'turnover': 17061276090, 'volume': 16230400, 'variation': 0.01270782, 'term': 526, 'id_tick_size_rule': 'BYMA_FIXED_INCOME'}
{'short_ticker': 'GD29', 'long_ticker': 'GD29-0002-C-CT-ARS', 'instrument_code': 'GD29', 'ext_code_cv': 81274, 'instrument_name': 'Bonos República Argentina U$S 1% Step Up V.09/07/29', 'instrument_short_name': 'Argentina 2029 (NY)', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'maturity': '2029-07-09', 'expires_at': '2029-07-09', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 1, 'currency': 'ARS', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 12488, 'tick_size': 10, 'date': '2024-12-20', 'open': 93560, 'high': 94560, 'low': 93050, 'close': 94040, 'prev_close': 93560, 'last': 94040, 'bid': None, 'ask': None, 'bids': [{'size': 0, 'price': 0}], 'asks': [{'size': 0, 'price': 0}], 'turnover': 322291420.5, 'volume': 343127, 'variation': 0.0051304, 'term': 1661, 'id_tick_size_rule': 'BYMA_FIXED_INCOME'}
{'short_ticker': 'GD30', 'long_ticker': 'GD30-0002-C-CT-ARS', 'instrument_code': 'GD30', 'ext_code_cv': 81086, 'instrument_name': 'Bonos República Argentina U$S Step Up V.09/07/30', 'instrument_short_name': 'Argentina 2030 (NY)', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'maturity': '2030-07-09', 'expires_at': '2030-07-09', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 1, 'currency': 'ARS', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 12220, 'tick_size': 10, 'date': '2024-12-20', 'open': 84750, 'high': 86100, 'low': 83890, 'close': 86100, 'prev_close': 84750, 'last': 86100, 'bid': 53480, 'ask': None, 'bids': [{'size': 50000, 'price': 53480}], 'asks': [{'size': 0, 'price': 0}], 'turnover': 26991079036.1, 'volume': 31810930, 'variation': 0.0159292, 'term': 2026, 'id_tick_size_rule': 'BYMA_FIXED_INCOME'}
{'short_ticker': 'GD35', 'long_ticker': 'GD35-0002-C-CT-ARS', 'instrument_code': 'GD35', 'ext_code_cv': 81088, 'instrument_name': 'Bonos República Argentina U$S Step Up V.09/07/35', 'instrument_short_name': 'Argentina 2035 (NY)', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'maturity': '2035-07-09', 'expires_at': '2035-07-09', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 1, 'currency': 'ARS', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 12276, 'tick_size': 10, 'date': '2024-12-20', 'open': 77900, 'high': 78590, 'low': 76470, 'close': 78430, 'prev_close': 77250, 'last': 78430, 'bid': 77000, 'ask': 80000, 'bids': [{'size': 1000, 'price': 77000}, {'size': 469, 'price': 75100}, {'size': 2987, 'price': 75000}, {'size': 50000, 'price': 43910}, {'size': 142978, 'price': 770}], 'asks': [{'size': 37700, 'price': 80000}, {'size': 74, 'price': 81500}, {'size': 83214, 'price': 88000}], 'turnover': 27472959949.4, 'volume': 35347052, 'variation': 0.01527508, 'term': 3852, 'id_tick_size_rule': 'BYMA_FIXED_INCOME'}
{'short_ticker': 'GD38', 'long_ticker': 'GD38-0002-C-CT-ARS', 'instrument_code': 'GD38', 'ext_code_cv': 81090, 'instrument_name': 'Bonos República Argentina U$S Step Up V.09/01/38', 'instrument_short_name': 'Argentina 2038 (NY)', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'maturity': '2038-01-09', 'expires_at': '2038-01-09', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 1, 'currency': 'ARS', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 12332, 'tick_size': 10, 'date': '2024-12-20', 'open': 82640, 'high': 83050, 'low': 81200, 'close': 82790, 'prev_close': 82050, 'last': 82790, 'bid': 80000, 'ask': None, 'bids': [{'size': 2000, 'price': 80000}, {'size': 1872, 'price': 79600}], 'asks': [{'size': 0, 'price': 0}], 'turnover': 8596081272.8, 'volume': 10436576, 'variation': 0.00901889, 'term': 4767, 'id_tick_size_rule': 'BYMA_FIXED_INCOME'}
{'short_ticker': 'GD41', 'long_ticker': 'GD41-0002-C-CT-ARS', 'instrument_code': 'GD41', 'ext_code_cv': 81092, 'instrument_name': 'Bonos República Argentina U$S Step Up V.09/07/41', 'instrument_short_name': 'Argentina 2041 (NY)', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'maturity': '2041-07-09', 'expires_at': '2041-07-09', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 1, 'currency': 'ARS', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 12387, 'tick_size': 10, 'date': '2024-12-20', 'open': 73500, 'high': 74380, 'low': 72010, 'close': 74080, 'prev_close': 72800, 'last': 74080, 'bid': None, 'ask': None, 'bids': [{'size': 0, 'price': 0}], 'asks': [{'size': 0, 'price': 0}], 'turnover': 4954673333.3, 'volume': 6774945, 'variation': 0.01758242, 'term': 6044, 'id_tick_size_rule': 'BYMA_FIXED_INCOME'}
{'short_ticker': 'GD46', 'long_ticker': 'GD46-0002-C-CT-ARS', 'instrument_code': 'GD46', 'ext_code_cv': 81093, 'instrument_name': 'Bonos República Argentina U$S Step Up V.09/07/46', 'instrument_short_name': 'Argentina 2046 (NY)', 'instrument_type': 'BONOS_PUBLICOS', 'instrument_subtype': 'NACIONALES_USD', 'maturity': '2046-07-09', 'expires_at': '2046-07-09', 'logo_file_name': 'ARG.jpg', 'id_venue': 'BYMA', 'id_session': 'CT', 'id_segment': 'C', 'settlement_days': 1, 'currency': 'ARS', 'price_factor': 100, 'contract_size': 1, 'min_lot_size': 1, 'id_security': 12441, 'tick_size': 10, 'date': '2024-12-20', 'open': 79300, 'high': 80770, 'low': 78000, 'close': 79300, 'prev_close': 79200, 'last': 79300, 'bid': 55500, 'ask': 84500, 'bids': [{'size': 35, 'price': 55500}], 'asks': [{'size': 416, 'price': 84500}], 'turnover': 174666778.7, 'volume': 221185, 'variation': 0.00126263, 'term': 7870, 'id_tick_size_rule': 'BYMA_FIXED_INCOME'}

'''for x in resulta['short_ticker']:
    print(x)'''

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


app.logout()