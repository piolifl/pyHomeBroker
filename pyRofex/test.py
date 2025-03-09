import xlwings as xw

wb = xw.Book('D:\\pyHomeBroker\\epgb.xlsb')
shtTickers = wb.sheets('pyRofex')
shtData = wb.sheets('MATRIZ OMS')

arsCI = {'BID': 0, 'ASK': 0, 'LAST': 0}
mepCI = {'BID': 0, 'ASK': 0, 'LAST': 0}
cclCI = {'BID': 0, 'ASK': 0, 'LAST': 0}


def excelRulo(celda):

    for valor in shtData.range('A65:A70').value:
        
        ticker = str(valor).split()

        bid = shtData.range('C'+str(int(celda+1))).value
        ask = shtData.range('D'+str(int(celda+1))).value
        last = shtData.range('F'+str(int(celda+1))).value

        

        if ticker[2] == 'CI':
            print(ticker[0], end=' ')
            moneda = ticker[0][-1:]
            if moneda == '0':
                if bid >= arsCI['BID']:  arsCI['BID'] = bid
                if ask >= arsCI['ASK']:  arsCI['ASK'] = ask
                if last >= arsCI['LAST']: arsCI['LAST'] = last
                print(arsCI)
            elif moneda == 'D':
                if ticker[0][-2:] == '7D': 
                    if bid >= arsCI['BID']:  arsCI['BID'] = bid 
                    if ask >= arsCI['ASK']:  arsCI['ASK'] = ask
                    if last >= arsCI['LAST']: arsCI['LAST'] = last
                    print(arsCI)
                else:    
                    if bid >= mepCI['BID']:  mepCI['BID'] = bid
                    if ask >= mepCI['ASK']:  mepCI['ASK'] = ask
                    if last >= mepCI['LAST']: mepCI['LAST'] = last
                    print(mepCI)
            elif moneda == 'C':
                if bid >= cclCI['BID']:  cclCI['BID'] = bid
                if ask >= cclCI['ASK']:  cclCI['ASK'] = ask
                if last >= cclCI['LAST']: cclCI['LAST'] = last
                print(cclCI)
            else:
                if bid >= arsCI['BID']:  arsCI['BID'] = bid 
                if ask >= arsCI['ASK']:  arsCI['ASK'] = ask
                if last >= arsCI['LAST']: arsCI['LAST'] = last
                print(arsCI)
        
        celda += 1

excelRulo(64)












        