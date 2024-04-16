from pyhomebroker import HomeBroker

def example_online():

    broker=265
    dni='26968339'
    user='bordame'
    password='Bordame02'

    hb = HomeBroker(int(broker), 
        on_open=on_open, 
        on_personal_portfolio=on_personal_portfolio, 
        on_error=on_error, 
        on_close=on_close)
        
    hb.auth.login(dni=dni, user=user, password=password, raise_exception=True)
    
    hb.online.connect()
    hb.online.subscribe_personal_portfolio()

    input('Press Enter to Disconnect...\n')

    hb.online.unsubscribe_personal_portfolio()
    hb.online.disconnect()

def on_open(online):
    
    print('=================== CONNECTION OPENED ====================')

def on_personal_portfolio(online, portfolio_quotes):
    
    print('------------------- Personal Portfolio -------------------')
    print(portfolio_quotes)

    
def on_error(online, exception, connection_lost):
    
    print('@@@@@@@@@@@@@@@@@@@@@@@@@ Error @@@@@@@@@@@@@@@@@@@@@@@@@@')
    print(exception)

def on_close(online):

    print('=================== CONNECTION CLOSED ====================')

if __name__ == '__main__':
    example_online()
  
#[ ]><   \n