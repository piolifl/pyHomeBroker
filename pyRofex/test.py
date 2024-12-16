'''import datetime
import time
import pyRofex

pyRofex._set_environment_parameter("url", "https://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)
pyRofex._set_environment_parameter("ws", "wss://api.veta.xoms.com.ar/", pyRofex.Environment.LIVE)
pyRofex.initialize(user="20263866623", password="Bordame01!", account="47352", environment=pyRofex.Environment.LIVE)
    '''

import pycocos
import pandas as pd

# Inicializar la conexión a Cocos
cocos = pycocos.Cocos()

# Definir los instrumentos que te interesan
instrumentos = ['ALUA', 'BBAR', 'GGAL']

# Obtener los datos de los instrumentos
datos = cocos.get_data(instrumentos)

# Convertir los datos a un DataFrame de pandas
df = pd.DataFrame(datos)

# Exportar los datos a un archivo Excel
df.to_excel('datos_acciones.xlsx', index=False)