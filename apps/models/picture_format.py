try:

    from pathlib import Path
    from datetime import datetime, date

    #----------  IMPORT APP -----------------------------------
    # Es super necesario llamar al APP para poder utilizar las libreria propia
    # Como cunado necesitamos llamar al directorio raiz de la applicacion a la hora de importar ficheros.
    from run import app
    import sys
    import time
    from PIL import Image
    import os
# Utilizo la excepción específica para módulos
except ModuleNotFoundError as err:
    print('Opssss... Looks like there is an error importing the package', err)
################################################################################

class Picture():
    #------------- CONSTRUCTOR ------------------
    # Inicializo el constructor
    # Cada vez que cree un objeto de tipo report se inicializan estos atributos
    # Se repiten en todos los metodos aunque cambia el valores# El valor lo cambio directo desde el archivo de rutas
    def __init__(self):
        # Defino la ruta donde deseo guardar las imagenes desde el principio
        self.SAVE_PATH = os.path.abspath('apps/static/assets/img/company-logos/')

    def resize(self,f,h1):

        # Recibo la imagen como parametro y abro el objeto
        with Image.open(f) as im:

            # esta funcion me permite rescalar imagenes sin saber su tamaño
            # Defino los topes que no puede pasar la imagenes
            # La imagen se ajusta a esos limites.
            im.thumbnail((150, 150))
            # Indico donde deseo guardar la imagen
            # Recibo por parametro el nombre de la imagen sin extension y aqui le digo que todas sean PNG
            im.save(self.SAVE_PATH + '\\' + h1 + '.png', format='png')

            #x = im.format
            #y = im.size
            #z = im.mode
            Image.Image.close(self)

        return True
