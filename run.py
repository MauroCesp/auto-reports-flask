# -*- encoding: utf-8 -*-
"""
Copyright (c) 2019 - present AppSeed.us
"""

import os
from   flask_migrate import Migrate
from   flask_minify  import Minify
from   sys import exit
#------------------------------------------------------------------------#
# Importo del archivo de configuracion la variable config_dict
# Las variables dentro del diccionario de config_dict son clases que basicamente me crean la base de datos y setean el DEGUB como True
"""
# Load all possible configurations
config_dict = {
    'Production': ProductionConfig,
    'Debug'     : DebugConfig
"""
from apps.config import config_dict
#from flask_ngrok import run_with_ngrok
#------------------------------------------------------------------------#
# Esta es una funcion que importo desde el app/__init__.Copyright
# Esto me inicializa la applicacion y los register_blueprints
# Ademas inicializa la base de datos.
"""
def create_app(config):
    app = Flask(__name__)
    app.config.from_object(config)
    register_extensions(app)
    register_blueprints(app)
    configure_database(app)
    return app
"""
from apps import create_app, db
#------------------------------------------------------------------------#

# WARNING: Don't run with debug turned on in production!
DEBUG = (os.getenv('DEBUG', 'False') == 'True')

# The configuration
get_config_mode = 'Debug' if DEBUG else 'Production'

try:

    # Load the configuration using the default values
    app_config = config_dict[get_config_mode.capitalize()]

except KeyError:
    exit('Error: Invalid <config_mode>. Expected values [Debug, Production] ')



app = create_app(app_config)
Migrate(app, db)

if not DEBUG:
    Minify(app=app, html=True, js=False, cssless=False)

if DEBUG:
    app.logger.info('DEBUG            = ' + str(DEBUG)             )
    app.logger.info('FLASK_ENV        = ' + os.getenv('FLASK_ENV') )
    app.logger.info('Page Compression = ' + 'FALSE' if DEBUG else 'TRUE' )
    app.logger.info('DBMS             = ' + app_config.SQLALCHEMY_DATABASE_URI)
    app.logger.info('ASSETS_ROOT      = ' + app_config.ASSETS_ROOT )

if __name__ == "__main__":
    # app.run(host='0.0.0.0',port=80)
    #app.run(host='0.0.0.0',port=5000)
    app.run()
