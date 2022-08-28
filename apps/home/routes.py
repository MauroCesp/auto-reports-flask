# -*- encoding: utf-8 -*-

import os

"""
Copyright (c) 2019 - present AppSeed.us
"""
######################  BLUE PRINT IMPORT #################################
#Flask Blueprints encapsulate functionality, such as views, templates, and other resources.
# Lo que hago es importar el archivo INIT desde la ruta especificada
# El archivo contiene la inicializacion del objeto de BLUEPRINT
"""blueprint = Blueprint(
    'home_blueprint',
    __name__,
    url_prefix='')
"""
from apps.home import blueprint
#----------------------------------------------------------------------
# Esto es para trabajar con archivos seguros
from werkzeug.utils import secure_filename

##########################  MIS MODELOS  ############################
# Este es el modelo que cree para realizar las transformaciones de los archivos de EXCEL
# Cree un folder MODELS
#from apps.models.classifier import Classifier as cls

from apps.models.journals_formatting import Format as journal

##########################  FLASK LIBRARIES  ############################
# Esto es para poder utilizar modals en la applicación
# Aun que por ahora estoy realizando modals de manera manual
from flask_modals import Modal
#
#Esta librería me permite retieve los template que tengo dentro del folder /home/templates/
# Intalamos la extension de FLASH para poder tirar mensajes rapidos
from flask import Flask, flash, request, redirect, url_for, render_template
from flask_login import login_required
from jinja2 import TemplateNotFound

##########################  IMPORT APP ##################################
# Es super necesario llamar al APP para poder utilizar las libreria propia
# Como cunado necesitamos llamar al directorio raiz de la applicacion a la hora de importar ficheros.
from run import app
##########################  EXCEL #######################################
# Todas estas librerias me permiten ttrabajar con archivos de EXCEL
import openpyxl
from openpyxl import Workbook

##########################  PANDAS #######################################
# Todas estas librerias me permiten ttrabajar con archivos de EXCEL
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
#
#
#	#	#	#	#	#	#	#	#	#	#	#	#	#	#	#	#	#	#
#
#-------------------------------------------------------------------------#
#------------------------  INDEX  ---------------------------------------#
#-------------------------------------------------------------------------#
# Como ya importamos el objeto de BLUEPRINT desde el init_.py y alo puedo utilizar
# El Objeto la llamamos "blueprint" entonces es el mobre que utilizamos en el decorador
@blueprint.route('/index')
@login_required
def index():
    # La ruta que llamamos es apps/templates/home/...
    return render_template('home/index.html', segment='index')
#
#
#-------------------------------------------------------------------------#
#------------------------  FORM REPORT  ----------------------------------#
#-------------------------------------------------------------------------#

# Como ya importamos el objeto de BLUEPRINT desde el init_.py ya lo puedo utilizar
# El Objeto la llamamos "blueprint" entonces es el mobre que utilizamos en el decorador

@blueprint.route('/new-report', methods = ['GET','POST'])
@login_required
def entitlement():
    # Como tengo tres tipos de reportes diferentes - journals,books,databases-  creé en la vista tres formularios
    # Necesito decirle a la ruta cual formulario deseo utilizar en cada caso.
    # Para ello en el valor VALUE del formulario especifique un identificador que utilizo aqui
	if request.method == 'POST':

# --------- JOURNALS --------------
        # Necesito comparar el nombre del formulario que me llaga por la request con el VALUE de ese formulario
		if request.form.get('journals') == 'journals':
            # Con el nombre que le dimos al archivo en el VIEW lo trabajamos aqui
			f = request.files['file']

            # ----------------------
            # Salvamos el archivo de manera segura
			filename = secure_filename(f.filename)

            # Con esta linea guardo el archivo en al ruta que definí en el INIT file -- apps/uploads/tempData
            # Aqui es donde llegan los archivos por defecto, de aquí los puedo mover a cualquier folder
            # Pero primero es importante poder encontrarlos
			f.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

            # Asi obtenemos el nombre del archivo
            #------------------------------------
            #		NEW COLUMN ----> Month-Year
            #------------------------------------
			rows, cols, newRows, newCols = journal.new_column(filename)

			#journal.set_table()
			journal.test1()

			return render_template('home/webstats/ui-newWSReport.html',rows = rows, cols = cols, newRows = newRows, newCols = newCols )

# --------- BOOKS --------------
        # Necesito comparar el nombre del formulario que me llaga por la request con el VALUE de ese formulario
		elif request.form.get('books') == 'books':
            # Esta ruta  es solo para probar
			return render_template('home/charts.html')
# --------- DB --------------
        # Si no es gallo es gallina
        # Como solo tengo tres formulario en la vista pues por defecto solo queda este para las bases de datos.
		else:
            # Esta ruta  es solo para probar
			return render_template('home/ui-buttons.html')
	else:
        # Esta es la ruta a la que accedo cuando se solicita la pagina
        # Es decir con el GET aqui llego
        # Una vez adentro es que realizo el post a esta ruta ENTITLEMENT
		return render_template('home/webstats/ui-webStatsReport.html', segment='index')


#-------------------------------------------------------------------------#
#------------------------  EXPORT  ---------------------------------------#
#-------------------------------------------------------------------------#

@blueprint.route("/export/", methods=['GET'])
def export_records():
    return excel.make_response_from_array([[1,2], [3, 4]], "xls",
                                          file_name="export_data")
#-------------------------------------------------------------------------#
#------------------------  ROUTES  ---------------------------------------#
#-------------------------------------------------------------------------#
# Como ya importamos el objeto de BLUEPRINT desde el init_.py y alo puedo utilizar
# El Objeto la llamamos "blueprint" entonces es el mobre que utilizamos en el decorador
# One of the most used decorators is route. It allows you to associate a view function to a URL route.
#
# When you bind a function with the help of the @blueprint.route decorator,
# The blueprint will record the intention of registering the function ROUTE_TEMPLATE on the application when it’s later registered.
# Additionally it will prefix the endpoint of the function with the name of the blueprint which was given to the Blueprint constructor (in this case also HOME_BLUEPRINT).
# The blueprint’s name does not modify the URL, only the endpoint.
#
#
# Entonces aqui lo que hacemos es tener una solo ruta para todos los archivos destro de la carpeta de home
# Solo tenemos que llamar a esta funcion <def route_template> dentro del documento HTML y pasarle com parametro el nombre del archivo entre comillas simples
# para que sea leido como string.
#
#
# Esta es un abuena forma de escalar proyectos y que no se nos convierta el codigo en algo grande y poco manejable
# En caso de que tengamos que pasarle informacion a cada pagina, solo tenemos que agregar otra variable dentro del HTML y l afuncion
@blueprint.route('/<template>')
@login_required
# En este caso el paso 'template' que es el nombre del fichero que quiero abrir.
# Tambien le paso un parametro adicional que es 'data'. Representa cualquier dato adicional que le tenga que pasar a la vista
# Se pueden pasar cuantos parametros sean necesarios pero es necesario sacarlos del request
# Usualmente se declara el parametro deltro de la funcion, pero como es un afuncion compartida vamos a utilizar el GET.ARGS.GET para sacarlos
def route_template(template):
    try:
        # Como solo indicamos el nombre del archivo es necesario agregarle el HTML para que lo podamos encontrar
        if not template.endswith('.html'):
            template += '.html'

        # Detect the current page
        # Llamamos a la funcion GET_SEGMENT mas abajo para obtener el segment que en este caso sera 'index'
        segment = get_segment(request)

        #------------ GET ARGUMENTS -------------#
        # Debemos de sacer los argumentos de esta forma para poderlos enviar a la ruta correspondiente
        # El nombre del argumento es igual pero en el HTML le asignamos valores diferentes de acuerdo a la vista que deseo mostrar
        # Pueden haber cuantos parametros queramos pero simepre los tenemos que sacar del request antes de enviarlos a la ruta.
        arg = request.args.get('param1')

        # Serve the file (if exists) from app/templates/home/FILE.html
        # Además le paso el parámetro que me llega por el request.
        # La idea es que este parametro cambie en cada ruta asi podemos realizar la operaciones en cada view.
        return render_template("home/" + template, segment=segment, arg = arg)

    except TemplateNotFound:
        return render_template('home/page-404.html'), 404

    except:
        return render_template('home/page-500.html'), 500

#-------------------------------------------------------------------------#
#------------------------  SEGMENT  --------------------------------------#
#-------------------------------------------------------------------------#
# Helper - Extract current page name from request
def get_segment(request):
    try:
        segment = request.path.split('/')[-1]

        if segment == '':
            segment = 'index'

        return segment
    except:
        return None
