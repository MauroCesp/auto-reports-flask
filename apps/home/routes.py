# -*- coding: utf-8 -*-
"""
Created on Thu Mar 17 17:50:04 2022
@author: Mauro
"""
# Lo primero que hago es utilizar las excepciones al importar los módulos
# Así puedo controlar los errores que surjan al importar

    # Voy a utilizar dos formas de importar paquetes
    # Primero importo solo lo que necesito directamente de cada sub-package
    # En este caso importo los modulos de cada uno para utilizar sus funciones
from apps.home import blueprint
    #----------------------------------------------------------------------
    # Esto es para trabajar con archivos seguros
from werkzeug.utils import secure_filename
    #----------  MIS MODELOS  --------------------------------------------
    # Este es el modelo que cree para realizar las transformaciones de los archivos de EXCEL
    # Cree un folder MODELS
    # Con esta clase vamos a crear los reportes
from apps.models.journal_format import Journal
from apps.models.book_format import Book
from apps.models.database_format import Database
from apps.models.picture_format import Picture

    #----------------- FLASK LIBRARIES  ------------------------------------
    # Esto es para poder utilizar modals en la applicación
    # Aun que por ahora estoy realizando modals de manera manual
from flask_modals import Modal
    #Esta librería me permite retieve los template que tengo dentro del folder /home/templates/
    # Intalamos la extension de FLASH para poder tirar mensajes rapidos
from flask import Flask, flash, request, redirect, url_for, render_template, send_file, session
from flask_login import login_required
from jinja2 import TemplateNotFound
    #----------------------  IMPORT APP -----------------------------------------
    # Es super necesario llamar al APP para poder utilizar las libreria propia
    # Como cunado necesitamos llamar al directorio raiz de la applicacion a la hora de importar ficheros.
from run import app
    # Ahora utilizo directamente el import para llamar a un modulo que no pertenece a ningún sub-paquete
    # El modulo se encuetra en el mismo nivel que los sub-paquetes
    # Tambien quize practicar un poco el Alias
    #--------------------------  EXCEL -----------------------------------------
from openpyxl import Workbook
    # Todas estas librerias me permiten ttrabajar con archivos de EXCEL
import openpyxl

import subprocess as sp

from openpyxl import load_workbook
    #----------------------------  PANDAS --------------------------------------
    # Todas estas librerias me permiten ttrabajar con archivos de EXCEL
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import time
import re

import json


#-----------------  BLUE PRINT IMPORT --------------------------------------------
#Flask Blueprints encapsulate functionality, such as views, templates, and other resources.
# Lo que hago es importar el archivo INIT desde la ruta especificada
# El archivo contiene la inicializacion del objeto de BLUEPRINT
"""blueprint = Blueprint(
    'home_blueprint',
    __name__,
    url_prefix='')
"""

#-------------------------------------------------------------------------#
#------------------------  INDEX  ---------------------------------------#
#-------------------------------------------------------------------------#
# Como ya importamos el objeto de BLUEPRINT desde el init_.py y alo puedo utilizar
# El Objeto la llamamos "blueprint" entonces es el mobre que utilizamos en el decorador
@login_required
@blueprint.route('/index')

def index():
    # La ruta que llamamos es apps/templates/home/...
    flash(' welcome','success')

    with open("apps/counter_webstats.txt") as f:
        ws = sum(int(line) for line in f)

    with open("apps/counter_atypon.txt") as f:
        at = sum(int(line) for line in f)

    with open("apps/counter_highwire.txt") as f:
        hw = sum(int(line) for line in f)

    return render_template('home/index.html',ws=ws,at = at,hw = hw )

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
# --------- JOURNALS ----------------------------------------------------------------------------------------------------
        # Necesito comparar el nombre del formulario que me llaga por la request con el VALUE de ese formulario
        if request.form.get('journals') == 'journals':

            KEY = 'journals'
            # Ahora comprobamos que se esté cargardo un archivo válido
            # Con el nombre que le dimos al archivo en el VIEW lo trabajamos aq
            type_graphic = request.form.get('comp_select')

            #---------------COLUMNS PARAMETERS -------------------
            session_select = request.form.get('Sessions')
            SR_select = request.form.get('Searches_Regular')
            TII_select = request.form.get('Total_Item_Investigations')
            UII_select = request.form.get('Unique_Item_Investigations')
            TIR_select = request.form.get('Total_Item_Requests')
            UIR_select = request.form.get('Unique_Item_Requests')




            #---------------TAB PARAMETERS -------------------
            Monthly_Trend = request.form.get('Monthly_Trend')
            Yearly_Comparison = request.form.get('Yearly_Comparison')
            Monthly_Trend1 = request.form.get('Monthly_Trend1')


            # Aqui traemos del form si quiere utilizar grid o no
            # En el for el VALUE es True cunado se seleccionar# SI no se selecciona el valor el NONE
            use_grid = request.form.get('grid')
            #no_logo = request.form.get('no_logo')

            image = request.form.get('image')

            # Ahora revisamos que se haya carga un archivo correctamente.
            if request.files['file']:

                # Aqui trabajo con el NAME
                f = request.files['file']
                # Esto es el timepo inical para medir cuanto tarda el programa en ejecutarse
                st = time.time()
                # Salvamos el archivo de manera segura
                SOURCE_FILE = secure_filename(f.filename)
                # Con esta linea guardo el archivo en al ruta que definí en el INIT file -- apps/uploads/tempData
                # Aqui es donde llegan los archivos por defecto, de aquí los puedo mover a cualquier folder
                # Pero primero es importante poder encontrarlos
                f.save(os.path.join(app.config['UPLOAD_FOLDER'], SOURCE_FILE))

                # Como envio los parametro al constructor los toma directamente del proyecto en el que estamos trabajando.
                FINAL_SAVE_PATH = 'apps/uploads/tempData/webstats/journals/'
                TEMP_SAVE_PATH = 'apps/uploads/tempData/'


                return render_template('home/loading.html',
                                        KEY = KEY,
                                        use_grid = use_grid,
                                        image = image,
                                        type_graphic = type_graphic,
                                        SOURCE_FILE = SOURCE_FILE,
                                        FINAL_SAVE_PATH = FINAL_SAVE_PATH,
                                        TEMP_SAVE_PATH = TEMP_SAVE_PATH,
                                        session_select = session_select,
                                        SR_select = SR_select,
                                        TII_select = TII_select,
                                        TIR_select = TIR_select,
                                        UIR_select = UIR_select,
                                        UII_select = UII_select,
                                        Monthly_Trend = Monthly_Trend,
                                        Yearly_Comparison = Yearly_Comparison,
                                        Monthly_Trend1 = Monthly_Trend1,
                                        )


            else:
                flash(' Please upload a valid file !', 'warning')
                return render_template('home/webstats/ui-webStatsReport.html')
# --------- BOOKS ----------------------------------------------------------------------------------------------------
        # Necesito comparar el nombre del formulario que me llaga por la request con el VALUE de ese formulario
# --------- BOOK ----------------------------------------------------------------------------------------------------
        elif request.form.get('books') == 'books':
            KEY = 'books'
            # Ahora comprobamos que se esté cargardo un archivo válido
            # Con el nombre que le dimos al archivo en el VIEW lo trabajamos aq
            type_graphic = request.form.get('comp_select')


            #---------------COLUMNS PARAMETERS -------------------
            session_select = request.form.get('Sessions')
            SR_select = request.form.get('Searches_Regular')
            TII_select = request.form.get('Total_Item_Investigations')
            TIR_select = request.form.get('Total_Item_Requests')
            UIR_select = request.form.get('Unique_Item_Requests')
            UII_select = request.form.get('Unique_Item_Investigations')

            #---------------TAB PARAMETERS -------------------
            Monthly_Trend = request.form.get('Monthly_Trend')
            Yearly_Comparison = request.form.get('Yearly_Comparison')
            Monthly_Trend1 = request.form.get('Monthly_Trend1')


            # Aqui traemos del form si quiere utilizar grid o no
            # En el for el VALUE es True cunado se seleccionar# SI no se selecciona el valor el NONE
            use_grid = request.form.get('grid')
            #no_logo = request.form.get('no_logo')
            image = request.form.get('image')


            if request.files['file']:

                # Aqui trabajo con el NAME
                f = request.files['file']
                # Esto es el timepo inical para medir cuanto tarda el programa en ejecutarse
                st = time.time()
                # Salvamos el archivo de manera segura
                SOURCE_FILE = secure_filename(f.filename)
                # Con esta linea guardo el archivo en al ruta que definí en el INIT file -- apps/uploads/tempData
                # Aqui es donde llegan los archivos por defecto, de aquí los puedo mover a cualquier folder
                # Pero primero es importante poder encontrarlos
                f.save(os.path.join(app.config['UPLOAD_FOLDER'], SOURCE_FILE))

                # Como envio los parametro al constructor los toma directamente del proyecto en el que estamos trabajando.
                FINAL_SAVE_PATH = 'apps/uploads/tempData/webstats/books/'
                TEMP_SAVE_PATH = 'apps/uploads/tempData/'

                return render_template('home/loading.html',
                                        KEY = KEY,
                                        use_grid = use_grid,
                                        image = image,
                                        type_graphic = type_graphic,
                                        SOURCE_FILE = SOURCE_FILE,
                                        FINAL_SAVE_PATH = FINAL_SAVE_PATH,
                                        TEMP_SAVE_PATH = TEMP_SAVE_PATH,
                                        session_select = session_select,
                                        SR_select = SR_select,
                                        TII_select = TII_select,
                                        TIR_select = TIR_select,
                                        UIR_select = UIR_select,
                                        UII_select = UII_select,
                                        Monthly_Trend = Monthly_Trend,
                                        Yearly_Comparison = Yearly_Comparison,
                                        Monthly_Trend1 = Monthly_Trend1

                                        )
            else:
                flash(' Please upload a valid file !', 'warning')
                return render_template('home/webstats/ui-webStatsReport.html')
# --------- DB ----------------------------------------------------------------------------------------------------
        # Como solo tengo tres formulario en la vista pues por defecto solo queda este para las bases de datos.
        elif request.form.get('databases') == 'databases':
            KEY = 'databases'
            # Ahora comprobamos que se esté cargardo un archivo válido
            # Con el nombre que le dimos al archivo en el VIEW lo trabajamos aq
            type_graphic = request.form.get('comp_select')


            #---------------COLUMNS PARAMETERS -------------------
            session_select = request.form.get('Sessions')
            SR_select = request.form.get('Searches_Regular')
            TII_select = request.form.get('Total_Item_Investigations')
            TIR_select = request.form.get('Total_Item_Requests')
            UIR_select = request.form.get('Unique_Item_Requests')
            UII_select = request.form.get('Unique_Item_Investigations')
            #---------------TAB PARAMETERS -------------------
            Monthly_Trend = request.form.get('Monthly_Trend')
            Yearly_Comparison = request.form.get('Yearly_Comparison')
            Monthly_Trend1 = request.form.get('Monthly_Trend1')


            # Aqui traemos del form si quiere utilizar grid o no
            # En el for el VALUE es True cunado se seleccionar# SI no se selecciona el valor el NONE
            use_grid = request.form.get('grid')
            #no_logo = request.form.get('no_logo')
            image = request.form.get('image')

            if request.files['file']:

                # Aqui trabajo con el NAME
                f = request.files['file']
                # Esto es el timepo inical para medir cuanto tarda el programa en ejecutarse
                st = time.time()
                # Salvamos el archivo de manera segura
                SOURCE_FILE = secure_filename(f.filename)
                # Con esta linea guardo el archivo en al ruta que definí en el INIT file -- apps/uploads/tempData
                # Aqui es donde llegan los archivos por defecto, de aquí los puedo mover a cualquier folder
                # Pero primero es importante poder encontrarlos
                f.save(os.path.join(app.config['UPLOAD_FOLDER'], SOURCE_FILE))

                # Como envio los parametro al constructor los toma directamente del proyecto en el que estamos trabajando.
                FINAL_SAVE_PATH = 'apps/uploads/tempData/webstats/databases/'
                TEMP_SAVE_PATH = 'apps/uploads/tempData/'

                return render_template('home/loading.html',
                                        KEY = KEY,
                                        use_grid = use_grid,
                                        image = image,
                                        type_graphic = type_graphic,
                                        SOURCE_FILE = SOURCE_FILE,
                                        FINAL_SAVE_PATH = FINAL_SAVE_PATH,
                                        TEMP_SAVE_PATH = TEMP_SAVE_PATH,
                                        session_select = session_select,
                                        SR_select = SR_select,
                                        TII_select = TII_select,
                                        TIR_select = TIR_select,
                                        UIR_select = UIR_select,
                                        UII_select = UII_select,
                                        Monthly_Trend = Monthly_Trend,
                                        Yearly_Comparison = Yearly_Comparison,
                                        Monthly_Trend1 = Monthly_Trend1

                                        )
            else:
                flash(' Please upload a valid file !', 'warning')
                return render_template('home/webstats/ui-webStatsReport.html')
# --------- ELSE ----------------------------------------------------------------------------------------------------
        else:
            flash('Please upload a valid file!', 'warning')
            return render_template('home/webstats/ui-webStatsReport.html')

    #------- GET -------------
    else:
        # Esta es la ruta a la que accedo cuando se solicita la pagina
        # Es decir con el GET aqui llego
        # Una vez adentro es que realizo el post a esta ruta ENTITLEMENT
        return render_template('home/webstats/ui-webStatsReport.html', segment='index')

#-------------------------------------------------------------------------#
#------------------------  UPLOAD PICTURE  ----------------------------------#
#-------------------------------------------------------------------------#
@blueprint.route('/upload-picture', methods = ['GET','POST'])
@login_required
def uploadPicture():
    # Como tengo tres tipos de reportes diferentes - journals,books,databases-  creé en la vista tres formularios
    # Necesito decirle a la ruta cual formulario deseo utilizar en cada caso.
    # Para ello en el valor VALUE del formulario especifique un identificador que utilizo aqui
    if request.method == 'POST':

# --------- JOURNALS --------------
        # Necesito comparar el nombre del formulario que me llaga por la request con el VALUE de ese formulario
        if request.form.get('upload') == 'upload':
            if request.files['picture']:
                # Con el nombre que le dimos al archivo en el VIEW lo trabajamos aqui
                f = request.files['picture']
                #--------- get picture name
                # Esto lo utilizamos para obtener el nombre del rchivo que llega por el request.
                # Primero lo buscamos en los headers
                d = f.headers['content-disposition']
                # Despues le decimos que nos extraiga el nombre.
                # Aqui todavia tenemos el nombre con la extension y se la tenemos que quitar para que pueda encontrar el nombre.
                h = re.findall("filename=(.+)", d)[0]
                # Definimos lo que queremos usar como separador
                separator = '.'
                # Ahora le decimos que nos elimine todo despues del punto.
                h1 = h.split(separator, 1)[0]
                h1 = h1.replace('"', '')
                #---------------------
                # CREO EL OBJETO
                #---------------------
                picture  = Picture()
                # Utilizo la función de resize para que todas las imagenes sean del mismo tamaño.
                # Le pasa el nombre sin la extension para pasarlas todas a PNG
                picture.resize(f, h1)
                flash(h1 + ' '+ ' has been properly uploaded','success')
                return render_template('home/webstats/ui-uploadpicture.html')
            else:
                flash(' Please upload a valid file !', 'warning')
                return render_template('home/webstats/ui-uploadpicture.html')
        else:
            pass
    else:
        # Esta es la ruta a la que accedo cuando se solicita la pagina
        # Es decir con el GET aqui llego
        # Una vez adentro es que realizo el post a esta ruta ENTITLEMENT
        return render_template('home/webstats/ui-uploadpicture.html', segment='index')
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

#-------------------------------------------------------------------------#
#------------------------  DOWNLOAD  --------------------------------------#
#-------------------------------------------------------------------------#
# Helper - Extract current page name from request
@blueprint.route('/download/', methods=('GET', 'POST'))
@login_required
# recibo la variable que me pasa de la vista
def download():

    # Hay que llamar al field que se queire utilizar con el NAME del formulario en el views# Asi es como conectamos la variabel y la jalamos.
    # Funciona como KEY:VALUE ------> NAME:VALUE en la vista
    type = str(request.form['type'])
    user = request.form['user']
    institution = request.form['institution']
    path = str(request.form['find_path'])


    # Aqui le vuelvo a quita los '_' al path para poder encontrarlos
    # Es que com oel nombre lleva espacio me haci escpe y no me jalaba el path
    #find_path = path.replace('_', ' ')

    print(type)
    print(user)
    print(institution)
    print(path)

    # La ruta la programammos con el final path para poder dejar limpio el temp folder y evitar los errores de repeticion.
    ruta1 = "uploads\\finalData\\webstats\\"
    ruta2 = str(type + "\\")

    ruta_final = str(ruta1+ruta2+path)
    return send_file(ruta_final, as_attachment=True)

#-------------------------------------------------------------------------#
#------------------------  Processing  --------------------------------------#
#-------------------------------------------------------------------------#
@blueprint.route("/processing")
@login_required
def processing():

    use_grid = request.args.get('use_grid')
    image = request.args.get('image')
    type_graphic = request.args.get('type_graphic')
    SOURCE_FILE = request.args.get('SOURCE_FILE')
    FINAL_SAVE_PATH = request.args.get('FINAL_SAVE_PATH')
    TEMP_SAVE_PATH = request.args.get('TEMP_SAVE_PATH')
    KEY = request.args.get('KEY')
    session_select = request.args.get('session_select')
    SR_select = request.args.get('SR_select')
    TII_select = request.args.get('TII_select')
    TIR_select = request.args.get('TIR_select')
    UIR_select = request.args.get('UIR_select')
    UII_select = request.args.get('UII_select')
    Monthly_Trend = request.args.get('Monthly_Trend')
    Yearly_Comparison = request.args.get('Yearly_Comparison')
    Monthly_Trend1 = request.args.get('Monthly_Trend1')

#--------------------------------------------------------------------
#-------------  JOURNALS  -------------------------------------------
    if KEY == 'journals':
        #------------------------------------
        #		CREATE OBJECT
        #------------------------------------
        #---------------------------------
        # Creamos el objeto
        # Es necesario pasarle los parametros de inicializacion del constructor
        journal = Journal(TEMP_SAVE_PATH, FINAL_SAVE_PATH, SOURCE_FILE, type_graphic, use_grid, session_select,TII_select, TIR_select, UIR_select , Monthly_Trend, Yearly_Comparison, Monthly_Trend1, SR_select,image, UII_select)

        #---- NO UTILIZO IMAGEN ----
        if image == 'True':
            institutionIs = journal.new_column()
            journal.set_table()
            journal.pivot_montly_trend()
            journal.pivot_stats()
            extra_path, col_end = journal.extra()
            journal.set_table_extra(extra_path, col_end)
            if Yearly_Comparison == 'Yearly_Comparison':
                journal.yearly_comp()
            else:
                pass

            average_yr, count_years,count_months,sum_session_yr ,sum_TIR_yr,sum_TII_yr,sum_UIR_yr,sum_UII_yr,sessions,TIR,TII,UIR,UII,sum_session_month ,sum_TIR_month,sum_TII_month,sum_UIR_month,sum_UII_month,sessions_m,TIR_m,TII_m,UIR_m,UII_m = journal.webGraph()

            #------------------------------------------
            #-------- COUNTER -------------------------
            #------------------------------------------
            # Esta es la funcion counter que me muestra los gráficos al principio.
            journal.counter(KEY)

            institution_name, institution_init_year, institution_end_year,find_path, type, pic_path  = journal.final()
            time.sleep(15)
            flash('Report successfully generated','success')
            journal.clean_folder()


            return render_template('home/webstats/ui-newWSReport.html',
                                institution_name = institution_name,
                                institution_init_year = institution_init_year,
                                institution_end_year = institution_end_year,
                                type = type,
                                find_path = find_path,
                                pic_path =pic_path,

                                average_yr = average_yr,
                                count_years = count_years,
                                count_months = count_months,
                                sum_session_yr  = sum_session_yr,
                                sum_TIR_yr  = sum_TIR_yr,
                                sum_TII_yr  = sum_TII_yr,
                                sum_UIR_yr  = sum_UIR_yr,
                                sum_UII_yr  = sum_UII_yr,
                                sessions = sessions,
                                TIR=TIR,
                                TII=TII,
                                UIR=UIR,
                                UII=UII,

                                sum_session_month  = sum_session_month,
                                sum_TIR_month = sum_TIR_month,
                                sum_TII_month  = sum_TII_month,
                                sum_UIR_month  = sum_UIR_month,
                                sum_UII_month  = sum_UII_month,
                                sessions_m = sessions_m,
                                TIR_m=TIR_m,
                                TII_m=TII_m,
                                UIR_m=UIR_m,
                                UII_m=UII_m,
                                Yearly_Comparison=Yearly_Comparison
                                )
        # ---- UTILIZO IMAGEN
        else:
            # Esto lo hacemos para poder escojer la imagen que necesitamos para el reporte
            # Resibimos del modelo el parametro que vamos a evaluar
            isExist = journal.set_image()
            # Asi obtenemos el nombre del archivo


            #------- NEW COLUMN ----> Month-Year
            # Sacamos el nombre de la institucion para que podamos revisar si tiene foto o no
            institutionIs = journal.new_column()

            # Aqui preguntamos si la imagen la tenemos grabada en el sistema
            # La funcion me regresa true or false
            if isExist == True:
                #------------------------------------
                #		CHECK FOR IMAGE TO CONTINUE
                #------------------------------------
                #------- Table format
                journal.set_table()
                #------ Montly Trend
                # Para esto ya necesitamos utilizar absolute path
                journal.pivot_montly_trend()
                #----- Stats General report
                journal.pivot_stats()

        #-------------------------------------------------------------------------------------
                # Aqui saco el path y la columna final donde terminar
                extra_path, col_end = journal.extra()

                journal.set_table_extra(extra_path, col_end)
                #database.extra_chart(path)
                # Pillo todas la variables que necesito pasar a la vista

                ##-------- Yealry comparison --------------
                if Yearly_Comparison == 'Yearly_Comparison':
                    journal.yearly_comp()
                else:
                    pass

                average_yr, count_years,count_months,sum_session_yr ,sum_TIR_yr,sum_TII_yr,sum_UIR_yr,sum_UII_yr,sessions,TIR,TII,UIR,UII,sum_session_month ,sum_TIR_month,sum_TII_month,sum_UIR_month,sum_UII_month,sessions_m,TIR_m,TII_m,UIR_m,UII_m = journal.webGraph()

                #------------------------------------------
                #-------- COUNTER -------------------------
                #------------------------------------------
                # Esta es la funcion counter que me muestra los gráficos al principio.
                journal.counter(KEY)

                institution_name, institution_init_year, institution_end_year,find_path, type, pic_path  = journal.final()
                time.sleep(15)
                flash('Report successfully generated','success')
                journal.clean_folder()


                #----------------------------------------------------------
                #----------------------------------------------------------
                #----------------------------------------------------------
                return render_template('home/webstats/ui-newWSReport.html',
                                    institution_name = institution_name,
                                    institution_init_year = institution_init_year,
                                    institution_end_year = institution_end_year,
                                    type = type,
                                    find_path = find_path,
                                    pic_path =pic_path,

                                    average_yr = average_yr,
                                    count_years = count_years,
                                    count_months = count_months,
                                    sum_session_yr  = sum_session_yr,
                                    sum_TIR_yr  = sum_TIR_yr,
                                    sum_TII_yr  = sum_TII_yr,
                                    sum_UIR_yr  = sum_UIR_yr,
                                    sum_UII_yr  = sum_UII_yr,
                                    sessions = sessions,
                                    TIR=TIR,
                                    TII=TII,
                                    UIR=UIR,
                                    UII=UII,

                                    sum_session_month  = sum_session_month,
                                    sum_TIR_month = sum_TIR_month,
                                    sum_TII_month  = sum_TII_month,
                                    sum_UIR_month  = sum_UIR_month,
                                    sum_UII_month  = sum_UII_month,
                                    sessions_m = sessions_m,
                                    TIR_m=TIR_m,
                                    TII_m=TII_m,
                                    UIR_m=UIR_m,
                                    UII_m=UII_m,
                                    Yearly_Comparison=Yearly_Comparison
                                    )

            else:
                journal.clean_folder()
                flash(institutionIs + ' '+ ' still don´t have an image in the system, please upload one here')
                return render_template('home/webstats/ui-webStatsReport.html')
                #----------------------------------------------------------
                #----------------------------------------------------------
                #----------------------------------------------------------



#--------------------------------------------------------------------
#-------------  BOOKS  -------------------------------------------
    elif KEY=='books':
        #------------------------------------
        #		CREATE OBJECT
        #------------------------------------
        #---------------------------------
        # Creamos el objeto
        # Es necesario pasarle los parametros de inicializacion del constructor
        book = Book(TEMP_SAVE_PATH, FINAL_SAVE_PATH, SOURCE_FILE, type_graphic, use_grid, session_select,TII_select, TIR_select, UIR_select , Monthly_Trend, Yearly_Comparison, Monthly_Trend1, SR_select,image, UII_select)

        #---- NO UTILIZO IMAGEN ----
        if image == 'True':
            institutionIs = book.new_column()
            book.set_table()
            book.pivot_montly_trend()
            book.pivot_stats()
            extra_path, col_end = book.extra()
            book.set_table_extra(extra_path, col_end)
            if Yearly_Comparison == 'Yearly_Comparison':
                book.yearly_comp()
            else:
                pass

            average_yr, count_years,count_months,sum_session_yr ,sum_TIR_yr,sum_TII_yr,sum_UIR_yr,sum_UII_yr,sessions,TIR,TII,UIR,UII,sum_session_month ,sum_TIR_month,sum_TII_month,sum_UIR_month,sum_UII_month,sessions_m,TIR_m,TII_m,UIR_m,UII_m = book.webGraph()
            #------------------------------------------
            #-------- COUNTER -------------------------
            #------------------------------------------
            # Esta es la funcion counter que me muestra los gráficos al principio.
            book.counter(KEY)

            institution_name, institution_init_year, institution_end_year,find_path, type, pic_path  = book.final()
            time.sleep(15)
            flash('Report successfully generated','success')
            book.clean_folder()

            return render_template('home/webstats/ui-newWSReport.html',
                                institution_name = institution_name,
                                institution_init_year = institution_init_year,
                                institution_end_year = institution_end_year,
                                type = type,
                                find_path = find_path,
                                pic_path =pic_path,

                                average_yr = average_yr,
                                count_years = count_years,
                                count_months = count_months,
                                sum_session_yr  = sum_session_yr,
                                sum_TIR_yr  = sum_TIR_yr,
                                sum_TII_yr  = sum_TII_yr,
                                sum_UIR_yr  = sum_UIR_yr,
                                sum_UII_yr  = sum_UII_yr,
                                sessions = sessions,
                                TIR=TIR,
                                TII=TII,
                                UIR=UIR,
                                UII=UII,

                                sum_session_month  = sum_session_month,
                                sum_TIR_month = sum_TIR_month,
                                sum_TII_month  = sum_TII_month,
                                sum_UIR_month  = sum_UIR_month,
                                sum_UII_month  = sum_UII_month,
                                sessions_m = sessions_m,
                                TIR_m=TIR_m,
                                TII_m=TII_m,
                                UIR_m=UIR_m,
                                UII_m=UII_m,
                                Yearly_Comparison=Yearly_Comparison
                                )
        # ---- UTILIZO IMAGEN
        else:
            # Esto lo hacemos para poder escojer la imagen que necesitamos para el reporte
            # Resibimos del modelo el parametro que vamos a evaluar
            isExist = book.set_image()
            # Asi obtenemos el nombre del archivo
            #------- NEW COLUMN ----> Month-Year
            # Sacamos el nombre de la institucion para que podamos revisar si tiene foto o no
            institutionIs = book.new_column()

            if isExist == True:


                #------------------------------------
                #		CHECK FOR IMAGE TO CONTINUE
                #------------------------------------
                #------- Table format
                book.set_table()
                #------ Montly Trend
                # Para esto ya necesitamos utilizar absolute path
                book.pivot_montly_trend()
                #----- Stats General report
                book.pivot_stats()

        #-------------------------------------------------------------------------------------
                # Aqui saco el path y la columna final donde terminar
                extra_path, col_end = book.extra()

                book.set_table_extra(extra_path, col_end)
                #database.extra_chart(path)
                # Pillo todas la variables que necesito pasar a la vista

                ##-------- Yealry comparison --------------
                if Yearly_Comparison == 'Yearly_Comparison':
                    book.yearly_comp()
                else:
                    pass

                average_yr, count_years,count_months,sum_session_yr ,sum_TIR_yr,sum_TII_yr,sum_UIR_yr,sum_UII_yr,sessions,TIR,TII,UIR,UII,sum_session_month ,sum_TIR_month,sum_TII_month,sum_UIR_month,sum_UII_month,sessions_m,TIR_m,TII_m,UIR_m,UII_m = book.webGraph()

                #------------------------------------------
                #-------- COUNTER -------------------------
                #------------------------------------------
                # Esta es la funcion counter que me muestra los gráficos al principio.
                book.counter(KEY)

                institution_name, institution_init_year, institution_end_year,find_path, type, pic_path  = book.final()

                # Seudo loading para que aparente estar cargando
                time.sleep(15)
                flash('Report successfully generated','success')

                book.clean_folder()

                #----------------------------------------------------------
                #----------------------------------------------------------
                #----------------------------------------------------------
                return render_template('home/webstats/ui-newWSReport.html',
                                    institution_name = institution_name,
                                    institution_init_year = institution_init_year,
                                    institution_end_year = institution_end_year,
                                    type = type,
                                    find_path = find_path,
                                    pic_path =pic_path,

                                    average_yr = average_yr,
                                    count_years = count_years,
                                    count_months = count_months,
                                    sum_session_yr  = sum_session_yr,
                                    sum_TIR_yr  = sum_TIR_yr,
                                    sum_TII_yr  = sum_TII_yr,
                                    sum_UIR_yr  = sum_UIR_yr,
                                    sum_UII_yr  = sum_UII_yr,
                                    sessions = sessions,
                                    TIR=TIR,
                                    TII=TII,
                                    UIR=UIR,
                                    UII=UII,

                                    sum_session_month  = sum_session_month,
                                    sum_TIR_month = sum_TIR_month,
                                    sum_TII_month  = sum_TII_month,
                                    sum_UIR_month  = sum_UIR_month,
                                    sum_UII_month  = sum_UII_month,
                                    sessions_m = sessions_m,
                                    TIR_m=TIR_m,
                                    TII_m=TII_m,
                                    UIR_m=UIR_m,
                                    UII_m=UII_m,
                                    Yearly_Comparison=Yearly_Comparison
                                    )
            else:
                book.clean_folder()
                flash(institutionIs + ' '+ ' still don´t have an image in the system, please upload one here')
                return render_template('home/webstats/ui-webStatsReport.html')

#--------------------------------------------------------------------
#-------------  DATABASES  -------------------------------------------
    elif KEY=='databases':
        #------------------------------------
        #		CREATE OBJECT
        #------------------------------------
        #---------------------------------
        # Creamos el objeto
        # Es necesario pasarle los parametros de inicializacion del constructor
        database = Database(TEMP_SAVE_PATH, FINAL_SAVE_PATH, SOURCE_FILE, type_graphic, use_grid, session_select,TII_select, TIR_select, UIR_select , Monthly_Trend, Yearly_Comparison, Monthly_Trend1, SR_select,image, UII_select)

        #---- NO UTILIZO IMAGEN ----
        if image == 'True':
            institutionIs = database.new_column()
            database.set_table()
            database.pivot_montly_trend()
            database.pivot_stats()
            extra_path, col_end = database.extra()
            database.set_table_extra(extra_path, col_end)
            if Yearly_Comparison == 'Yearly_Comparison':
                database.yearly_comp()
            else:
                pass

            average_yr, count_years,count_months,sum_session_yr ,sum_TIR_yr,sum_TII_yr,sum_UIR_yr,sum_UII_yr,sessions,TIR,TII,UIR,UII,sum_session_month ,sum_TIR_month,sum_TII_month,sum_UIR_month,sum_UII_month,sessions_m,TIR_m,TII_m,UIR_m,UII_m = database.webGraph()

            #------------------------------------------
            #-------- COUNTER -------------------------
            #------------------------------------------
            # Esta es la funcion counter que me muestra los gráficos al principio.
            database.counter(KEY)

            institution_name, institution_init_year, institution_end_year,find_path, type, pic_path  = database.final()
            time.sleep(15)
            flash('Report successfully generated','success')
            database.clean_folder()

            return render_template('home/webstats/ui-newWSReport.html',
                                institution_name = institution_name,
                                institution_init_year = institution_init_year,
                                institution_end_year = institution_end_year,
                                type = type,
                                find_path = find_path,
                                pic_path =pic_path,

                                average_yr = average_yr,
                                count_years = count_years,
                                count_months = count_months,
                                sum_session_yr  = sum_session_yr,
                                sum_TIR_yr  = sum_TIR_yr,
                                sum_TII_yr  = sum_TII_yr,
                                sum_UIR_yr  = sum_UIR_yr,
                                sum_UII_yr  = sum_UII_yr,
                                sessions = sessions,
                                TIR=TIR,
                                TII=TII,
                                UIR=UIR,
                                UII=UII,

                                sum_session_month  = sum_session_month,
                                sum_TIR_month = sum_TIR_month,
                                sum_TII_month  = sum_TII_month,
                                sum_UIR_month  = sum_UIR_month,
                                sum_UII_month  = sum_UII_month,
                                sessions_m = sessions_m,
                                TIR_m=TIR_m,
                                TII_m=TII_m,
                                UIR_m=UIR_m,
                                UII_m=UII_m,
                                Yearly_Comparison=Yearly_Comparison
                                )
        # ---- UTILIZO IMAGEN
        else:
            # Esto lo hacemos para poder escojer la imagen que necesitamos para el reporte
            # Resibimos del modelo el parametro que vamos a evaluar
            isExist = database.set_image()
            # Asi obtenemos el nombre del archivo
            #------- NEW COLUMN ----> Month-Year
            # Sacamos el nombre de la institucion para que podamos revisar si tiene foto o no
            institutionIs = database.new_column()

            if isExist == True:


                #------------------------------------
                #		CHECK FOR IMAGE TO CONTINUE
                #------------------------------------
                #------- Table format
                database.set_table()
                #------ Montly Trend
                # Para esto ya necesitamos utilizar absolute path
                database.pivot_montly_trend()
                #----- Stats General report
                database.pivot_stats()

        #-------------------------------------------------------------------------------------
                # Aqui saco el path y la columna final donde terminar
                extra_path, col_end = database.extra()

                database.set_table_extra(extra_path, col_end)
                #database.extra_chart(path)
                # Pillo todas la variables que necesito pasar a la vista

                ##-------- Yealry comparison --------------
                if Yearly_Comparison == 'Yearly_Comparison':
                    database.yearly_comp()
                else:
                    pass

                average_yr, count_years,count_months,sum_session_yr ,sum_TIR_yr,sum_TII_yr,sum_UIR_yr,sum_UII_yr,sessions,TIR,TII,UIR,UII,sum_session_month ,sum_TIR_month,sum_TII_month,sum_UIR_month,sum_UII_month,sessions_m,TIR_m,TII_m,UIR_m,UII_m = database.webGraph()

                #------------------------------------------
                #-------- COUNTER -------------------------
                #------------------------------------------
                # Esta es la funcion counter que me muestra los gráficos al principio.
                database.counter(KEY)

                institution_name, institution_init_year, institution_end_year,find_path, type, pic_path  = database.final()

                # Seudo loading para que aparente estar cargando
                time.sleep(15)
                flash('Report successfully generated','success')

                database.clean_folder()

                #----------------------------------------------------------
                #----------------------------------------------------------
                #----------------------------------------------------------
                return render_template('home/webstats/ui-newWSReport.html',
                                    institution_name = institution_name,
                                    institution_init_year = institution_init_year,
                                    institution_end_year = institution_end_year,
                                    type = type,
                                    find_path = find_path,
                                    pic_path =pic_path,

                                    average_yr = average_yr,
                                    count_years = count_years,
                                    count_months = count_months,
                                    sum_session_yr  = sum_session_yr,
                                    sum_TIR_yr  = sum_TIR_yr,
                                    sum_TII_yr  = sum_TII_yr,
                                    sum_UIR_yr  = sum_UIR_yr,
                                    sum_UII_yr  = sum_UII_yr,
                                    sessions = sessions,
                                    TIR=TIR,
                                    TII=TII,
                                    UIR=UIR,
                                    UII=UII,

                                    sum_session_month  = sum_session_month,
                                    sum_TIR_month = sum_TIR_month,
                                    sum_TII_month  = sum_TII_month,
                                    sum_UIR_month  = sum_UIR_month,
                                    sum_UII_month  = sum_UII_month,
                                    sessions_m = sessions_m,
                                    TIR_m=TIR_m,
                                    TII_m=TII_m,
                                    UIR_m=UIR_m,
                                    UII_m=UII_m,
                                    Yearly_Comparison=Yearly_Comparison
                                    )
            else:
                database.clean_folder()
                flash(institutionIs + ' '+ ' still don´t have an image in the system, please upload one here')
                return render_template('home/webstats/ui-webStatsReport.html')
#--------------------------------------------------------------------
#-------------  SI NO es ninguno de los tres(journal, databases, books)
    else:
        flash('An error has ocurred, please try again!')
        return render_template('home/webstats/ui-webStatsReport.html')

#------------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------
#----------------------------------------------             getOrders                   ---------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------#
#------------------------  INDEX  ---------------------------------------#
#-------------------------------------------------------------------------#
# Como ya importamos el objeto de BLUEPRINT desde el init_.py y alo puedo utilizar
# El Objeto la llamamos "blueprint" entonces es el mobre que utilizamos en el decorador
@blueprint.route('/orders', methods=('GET', 'POST'))
@login_required
def getOrders():
    # Como tengo tres tipos de reportes diferentes - journals,books,databases-  creé en la vista tres formularios
    # Necesito decirle a la ruta cual formulario deseo utilizar en cada caso.
    # Para ello en el valor VALUE del formulario especifique un identificador que utilizo aqui
    if request.method == 'POST':
# --------- JOURNALS ----------------------------------------------------------------------------------------------------
        # Necesito comparar el nombre del formulario que me llaga por la request con el VALUE de ese formulario
        if request.form.get('find') == 'find':
            # Ahora comprobamos que se esté cargardo un archivo válido
            # Con el nombre que le dimos al archivo en el VIEW lo trabajamos aq
            var = request.form.get('group')

            os.system("php test.php %s"%(var))

            return render_template('home/orders/getOrders.html')

    else:
        # Esta es la ruta a la que accedo cuando se solicita la pagina
        # Es decir con el GET aqui llego
        # Una vez adentro es que realizo el post a esta ruta ENTITLEMENT
        flash('Report successfully generated','success')
        return render_template('home/orders/getOrders.html')
