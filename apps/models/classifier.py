import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
import datetime
import os

class Classifier():
    	#----------------------------------------------------#
        #					SET-UP
        #----------------------------------------------------#
        def setUp(f):
            #The above function creates a dictionary with sheet names in the Excel files as
            #keys and dataframe as values. You can now access the dataframe with its sheet name.
            # Leo el archivo queme llega por request

            df = pd.read_excel(f,sheet_name=None, engine='openpyxl')

            # Aqui obtenemos el numero de tabs que vienen en la hoja
            # Hacemos una lista para guardar el numero de get_sheet_names
            holder = []
            #--------------------------------------------------

            # Recorrermos el array por keys que son las tabs# Esto porque el archivo se guarda com dicccionario
            for i in df.keys():
                holder.append(i)

            return holder, df
        
    	#----------------------------------------------------#
        #					WRITE FILES
        #----------------------------------------------------#
        def write_file(holder, df):

            # Primero especificamos la ruta donde vamos a guardar los archivos 
            out_path = r"/workspace/flaskAtlnatisDashboa/apps/templates/tempData/webStatsReport.xlsx"
            
            # Aqui unimos cada taba que viene en el reporte
            #Guardamos todo en un solo reporte para poder manipularlo.
            with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:

                for i in holder:
                    #Call os.path.join(components) with components as the constituent parts of the
                    #file path to join them together to create a full file path.
                    a_path =r"apps/templates/tempData"
                    # Asigno una variable para que me averigue la ruta
                    a_file = i+".html"

                    # La variable me guarda el path exacto a cada tab de la hoja de excel.
                    joined_path = os.path.join(a_path, a_file)

                    # Hago las variables dinamicas
                    df[i].to_excel(writer, sheet_name= i , index=False)

                    #df[i].to_html(i+".html")
                    df[i].to_html(joined_path)

                return joined_path
            

