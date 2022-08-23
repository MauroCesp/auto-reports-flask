import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from datetime import datetime, date
from openpyxl import Workbook

# Esto libreria nos ayuda a crear pivot tables
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

##########################  IMPORT APP ##################################
# Es super necesario llamar al APP para poder utilizar las libreria propia 
# Como cunado necesitamos llamar al directorio raiz de la applicacion a la hora de importar ficheros.
from run import app 

######################
#Tablas e imagenes
# Primero importo algunas funciones para tablas e imagenes
from openpyxl.worksheet.table import Table, TableStyleInfo

from openpyxl.chart import PieChart, Reference, Series,PieChart3D,LineChart, BarChart
 

class Format:
    
#---------- CREATE NEW COLUMN -MONTH-YEAR- --------------   
    def new_column(filename):
        
        PATH = 'apps/uploads/tempData/'
        ##################
        #  OPEN FILE
        ##################
        # Ahora vamos a buscar el archivo y cargarlo
        wb = load_workbook(PATH + filename)
        ws = wb.active

        # Ahora creamos la cell que queremos agregar
        # Hay que tener en cuanta el tamaño del sheet para poder ubicarla bien en la columna correcta
        # Creo una variable donde guardo la cell

        newCol = ws['AE1']
        # test.font = Font(bold=True)
        newCol.value = 'Month-Year'

        #################
        #  FIND SIZE DF
        #################
        # Podemos averiguar el tamaño de la coumna facilment con pandas
        # Llamo a un dataframe para buscar la info rapido y asigno una variable
        df = pd.read_excel(PATH + filename, engine='openpyxl')
        rows = df.shape[0]
        cols = df.shape[1]
        
        # Esto es para poder saber en donde tengo que poner la ultima columna en caso de que cambie el numero de columnas
        size = df.shape[0]
        ##########################
        #  FOR LOOP NEW COLUMN
        ##########################
        # Ahora creamos un forloop para interactuar con todos los rows en estas columnas
        # RECORDAR que no debemos incluir el primer row porque son HEADING y no values
        # Ponemos de limite la variable 'df'
        # Le ponemos el size mas 2 porque me fglatan dos espacio
        # Le decimos que comience desde 2 pero el DF me cuenta solo los spacios con data y los los encabezados
        # Para compensar esto hacemos el truco
        for row in range(2,(size+2)):
            # En cada interaccion pillamos el numero de row para el MES y el año
            # Lo convertimos a STR para poder hacerlo una cadena de texto
            m = str(ws[f'AC{row}'].value)
            y = str(ws[f'AD{row}'].value)

            # Creamos la cadena de texto con la fecha complete
            a_date = "1"+"/"+m+"/"+y

            # En el último parametro especifico el formato de fecha que deseamos tener '%B %Y'
            # B% da el mes con nombre completo
            # m% nos da el numero del mes

            #############################
            #       NEW columna
            #############################
            # En la columna AE ya le hemos asignado nombre
            # Ahora necesitamos asignarle los valores
            ws[f'AE{row}'] = datetime.strptime(a_date, "%d/%m/%Y").strftime('%B %Y')

        #Salvamos el doc como excel
        wb.save(PATH + 'new_column.xlsx')
        
        df1 = pd.read_excel(PATH +'new_column.xlsx')

        # select  columns to display
        dframe = df1[['Title',
                'Unique_Item_Requests',
                'Total_Item_Requests',
                'Platform',
                'Subject',
                'OrderDescription',
                'OrderNumber',
                'UsedByCustomer',
                'Group',
                'User',
	            'Month',
                'Year',
                'Month-Year']]
        
        ####################
        #  FIND SIZE NEW DF
        ####################
        # Llamo a un dataframe para buscar la info rapido y asigno una variable
        newRows = dframe.shape[0]
        newCols = dframe.shape[1]
        
        # A esta tabla le quitamos el formato de tabla para incluir la ultima columna que creamos con la fecha
        # Le ponemos el index flase para que no aparezca la coumna index
        dframe.to_excel(PATH + 'table_test.xlsx', index = False) 
        
        return rows, cols, newRows, newCols
        
        
#---------- SET UP THE TABLE FORMAT- --------------        
    def set_table():
        
        PATH = 'apps/uploads/tempData/'
        
        #################
        #  FIND SIZE DF
        #################
        df = pd.read_excel(PATH + 'table_test.xlsx')
        
        # Aqui buscamos baer el tamaño de la table_test
        # Agregamos 1 para que cubra el ultimo row de abajo
        # Lo pasamos a string para poder pegarlo en una cadena
        size = str(1+(df.shape[0]))
        
        #################
        #  UPLOAD BOOK
        #################
        # Primero abro el libro que ya tiene la columna incluida
        wb = load_workbook(PATH + 'table_test.xlsx')
        # grab the active worksheet
        # This will create the active sheet on this work BOOK
        ws = wb.active
        
        # Usualmente le ponemos un punto como nombre a la tabla con todos los datos.
        ws.title = "."
        
        ######################
        #  CREATE TABLE RANGE
        ######################
        # Creamos un objeto de data que guarde los datos de la tabla
        # el parametro 'ref' indica el rango de los datos
        # Necesitamos cojer desde la columna A hasta la M--- porque el el numero de columnas que dejamos
        # A esto le agregamos el numero de rows --SIZE, de esta forma cubrimos todos los datos dentro de la tabla
        tab = Table(displayName='Table1',ref = 'A1:M'+size)
        
        ##################
        #  STYLE TO TABLE
        ##################
        # Ahora le damos estilo a la tabla
        # Creo una variable que guarde el objeto de estilo
        # Solo quiero que los rows tengan strpes. Las columnas las mantengo normales
        style = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False, showLastColumn= False,
                                                            showRowStripes=True, showColumnStripes=False)
        # ahora asigno la variable
        tab.tableStyleInfo = style

        # Por ultimo agregamo la tabla al worksheet
        ws.add_table(tab)

        wb.save(PATH + 'final_table.xlsx')

        
#---------- SET UP THE TABLE FORMAT- --------------  
    def set_pivot():
        
        PATH = 'apps/uploads/tempData/'
        #-------------------------------------
        # Primero leemos el dataframe_to_rows
        # -------------------------------------
        df = pd.read_excel(PATH + 'final_table.xlsx')
        
        ####################################
        #  UPLOAD BOOK AND CREATE NEW SHEET
        ####################################
        
        # Primero abro el libro que ya tiene la columna incluida
        wb = load_workbook (PATH + 'final_table.xlsx')
        
        # Creamos otro tab pra guardar el Pivot table
        ws_1 = wb.create_sheet("Journal Stats")

        #-------------------------------------
        # Trabajamos el dataframe con funcion nativa de Openpyxl
        # -------------------------------------
        #this fucntion is a native openpyxl function that allows us to work specifically with Pandas DataFrames
        #it allows us to iterate through the rows and append each one to our active worksheet.

        for r in dataframe_to_rows(df, index=True, header=True):
                ws_1.append(r)

        # Values--------> Es lo que va a rellenar los espacios
        # Index ---------> Es lo que va a servir como indice vertical
        # Columns -----------> Son las columnas que se despliegan

        data_piv = df.pivot_table(values=['Unique_Item_Requests','Total_Item_Requests'],index='Platform',columns='Year',aggfunc='sum')

        #create pivoyt table using Pandas

        for r in dataframe_to_rows(data_piv, index=True):
                ws_1.append(r)

        wb.save(PATH + 'report.xlsx')       
        
        
    def set_graphic():

        # Primero abro el libro que ya tiene la columna incluida
        wb = load_workbook('docs/temp/final_table.xlsx')
        ws = wb.active

        df = pd.read_excel('docs/temp/final_table.xlsx')

        size = df.shape[0]
        print(size)
        #-------------------------------------
        # Creamos otro tab pra guardar el grafico
        # With teh work book object we call the create fucntion
        ws_1 = wb.create_sheet("Statistics")

        #-------------------------------------
        # CREAMOS grafico
        #-------------------------------------

        # Ahora que tenemos informacion tenemos que decidir como se presenta en la grafica
        chart = BarChart()

        # Especificamos desde donde, hasta donde queremos la data
        # Tenemos que encontrar el tamaño de la tabla
        labels = Reference(ws,min_col=2, min_row=1, max_row=size)

        data = Reference(ws,min_col=2, max_col=ws.max_column, min_row=1, max_row=ws.max_row)


        chart.add_data(data,titles_from_data=True)
        chart.set_categories(labels)
        chart.title = 'Ice crema Flavor'

        # Añadimos el gráfico a la pestaña correcta
        # Decimos en que celda deseo que comience.
        ws_1.add_chart(chart,'Q5')

        # Cunado tenemos todo lo que queremos guardamos el archivo
        wb.save('docs/temp/graphs.xlsx')

        return wb




