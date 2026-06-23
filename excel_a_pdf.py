#Objetivo: pasar información deun excel a un documento PDF
#Autor : Ian Antonio Pérez Liska
#Lenguaje: Python
#versión de python: 3.12
#Procesos falntantes: ninguno
#Historia:
#         inicio: 19/6/2026
#         fin: 23/6/2026

#Entrada: documento excel a pasar y PDF a usar de plantilla
#Proceso: 
#       1. Leer cada fila del excel 
#       2. Ver cuantas columnas tiene
#       3. tomar la información de cada fila
#       4. meter esa información en el PDF
#       5. repetir el proceso hasta que ya no hallan más colunas por leer
#Salida: guardar todos los PDF con la información ingresada


#Librerías
from tkinter import * # interface gráfica
from tkinter import filedialog as f, messagebox as ms # para diálogos de selección de archivos y mensajes
import pandas as pd # para manipulación de datos en Excel
from PyPDF2 import PdfReader, PdfWriter # para manipulación de archivos PDF

class App(Tk):
    def __init__(self):
        super().__init__()

        #configuración de la pantalla
        self.geometry('380x150')
        self.title('De Excel a Pdf')
        self.config(bg='black')

        #Objetos de la interfaz
        Label(self, text='↓ Elige una opción ↓', bg='black', fg='white', font='arial 12').pack()
        Button(self, text='Generar', command=self.generar,width=20, height=5, bg='lightblue', font='arial 9').pack() # boton para la función de generar
    

    #Funcion para hacer el paso de excel a pdf
    def generar(self):
        #pedidas de archivos y ruta de guardado
        ms.showinfo('Aviso', 'Selecciona el archivo Excel')
        excel = f.askopenfilename(filetypes=[('Excel', '*.xlsm *.xlsx')])

        ms.showinfo('Aviso', 'Selecciona la plantilla PDF')
        pdf = f.askopenfilename(filetypes=[('PDF', '*.pdf')])

        ms.showinfo('Aviso', 'Elige la carpeta donde se guardarán los PDFs')
        ruta = f.askdirectory()

        #Lectura de hoja específica del excel
        df = pd.read_excel(excel, sheet_name='FORMATO PARA NUEVOS CODIGOS')

        # Solo filas con datos (columna 1 tiene el número de fila)
        filas_con_datos = []
        for _, fila in df.iterrows():
            valor_col1 = str(fila.iloc[1]).strip()
            if valor_col1.isdigit():
                filas_con_datos.append(fila)
        datos = pd.DataFrame(filas_con_datos)

        for indi in datos.index: # ciclo para recorrer cada columna
            for fila in [datos.loc[indi]]: # ciclo para recorrer cada fila

                #función para obtener cada valor de cada celda de cada fila de cierta columna
                def var(col):
                    val = fila.iloc[col]
                    if pd.isna(val): return ''
                    # if hasattr(val, 'strftime'): return val.strftime('%d/%m/%Y')
                    return str(val).strip()

                reescrituras = {
                    # página 1 de la declaración jurada
                    'No de Certificado':var(2),
                    'Texto4':'',
                    'Texto3': var(25)[:2],
                    'Texto1': var(25)[3:-5],
                    'Texto2': var(25)[6:],
                    'Nombre de la Sociedad Mercantil Empresa Mercantil o Propietario':var(4),
                    'Dirección fiscal':var(22),
                    'Correo electrónico':var(20),
                    'No teléfonos':var(24),
                    'Número de Identificación Tributaria NIT':var(19),
                    'No tarjeta de circulación':var(47),
                    'No código correlativo':var(85),
                    'Texto5':var(53),
                    'Código único de identificación CUI':var(52),
                    'Uso':var(54),
                    'Placa':var(55),
                    'Tipo':var(49),
                    'Texto6':var(56),
                    'Línea':var(57),
                    'Modelo':var(58),
                    'Chasis':var(59),
                    'Texto7':var(60),
                    'Serie':var(61),
                    'Motor':var(62),
                    'CC':var(66),
                    'Cilindros':var(65),
                    'Asientos':var(63),
                    'Ejes':var(64),
                    'Color':var(50),
                    'Toneladas':var(67),
                    'Tipo de sistema':var(87),
                    'Número de Registro como implementador proporcionado por Provial':var(2),
                    'No Folio':f'Reg:{var(6)} Folio:{var(7)} Libro:{var(8)}',
                    'No Regisro':var(11),
                    'Texto10':var(12),
                    'Texto11':var(13),
                    'Nombre completo del propietario yo  representante legal':var(9),
                    'Edad':var(15),
                    'Estado civil':var(16),
                    'Nacionalidad':var(17),
                    'Profesión u oficio':var(18),
                    'No Regisro_2':var(6),
                    'Texto8':var(7),
                    'Texto9':var(8),
                    'Dirección':var(42),
                    'No teléfonos_2':var(43),

                    # página 2 de la declaración jurada
                    'calidad': var(25)[:2], # día
                    'de': self.meses(var(25)[3:-5]), # mes
                    'siendo las': var(25)[6:], # año
                    #falta las horas 
                    'entidad': var(9), # nombre del requerido
                    'quien requiere de': var(22), #ubicacion
                    'estando ubicados en': var(18), # en su calidad
                    'correctos y exactos de no ser asi expresamente renuncia al fuero de su domicilio y se somete a las acciones judiciales': var(2), # certificado de implementación del sistema limitador de velocidad, número...
                    'mis servicios profesionales para prestar Declaración Jurada sobre hechos de su interés para lo cual se le hace saber las penas': var(4), # entidad
                    'requirente quien enterado de su contenido objeto válidez y demás efectos legales la acepta ratifica y firma': '10' # minutos despues

                }

                reader = PdfReader(pdf) # toma el PDF seleccionado como plantilla
                writer = PdfWriter()# crea un nuevo PDF para escribir la información
                writer.append(reader)# agrega las páginas del PDF de plantilla al nuevo PDF
                writer.update_page_form_field_values(writer.pages[0], reescrituras) # reescribe las cosas en la primera página del pdf
                writer.update_page_form_field_values(writer.pages[1], reescrituras) # reescribe las cosas en la segunda página del pdf

                #guarda el pdf con todo ya cambiado
                with open(f'{ruta}/Certificado_{var(1)}_{var(55)}.pdf', 'wb') as out:
                    writer.write(out)

        ms.showinfo('Listo', f'Se generaron {len(datos)} certificados en:\n{ruta}') # aviso de que se guardaron ya los archivos

    def meses(self, n:str):
        meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']

        if n == '01':
            return meses[0]
        if n == '02':
            return meses[1]
        if n == '03':
            return meses[2]
        if n == '04':
            return meses[3]
        if n == '05':
            return meses[4]
        if n == '06':
            return meses[5]
        if n == '07':
            return meses[6]
        if n == '08':
            return meses[7]
        if n == '09':
            return meses[8]
        if n == '10':
            return meses[9]
        if n == '11':
            return meses[10]
        if n == '12':
            return meses[11]

App().mainloop()
