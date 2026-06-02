#Autro : Ian Antonio Pérez Liska
#Lenguaje: Python
#versión de python: 3.12
#Procesos falntantes: ninguno
#Historia:
#         inicio: 27/5/2026
#         proceso: se instalaron librerías para la reducción de código y manipulación de archivos
#                  Interfaz gráfica para la selección de archivos y ruta de guardado
#                  Funciones para convertir números a letras y para convertir placas a letras
#                  Funciones para reescribir la información de excel a word para persona individual y persona jurídica
#                  Arreglos en errores menores y pruebas de funcionamiento                                             
#         fin: 01/6/2026

#librerías
from tkinter import *; from tkinter import filedialog as f, messagebox as ms
from docxtpl import DocxTemplate as DT; import pandas as pd
from num2words import num2words

class app(Tk):
    def __init__(self):
        super().__init__()

        #configuración de la ventana
        self.geometry('380x150')
        self.title('Paso de información de excel a word')
        self.config(bg='black')
        
        #opciones
        Label(text='↓Elige una opción↓',bg='black',fg='white',font='arial 12').pack()
        Button(self,text='Reescribir\n para\n Persona individual',command=self.reescribirIndi,width=20,height=5,bg='lightblue',font='arial 9').pack(side='right')
        Button(self,text='Reescribir\n para\n Persona jurídica',command=self.reescribirentity,width=20,height=5,bg='cyan',font='arial 9').pack(side='left')
    
#Funciones

    #declaración para persona individual
    def reescribirIndi(self):
        ms.showinfo('Aviso','Selecciona los archivos a usar para la reescritura')
        #deja que el usuario ingrese los documentos a usar de word y excel
        excel = f.askopenfilename(filetypes=[('Excel files', '*.xlsm *.xlsx')])
        doc = f.askopenfilename(filetypes=[('Word files', '*.docx')])

        if excel == '' or doc == '': # si no se escogio ninguno documento y sea word o excel
            ms.showerror('Error','No se han seleccionado los archivos')
        else:
            ms.showinfo('Información','Elige la ruta donde se guardarán los documentos') 
            ruta = f.askdirectory()# hace que el usuario ingrese su ruta de guardado de todos los documentos
            excel_file = pd.read_excel(excel, sheet_name='Persona_individual') # abre el excel y lee solo la hoja que se llame 'Persona_individual'
            for i, fila in excel_file.iterrows(): # recorre todo el documento

                #Pasa la información fila por fila del excel a un word con la misma estructura que la plantilla de word
                reescrituras = {'day_of_month': fila["day_of_month"],
                                'month_name':fila["month_name"],
                                'hour_of_day':self.hora(str(fila["hour_of_day"])) +'('+ str(fila['hour_of_day'])+')',
                                'lawyer_name':fila["lawyer_name"],
                                'lawyer_code':fila["lawyer_code"],
                                'notary':fila["notary"],
                                'constituted':fila["constituted"],
                                'type_of_person':fila["type_of_person"],
                                'owner_name':fila["owner_name"],
                                'id_in_words':self.nit(str(fila["id_in_numbers"])[:4])+', '+self.nit(str(fila["id_in_numbers"])[4:-4])+', '+self.nit(str(fila["id_in_numbers"])[9:]),
                                'id_in_numbers':str(fila["id_in_numbers"])[:4] +' '+str(fila["id_in_numbers"])[4:-4]+' '+str(fila["id_in_numbers"])[9:],
                                'tc_in_words':self.nit(str(fila["tc_in_numbers"])),
                                'tc_in_numbers':fila["tc_in_numbers"],
                                'vehicle_color':fila["vehicle_color"],
                                'nit_in_words':self.nit(str(fila["nit_in_numbers"])),
                                'nit_in_numbers':fila["nit_in_numbers"],
                                'vehicle_owner':fila["vehicle_owner"],
                                'use_type':fila["use_type"],
                                'vehicle_type':fila["vehicle_type"],
                                'license_plate_in_words':self.placa(str(fila["license_plate_in_numbers"])),
                                'license_plate_in_numbers':fila["license_plate_in_numbers"],
                                'vehicle_make':fila["vehicle_make"],
                                'line_in_words':self.convertir(str(fila["line_in_numbers"])),
                                'line_in_numbers':fila["line_in_numbers"],
                                'year_in_words':num2words(int(fila['year_in_numbers']), to='year', lang='es'),
                                'year_in_numbers':fila["year_in_numbers"],
                                'vin_in_words':self.convertir(str(fila["vin_in_numbers"])),
                                'vin_in_numbers':fila["vin_in_numbers"],
                                'weight_in_words':num2words(int(fila["weight_in_numbers"]),lang='es'),
                                'weight_in_numbers':fila["weight_in_numbers"],
                                'slv_model':fila["slv_model"],
                                'provial_code_in_words':self.convertir(fila["provial_code"]),
                                'provial_code':fila["provial_code"],
                                }

                doc_file = DT(doc)
                doc_file.render(reescrituras)#hace todas las reescrituras
                doc_file.save(ruta + f"/declaration_{fila["license_plate_in_numbers"]} rev.docx")# guarda el documento con un nombre específico

    #declaración para persona jurídica
    def reescribirentity(self):
        ms.showinfo('Aviso','Selecciona los archivos a usar para la reescritura')
        #deja que el usuario ingrese los documentos a usar
        excel = f.askopenfilename(filetypes=[('Excel files', '*.xlsm *.xlsx')])
        doc = f.askopenfilename(filetypes=[('Word files', '*.docx')])

        if excel == '' or doc == '': # verifica que se hayan ingresado los dos documentos
            ms.showerror('Error','No se han seleccionado los archivos')
        else:
            ms.showinfo('Información','Elige la ruta donde se guardarán los documentos')
            ruta = f.askdirectory()# hace que el usuario ingrese la ruta de guardado de sus documentos
            excel_file = pd.read_excel(excel, sheet_name='Persona_juridica')# abre el excel y lee solo la hoja de nombre 'Persona_jurídica'
            for i, fila in excel_file.iterrows():# recorre fila y columna del excel

                #Pasa la información fila por fila del excel a un word con la misma estructura que la plantilla de word
                reescrituras = {'day_of_month': fila["day_of_month"],
                                'month_name':fila["month_name"],
                                'hour_of_day':self.hora(str(fila["hour_of_day"])) +'('+ str(fila['hour_of_day'])+')',
                                'lawyer_name':fila["lawyer_name"],
                                'lawyer_code':fila["lawyer_code"],
                                'notary':fila["notary"],
                                'constituted':fila["constituted"],
                                'type_of_person':fila["type_of_person"],
                                'owner_name':fila["owner_name"],
                                'id_in_words':self.nit(str(fila["id_in_numbers"])[:4])+', '+self.nit(str(fila["id_in_numbers"])[4:-4])+', '+self.nit(str(fila["id_in_numbers"])[9:]),
                                'id_in_numbers':str(fila["id_in_numbers"])[:4] +' '+str(fila["id_in_numbers"])[4:-4]+' '+str(fila["id_in_numbers"])[9:],
                                'tc_in_words':self.nit(str(fila["tc_in_numbers"])),
                                'tc_in_numbers':fila["tc_in_numbers"],
                                'vehicle_color':fila["vehicle_color"],
                                'nit_in_words':self.nit(str(fila["nit_in_numbers"])),
                                'nit_in_numbers':fila["nit_in_numbers"],
                                'vehicle_owner':fila["vehicle_owner"],
                                'use_type':fila["use_type"],
                                'vehicle_type':fila["vehicle_type"],
                                'license_plate_in_words':self.placa(str(fila["license_plate_in_numbers"])),
                                'license_plate_in_numbers':fila["license_plate_in_numbers"],
                                'vehicle_make':fila["vehicle_make"],
                                'line_in_words':self.convertir(str(fila["line_in_numbers"])),
                                'line_in_numbers':fila["line_in_numbers"],
                                'year_in_words':num2words(int(fila['year_in_numbers']), lang='es'),
                                'year_in_numbers':fila["year_in_numbers"],
                                'vin_in_words':self.convertir(str(fila["vin_in_numbers"])),
                                'vin_in_numbers':fila["vin_in_numbers"],
                                'weight_in_words':num2words(int(fila["weight_in_numbers"]), lang='es'),
                                'weight_in_numbers':fila["weight_in_numbers"],
                                'slv_model':fila["slv_model"],
                                'provial_code_in_words':self.convertir(fila["provial_code"]),
                                'provial_code':fila["provial_code"],
                                'company_name':fila["company_name"],
                                'register_in_words':self.convertir(str(fila["register_in_numbers"])),
                                'register_in_numbers':fila["register_in_numbers"],
                                'folio_in_words':num2words(int(fila["folio_in_numbers"]), lang='es') if str(fila["folio_in_numbers"]).isdigit() else self.convertir(str(fila["folio_in_numbers"])),
                                'folio_in_numbers':fila["folio_in_numbers"],
                                'book_in_words':num2words(int(fila["book_in_numbers"]), lang='es'),
                                'book_in_numbers':fila["book_in_numbers"],
                                }

                doc_file = DT(doc)
                doc_file.render(reescrituras)#hace todas las reescrituras
                doc_file.save(ruta + f"/declaration_{fila["license_plate_in_numbers"]} rev.docx")# guarda el documento con un nombre específico

    #funcion para pasar las horas a letras
    def hora(self,h:str):
        partes = h.split(':') #quita el dos puntos del string
        uno = partes[0] # toma los numeros de las horas
        dos = partes[1] # toma los numeros de los minutos

        #convertimos a letras
        horas = num2words(int(uno), lang='es')
        minutos = num2words(int(dos), lang='es')

        return horas + ' horas' + ' con ' + minutos + ' minutos' #devuelve las letras de los números convertidos

    def placa(self, placa: str):
        numeros = ['cero','uno', 'dos', 'tres', 'cuatro', 'cinco', 'seis', 'siete', 'ocho', 'nueve']
        
        if not placa: # si ya no hay nada que convertir se detiene
            return ''
        
        dih = placa[:1] # primer caracter
        resto = placa[1:] # resto de caracteres

        if dih in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ': #si el caracter son letras lo deja como esta y sigue con los demás caracteres
            return dih + ' ' + self.placa(resto)
        if dih in '0123456789':# si el caracter son números
            return numeros[int(dih)] + ' ' + self.placa(resto)
        return self.placa(resto)  # ignorar caracteres no reconocidos

    #funcion para pasar el nit a letras
    def nit(self, pnit: str):
        if not pnit:# si ya no hay nada que convertir se detiene
            return ''
        
        dih = pnit[:1] # primer caracter
        resto = pnit[1:] # resto de caracteres
        j = '' #string de acumulación

        if dih == '0': #si hay un cero antes de todo escribe cero
            return 'cero' + ' ' + str(self.nit(resto))
        
        if dih in '1234567890':
            while pnit and pnit[0] in '1234567890': #vemos si hay más numeros despues de ese caracter
                j += pnit[0] #acumulamos los números en el string de acumulación
                pnit = pnit[1:] #quitamos el número ya acumulado
            return num2words(int(j), lang='es') + ' ' + str(self.nit(pnit))
        
        if dih == '-': # si hay un guió escribir la palabra guion y seguir con los siguientes caracteres
            return 'guion' + ' ' + str(self.nit(resto))
        

    #función para pasar numeros y placas a letras
    def convertir(self, c: str):
        if not c:  #cuando ya no quedan caracteres, detener la recursión
            return ''

        car = c[:1] #toma el primer caracter
        demas = c[1:] #toma el resto de caracteres
        s = '' #string de acumulación

        if car == '0': #si hay un cero antes de todo escribe cero
            return 'cero' + ' ' + self.convertir(demas)
        
        if car == 'y': #si hay dos números separados por una y en medio escribir la 'y'
            return 'y' + ' ' + self.convertir(demas)
        
        if car in 'ABCDEFGHIJHIJKLMNOPQRSTUVWXYZ': #si el caracter son letras lo deja como esta y sigue con los demás caracteres
            return car + ' ' + self.convertir(demas)
        if car in '0123456789':# si el caracter son números
            while c and c[0] in '0123456789': #vemos si hay más numeros despues de ese caracter
                s += c[0] #acumulamos los números en el string de acumulación
                c = c[1:] #quitamos el número ya acumulado
            return num2words(int(s), lang='es') + ' ' + self.convertir(c)
        
        if car == '-':# si hay un guió escribir la palabra guion y seguir con los siguientes caracteres
            return 'guion ' + self.convertir(demas)
        if car == '/': #si hay una diagona escribir la palabra diagonal y seguir con los siguientes caracteres
            return 'diagonal ' + self.convertir(demas)
        return self.convertir(demas)  # ignorar caracteres no reconocidos

app().mainloop()