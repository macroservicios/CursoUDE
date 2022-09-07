# -*- coding: utf-8 -*-
"""
Created on Wed Aug 24 18:21:12 2022

PROYECTO 1
Nombre del proyecto: Tkinter y Excel

Objetivo
Conocer el funcionamiento básico de la herramienta GUI Tkinter y el uso de funciones
para gestión de archivos Excel.

Descripción
La idea es generar un formulario con 4 o 5 campos usando la herramienta Tkinter y luego
almacenarlos en una planilla Excel.
El funcionamiento se puede asimilar al de un formulario de Google. Se puede considerar
realizar validación de campos en la medida que el usuario va ingresando la información
y al final se confirmará el formulario y se grabará una planilla Excel con los datos
ingresados.

@author: silvah
"""

from tkinter import *         #Importamos tkinter para diseño de la interface grafica
import pandas as pd           #importamos la biblioteca pandas para manipulación de archivos


formulario = Tk()                                     #Se crea el objeto formulario

formulario.title("Agenda de contactos")              #Se le asigna un titulo a la ventana formulario
formulario.configure(bg="#48D1CC")                   #Se le asigna un color a la ventana formulario
formulario.iconbitmap('C:/Users/SilvaH/OneDrive - Anheuser-Busch InBev/Documents/Datos/agenda.ico')

#Se obtienen el ancho y alto de la venbtana
x_ventana = formulario.winfo_screenwidth() // 2 - 500 // 2   
y_ventana = formulario.winfo_screenheight() // 2 - 450 // 2

#Se define la variable posición para centran la ventana en el medio de la pantalla
posicion = str(500) + "x" + str(450) + "+" + str(x_ventana) + "+" + str(y_ventana)

formulario.geometry(posicion)                        #Se le dan las medidas de ancho y alto a la ventana



#Cargamos los datos de la archivo excel y los guardamos en df1
df1 = pd.read_excel('C:/Users/SilvaH/OneDrive - Anheuser-Busch InBev/Documents/Datos/Agenda.xlsx')
df1 = df1[['Nombre', 'Apellido', 'Teléfono', 'Email', 'Dirección']]  #Tomamos de df1 las columnas que nos interesan


df2 = pd.DataFrame()  #Creamos un DataFrame vacio donde luego se guardarán los datos del formulario




#Declaramos la función guardar
def guardar():
          
        #Guardamos los datos que se digitan en las entradas en df2
        df2['Nombre'] = [text_nombre.get()]
        df2['Apellido'] = [text_apellido.get()]
        df2['Teléfono'] = [text_telefono.get()]
        df2['Email'] = [text_email.get()]
        df2['Dirección'] = [text_direccion.get()]
             
        #Borramos las entradas luego de ser guardadas
        text_nombre.delete(0,END)
        text_apellido.delete(0,END)
        text_telefono.delete(0,END)
        text_email.delete(0,END)
        text_direccion.delete(0,END)
         
        #Concatenamos df1 con df2 y creamos df3 
        df3 = pd.concat([df1,df2], ignore_index=True)
        #guardamos df3 en el archivo excel
        df3.to_excel('C:/Users/SilvaH/OneDrive - Anheuser-Busch InBev/Documents/Datos/Agenda.xlsx')

        #Se confirman los datos guardados
        messagebox.showinfo('Mensaje','Los datos fueron grabados con exito')
  

#Creamos las etiquetas y las entradas y le damos una ubicación en el formulario
Label(formulario, text="Nombre: ", width=15, bg="#48D1CC").grid(column=0, row=0, pady=20, padx=10)
text_nombre = Entry(formulario, width=30)
text_nombre.grid(column=1, row=0)

Label(formulario, text="Apellido: ", width=15, bg="#48D1CC").grid(column=0, row=1, pady=20, padx=10)
text_apellido = Entry(formulario, width=30)
text_apellido.grid(column=1, row=1)

Label(formulario, text="Teléfono: ", width=15, bg="#48D1CC").grid(column=0, row=2, pady=20, padx=10)
text_telefono = Entry(formulario, width=30)
text_telefono.grid(column=1, row=2)

Label(formulario, text="email: ", width=15, bg="#48D1CC").grid(column=0, row=3, pady=20, padx=10)
text_email = Entry(formulario, width=30)
text_email.grid(column=1, row=3)

Label(formulario, text="Direción: ", width=15, bg="#48D1CC").grid(column=0, row=4, pady=20, padx=10)
text_direccion = Entry(formulario, width=30)
text_direccion.grid(column=1, row=4)


#creamos el botón guardar, lo ubicamos en el formulario y le asignamos el comando 
Button(formulario, text="Grabar datos", width=15, command=guardar).grid(column=1, row=5, pady=30)


formulario.mainloop()


