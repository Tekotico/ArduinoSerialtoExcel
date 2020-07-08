from tkinter import ttk
from tkinter import *
from serialToExcel import SerialToExcel

columnas = []
class Principal:

    
    def __init__(self, window):
        self.wind = window
        self.wind.title('Arduino Serial to Excel')

        

        #creando un Frame Container
        frame = LabelFrame(self.wind, text = 'Configuracion')
        frame.grid(row = 0, column = 0, columnspan = 5, pady= 20) 

        #Entrada Puerto
        Label(frame, text = 'Puerto: ').grid(row = 1, column= 0 )
        self.puerto = Entry(frame)
        self.puerto.focus() #Focus es para que el cursor aparezca en la casilla
        self.puerto.grid(row = 1, column = 1)

        #Entrada Velocidad, Combobox para seleccion de velocidad
        Label(frame, text = ' Velocidad (Bd): ').grid(row = 2, column = 0)
        self.combo = ttk.Combobox(frame, state="readonly")
        self.combo.grid(row = 2, column = 1)
        self.combo["values"] = ["300", "600", "1200", "2400", "4800", "9600", "14400", "19200", "28800", "38400", "57600", "115200"]
        
        #Boton Ejecutar
        ttk.Button(frame, text = 'Ejecutar', command = self.Ejecutar).grid(row = 18, columnspan = 2, sticky = W + E)

        #Mensajes de salida
        self.message = Label(text = '', fg = 'red')
        self.message.grid(row = 6, column = 0, columnspan = 2, sticky = W + E)

        #Tabla vista previa
        self.tree = ttk.Treeview(height = 10)
        self.tree.grid(row = 0, column = 6, columnspan = 5)
        self.tree.heading('#0', text = 'Columnas', anchor = CENTER)
        

        #Agregar o quitar columnas
        Label(frame, text = 'Agregar o quitar columnas.').grid(row = 14, columnspan = 5, sticky = W + E)

        #Nombre de archivo
        Label(frame, text = 'Nombre del Archivo Excel: ').grid(row = 4, column = 0 )
        Label(frame, text = '.xls').grid(row = 4, column = 2, sticky = W )
        self.archivo = Entry(frame)
        self.archivo.grid(row = 4, column = 1)


        #Numero de datos
        Label(frame, text = 'Cantidad de muestras: ').grid(row = 5, column = 0)
        self.muestras = Entry(frame)
        self.muestras.grid(row = 5, column = 1)

        #Nueva Columna
        Label(frame, text = 'Nueva Columna:').grid(row = 15, column = 0)
        self.new_column = Entry(frame)
        self.new_column.grid(row = 15,column = 1)

         
        #Boton Agregar columna
        ttk.Button(frame, text = 'Agregar', command = self.add_column).grid(row = 16, columnspan = 1, sticky = W + E)

        #Boton Eliminar columna
        ttk.Button(frame, text = 'Eliminar', command = self.delete_column).grid(row = 17, columnspan = 1, sticky = W + E)


        #Datos cargados
        self.messagedatos = Label(text = '', fg = 'blue')
        self.messagedatos.grid(row = 16, column = 4, sticky = W + E)

        

       
        
        



    def Ejecutar(self):
        self.message['text'] = ''
        a = 0
        if  len(self.muestras.get()) == 0: a = 4
        else :
            if int (self.muestras.get()) > 1048500 : a = 5
        if  len(self.archivo.get()) == 0: a = 3
        if  len(self.combo.get()) == 0 : a = 2
        if  len(self.puerto.get()) == 0: a = 1

        if self.switch(a):
            self.message['fg'] = 'red'
            self.message['text'] = self.switch(a)
        else:
            self.message['fg'] = 'red'
            self.message['text'] = 'guardando'   #No funciona (????
            serialToExcel = SerialToExcel(self.puerto.get().upper(),int (self.combo.get()))
            global columnas   
            serialToExcel.setColumns(columnas)
            serialToExcel.setRecordsNumber(int (self.muestras.get()))
            serialToExcel.readPort()

            if serialToExcel.writeFile(self.archivo.get() + ".xls") :
             self.message['fg'] = 'green'
             self.message['text'] = 'Archivo ' + self.archivo.get() + ' creado con éxito.'

    
    
    def switch(self, a):
            switcher = {
                1: 'Ingrese puerto (Ej: COM1).',
                2: 'Seleccione Velocidad (EJ: 9600).',
                3: 'Ingrese un nombre de Archivo.',
                4: 'Ingrese un número de muestras.',
                5: 'Ingrese un número entero menor a 1048500.',
            } 
            return switcher.get(a, False)



    def add_column(self):
        if len(self.new_column.get()) != 0:
            self.message['text'] = ''
            new_columns = [self.new_column.get()]
            global columnas
            columnas.extend(new_columns)
            self.tree.insert('', 'end', text = self.new_column.get())
            self.message['fg'] = 'green'
            self.message['text'] = 'Columna Agregada.'
        else:
            self.message['fg'] = 'red'
            self.message['text'] = 'Introduzca un Nombre para Columna.'
            return
        self.new_column.delete(0, 'end')
   

    def delete_column(self):
        global columnas
        self.message['text'] = ''
        try:
            self.tree.item(self.tree.selection())['text'][0]
        except IndexError as e:
            self.message['fg'] = 'red'
            self.message['text'] ='Seleccione Columna a Eliminar.'
            return
        self.message['text'] = ''
        del_col = self.tree.item(self.tree.selection())['text']
        columnas.remove(del_col)
        self.tree.delete(self.tree.selection())
        self.message['fg'] = 'green'
        self.message['text'] = 'Columna eliminada.'

    

if __name__ == '__main__':
    window = Tk()
    application = Principal(window)
    window.mainloop()  