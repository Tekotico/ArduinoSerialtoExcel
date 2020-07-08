"# ArduinoSerialtoExcel" 

Instrucciones

1- Los datos enviados por el puerto serie deben estar separados por comas para ser interpretados como columnas distintas.
 
2- En la casilla "Puerto" debe escribirse el puerto al que esta conectado el Arduino. Debe escribirse de la forma " COM# " y puede estar en minúsculas. Ej: com1, COM4.

3- La velocidad debe coincidir con la que se configuró el puerto en Arduino.

4- Ingresar el nombre del archivo. No se requiere agregar la extensión. Si existe un archivo con el mismo nombre, éste se sobreescribirá.

5- La cantidad de muestras debe ser un número entero menor a 1048500 (está limitado por la cantidad máxima de filas que puede contener un archivo .xls)

6- Por defecto las columnas no tienen nombre. Para agregar una columna se debe ingresar el nombre en la casilla "Nueva columna" y presionar "Agregar".
	La columna agregada aparecerá en la tabla "Columnas". Para eliminar una columna basta con seleccionarla y presionar "Eliminar".

7- Una vez finalizada la configuración debe presionarse "Ejecutar". Cuando el archivo esté listo aparecerá el mensaje "Archivo * creado con éxito.".


Cosas a Mejorar:

1- La tabla de Columnas solo muestra 10 nombres de columnas. Si se agregan mas el programa las almacena y aparecerán en el archivo Excel, pero no se verán en la tabla.
    Para solucionarlo debería agregarse un scroll en la tabla.

2- Falta validación para entrada de puerto.

3- Añadir la posibilidad de agregar columnas presionando enter.

4- Emitir un mensaje que indique el porcentaje de datos guardados y/o que indique que se estan guardando (Está pero por algún motivo no anda).

