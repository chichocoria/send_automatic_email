# send_automatic_email
Script en Phyton que envia mails desde un listado de excel y adjunta archivos.

Autor: Ruben Dario Coria

## Descripcion General
El script esta diseñado para enviar mails masivos a una lista de personas con su numero de legajo
La lista se encuentra en una archivo excel.
A su vez cuenta con un archivo txt que se puede editar para añadir el cuerpo del mensaje.
Para que funcione el script, el nombre del archivo debe contener el formato "leg_XX.pdf" donde XX es el numero de legajo de cada empleado. (no es key sensitive)

## Como funciona?
Al ejecutar el script lee desde el archivo excel si coincide el nombre del archivo PDF con la columna correspondiente, si es asi, se envia un mail con el PDF adjunto. Se puede enviar cualquier archivo adjunto, el ejemplo es con un PDF que coincida con la columna de legajos del archivo excel.

## Limitaciones
-Al acceder una nueva persona se debe agregar en el archivo readfile.xlsx
-Es de suma importancia que el archivo PDF contenga al final del nombre el formato "LEG_XX.pdf"  donde XX es el numero de legajo de cada persona.

## Releases

### V1.01
-Se cambio el formato de la tabla de excel para unificar los valores y hacer mas facil la busqueda.


### V1.0
- Se agrega un log con level INFO en donde se puede visualizar fecha y hora del inicio del sistema, muestra una lista de mails enviados.

## Compilacion
Para compilar el archivo se utilizo.
 * Windows 10 64bits
 * Python 3.10.2
 * pyinstaller 4.8

### Correr en consola

pip install pyinstaller

### Compilar el archivo .py

pyinstaller --onefile --windowed nombre_archivo.py
