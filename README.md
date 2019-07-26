# WebScraping
Proyecto de validación de información

El proceso de validación de contratos se divide en dos partes:

1. La primera parte exporta a un archivo de excel una tabla con el número de todos los contratos y su respectiva url, tanto para test como para bios.
Después exporta a un archivo de excel los contratos en donde el número de contrato esta tanto en bios como en test, y sus respectivas url's.

2. La segunda parte entraría a comparar la información que tiene cada contrato en bios y en test. Exporta a un archivo de excel el resultado,
en donde hay una columna para cada uno de los campos, y se encuentra una diferencia, escribe en la celda el texto que tiene en bios y en test.
En otro caso escribe "ok".

Para que el código funcione en cada computado, toca cambiar algunas cosas:
 1. Cambiar el directorio
 2. Tener el archivo configTest.py que contiene las credenciales y tenerlo en la carpeta directorio
 3. Descargar el chromedriver de https://chromedriver.storage.googleapis.com/index.html?path=73.0.3683.68/ y extraerlo del zip.
 Aunque no es necesario tenerlo en el directorio, toca cambiar la ruta en los archivos de python.
 4. En el arcvio Detail.txt esta establecido los campos a comparar
 5. En el directorio, toca crear una carpeta que se llame "pantallazo"
 
