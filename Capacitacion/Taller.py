################################################################
############# Taller Selenium 3 de Julio #######################
################################################################

# El siguiente taller tiene como objetivo realizar una primera
# tarea utilizando Selenium. La tarea consiste en programar las
# funcion inicio_sesion() y unidades_de_medida
# Las presentaciones las puede descargar de este repositorio
# https://github.com/jjbuitrago/WebScraping

# Algunas librerias que pueden utilizar
# Recuerde que para instalar una libreria, utilice el comando
# "pip3 install libreria" en la terminal
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd

############# Escribir la ubicación del driver en esta parte ################
driver = webdriver.Chrome()
#############################################################################

def inicio_sesion():
    # El objetivo de esta función es iniciar sesión en cualquier
    # instancia de Coupa.
    pass

def unidades_de_medida():
    # El objetivo de esta función es guardar en cualquier estructura de texto
    # (lista, diccionario, data frame) las unidades de medida que soporta la
    # instancia y su código
    pass

if __name__ == "__main__":
    inicio_sesion()
    unidades_de_medida()