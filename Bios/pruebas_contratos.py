from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import os
import configTest
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

# Definir el directorio
os.chdir("G:\Mi unidad\Data\Grupo Bios\Validacion contratos")
# Crear los drivers como variables globales
driver_test = webdriver.Chrome("C:\Python\Python37\chromedriver")
driver_bios = webdriver.Chrome("C:\Python\Python37\chromedriver")
#Importar la base de datos a trabajar 
data_frame =pd.read_excel('pruebas_contratos.xlsx', header =0)
#Crear un arreglo con los componentes a comparar de cada contrato
data_test = []
with open("Detail.txt") as reader:
    for line in reader:
        a =  str(line).replace('\n','')
        data_test.append(a)
        data_frame[a] = "NaN"

#Para utilizar los primeros 5 contratos
#data_frame = data_frame.iloc[0:5,]

def pruebas():
    pruebas = 0
    fallos = 0
    #Inicio de sesion en test
    driver_test.get("https://grupobios-test.coupahost.com/")
    nombre_usuario = driver_test.find_element_by_name("user[login]")
    nombre_usuario.clear()
    nombre_usuario.send_keys(configTest.user_test)
    contrasena = driver_test.find_element_by_name("user[password]")
    contrasena.clear()
    contrasena.send_keys(configTest.password_test)
    driver_test.find_element_by_class_name("button").click()
    #Inicio de sesion en bios
    driver_bios.get("https://grupobios-test.coupahost.com/")
    nombre_usuario = driver_bios.find_element_by_id('user_login')
    nombre_usuario.clear()
    nombre_usuario.send_keys(configTest.user_test)
    contrasena = driver_bios.find_element_by_id('user_password')
    contrasena.clear()
    contrasena.send_keys(configTest.password_test)
    driver_bios.find_element_by_class_name("button").click()
    #Crear una instancia de Webdriver. Esto es para que el navegador espere hasta que se cumpla una condicion
    wait = WebDriverWait(driver_bios, 10)
    #Ciclo que se hace en cada una de las filas de la tabla
    for x in range(0,len(data_frame)):
        screenshot = False
        #Se cargan as paginas correspondientes
        driver_test.get(data_frame.loc[x,'url_test'])
        driver_bios.get(data_frame.loc[x,'url_bios'])
        #Este es un  ciclo que compara cada uno de los componentes que se desean comprarar de pagina
        for line in data_test:
            #La pagina espera hasta que aparezca el elemento en id=field_supplier
            wait.until(EC.element_to_be_clickable((By.ID, 'field_supplier')))
            #Si al comprarar son iguales, que no haga nada. En otro caso, deja en la base que elemnto no coincide
            try:
                if driver_test.find_element_by_id(line).text == driver_bios.find_element_by_id(line).text:
                    data_frame.loc[x,line] = "Ok"
                    continue
                else:
                    data_frame.loc[x,line] = str(driver_test.find_element_by_id(line).text) + "/" + str(driver_bios.find_element_by_id(line).text)
                    screenshot=True
            except:
                data_frame.loc[x,line] = "No tiene"
        #Lleva la cuenta de cuantas pruebas van
        #print ('Fin de la prueba ' + str(x))
        pruebas = pruebas + 1
        if screenshot == True:
            #nom_arch_1 = 'pantallazo/' + str(data_frame.loc[x,'Numero']) + "-1.png"
            elem1 = driver_bios.find_element_by_class_name('columnsWrapper')
            elem1.location_once_scrolled_into_view   
            driver_bios.save_screenshot('pantallazo/' + str(data_frame.loc[x,'Numero']) + "-1.png")
            nom_arch_2 = 'pantallazo/' + str(data_frame.loc[x,'Numero'])+"-2.png"
            elem2 = driver_test.find_element_by_class_name('columnsWrapper')
            elem2.location_once_scrolled_into_view  
            driver_test.save_screenshot('pantallazo/' + str(data_frame.loc[x,'Numero'])+"-2.png")
            fallos = fallos + 1
    #Al final exporta a un excel los resultados
    driver_bios.close()
    driver_test.close()
    data_frame.to_excel('resultado_validacion.xlsx', index = None, header=True,encoding='utf-8')
    print("Se realizaron %s pruebas. Se encontraron diferencias en %d" % (pruebas, fallos))

if __name__ == "__main__":
    pruebas()