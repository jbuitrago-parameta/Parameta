#El siguiente programa sirve para exportar una base de datos que contiene el numero del contrato,
#la url en test y la url en bios. Además exporta un archivo con el numero de todos los contratos
#y la url, tanto en test como en bios

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import os
import configTest

#########Cambiar el directorio a su respectiva carpeta###########
os.chdir("G:\Mi unidad\Data\Grupo Bios\Validacion contratos")

contratos = pd.DataFrame(columns=['Numero','url'])
contratos_bios = pd.DataFrame(columns=['Numero','url'])

# Se crean instancias de los browser como variables globales
# Es necesario tener el brower descargado y cambiar la dirección en donde esta el archivo
driver_test = webdriver.Chrome("C:\Python\Python37\chromedriver")
driver_bios = webdriver.Chrome("C:\Python\Python37\chromedriver")


def nombres_url_test():
    driver_test.get("https://grupobios-test.coupahost.com/")
    nombre_usuario = driver_test.find_element_by_name("user[login]")
    nombre_usuario.clear()
    nombre_usuario.send_keys(configTest.user_test)
    contrasena = driver_test.find_element_by_name("user[password]")
    contrasena.clear()
    contrasena.send_keys(configTest.password_test)
    driver_test.find_element_by_class_name("button").click()
    driver_test.find_element_by_link_text('Suppliers').click()
    driver_test.find_element_by_link_text('Contracts').click()
    i=0
    for t in range(3):
        time.sleep(3)
        tbody = driver_test.find_element_by_id("contract_tbody")
        rows  = tbody.find_elements_by_tag_name("tr")        
        #Con este ciclo obtengo cada uno de los # de contrato
        for row in rows:
            try:
                regla = row.find_elements_by_tag_name("td")
                if str(regla[5].text) == "Published":
                    name = row.find_element_by_tag_name("td").text
                    contratos.loc[i,"Numero"]=name
                    link = row.find_element_by_tag_name("td").find_element_by_tag_name("a").get_attribute("href")
                    contratos.loc[i,'url']=link
                else:
                    continue
            except:
                continue
            i=i+1
        print ("Pagina " + str(t+1))
        try:
            driver_test.find_element_by_link_text("Next").click()
        except:
            print ("Fin test")
            break
    driver_test.close()
    contratos.to_excel('contratos_test.xlsx', index = None, header=True,encoding='utf-8')
    print ("Exportacion exitosa test")

def nombres_url_bios():
    driver_bios.get("https://grupobios.coupahost.com/")
    nombre_usuario = driver_bios.find_element_by_id('user_login')
    nombre_usuario.clear()
    nombre_usuario.send_keys(configTest.user)
    contrasena = driver_bios.find_element_by_id('user_password')
    contrasena.clear()
    contrasena.send_keys(configTest.password)
    driver_bios.find_element_by_class_name("button").click()
    driver_bios.find_element_by_link_text('Suppliers').click()
    driver_bios.find_element_by_link_text('Contracts').click()
    i=0
    for t in range(3):
        time.sleep(3)
        tbody = driver_bios.find_element_by_id("contract_tbody")
        rows  = tbody.find_elements_by_tag_name("tr")        
        #Con este ciclo obtengo cada uno de los # de contrato
        for row in rows:
            try:
                regla = row.find_elements_by_tag_name("td")
                if str(regla[5].text) == "Published":
                    name = row.find_element_by_tag_name("td").text
                    contratos_bios.loc[i,"Numero"]=name
                    link = row.find_element_by_tag_name("td").find_element_by_tag_name("a").get_attribute("href")
                    contratos_bios.loc[i,'url']=link
                else:
                    continue
            except:
                continue
            i=i+1
        print ("Pagina " + str(t+1))
        try:
            driver_bios.find_element_by_link_text("Next").click()
        except:
            print ("Fin bios")
            break
    driver_bios.close()
    contratos_bios.to_excel('contratos_bios.xlsx', index = None, header=True,encoding='utf-8')
    print ("Exportacion exitosa bios")

def match():
    contratos_bios_mod = contratos_bios
    i=0
    match_df = pd.DataFrame(columns=['Numero','url_test','url_bios'])
    for test in contratos.itertuples():
        z=False
        for bios in contratos_bios_mod.itertuples():
            if test[1] == bios[1]:
                match_df.loc[i,'Numero'] = test[1]
                match_df.loc[i,'url_test'] = test[2]
                match_df.loc[i,'url_bios'] = bios[2]
                a = int(bios[0])
                contratos_bios_mod = contratos_bios_mod.drop(a)
                i=i+1
                z = True
                break
        if z == False:
            match_df.loc[i,'Numero'] = test[1]
            match_df.loc[i,'url_test'] = test[2]
            match_df.loc[i,'url_bios'] = "no tiene"
            i=i+1
    match_df.to_excel('duplas_test_bios.xlsx', index = None, header=True,encoding='utf-8')
    print ("Exportacion exitosa match")


if __name__ == "__main__":
    nombres_url_test()
    nombres_url_bios()
    match()