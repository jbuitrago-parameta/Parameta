from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import os
import credeMulti
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

#os.chdir("C:\Users\Oscar\Documents\Multi")
browser = webdriver.Chrome("C:\Python\Python37\chromedriver")
urls = pd.DataFrame(columns=['url'])

def obtener_urls():
    browser.get("https://multibank-sandbox.coupahost.com/sessions/new")
    nombre_usuario = browser.find_element_by_name("user[login]")
    nombre_usuario.clear()
    nombre_usuario.send_keys(credeMulti.usuario)
    contrasena = browser.find_element_by_name("user[password]")
    contrasena.clear()
    contrasena.send_keys(credeMulti.contrasena)
    browser.find_element_by_class_name("button").click()
    browser.find_element_by_link_text('Requests').click()
    ix = 0
    for t in range(20):
        time.sleep(7)
        print('Pagina ' + str(ix+1))
        tbody = browser.find_element_by_id("requisition_header_tbody")
        rows  = tbody.find_elements_by_tag_name("tr")
        for row in rows:
            regla = row.find_elements_by_tag_name("td")
            try:
                if regla[3].text == 'Pending Buyer Action':
                    url = regla[6].find_element_by_tag_name('a').get_attribute('href')
                    urls.loc[ix,'url'] = url
                    print(url)
                    ix = ix+1
            except:
                print(regla[0].text)

        try:
            browser.find_element_by_link_text("Next").click()
        except:
            print ("Fin test")
            break


def cambios():
    #Aca comenzaria el ciclo
    for x in range(0,len(urls)):
        browser.get(urls.loc[x,'url'])
        time.sleep(1)
        browser.find_element_by_id('bypass_approvals_and_order').click()
        #Esto es para llevar seguimiento del proceso
        print (str(x+1) + ' cambio realizado')
    #browser.close()

if __name__ == "__main__":
    obtener_urls()
    #cambios()


