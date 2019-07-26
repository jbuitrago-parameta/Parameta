###############################
#### Automaztización UAT´s ####
###############################

from selenium import webdriver
import os
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

os.chdir(r"G:\Mi unidad\DATA\Multibank\Python Selenium\UAT")

dict_info = {}

with open("datos UAT.txt") as reader:
    for line in reader:
        a =  str(line).replace('\n','')
        key = a.split("->")[0].strip()
        valor = a.split("->")[1].strip()
        dict_info[key] = valor

browser = webdriver.Chrome("C:\Python\Python37\chromedriver")

def cadena():
    browser.get(dict_info['url'])
    nombre_usuario = browser.find_element_by_name("user[login]")
    nombre_usuario.clear()
    nombre_usuario.send_keys(dict_info['user'])
    contrasena = browser.find_element_by_name("user[password]")
    contrasena.clear()
    contrasena.send_keys(dict_info['password'])
    browser.find_element_by_class_name("button").click()
    browser.find_element_by_id('cart').click()

    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, "approvers")))
    sec_aprobadores = browser.find_element_by_id('approvers')
    #lista_aprobadores = sec_aprobadores.find_elements_by_class_name('name-date')
    #lista_aprobadores = sec_aprobadores.find_elements_by_xpath('//*[@class ="name-date"]/div')
    lista_aprobadores = sec_aprobadores.find_elements_by_xpath('//*[@class="approver-inner hover_target"]/div[2]/div[1]')
    for l in lista_aprobadores:
        print (l.text)
    #print(browser.find_element_by_xpath('//*[@id="approvers"]/div[4]/div[1]/div[2]/div[1]').text)

def orden_de_compra():
    ##### Inicio de session
    browser.get(dict_info['url'])
    nombre_usuario = browser.find_element_by_name("user[login]")
    nombre_usuario.clear()
    nombre_usuario.send_keys(dict_info['user'])
    contrasena = browser.find_element_by_name("user[password]")
    contrasena.clear()
    contrasena.send_keys(dict_info['password'])
    browser.find_element_by_class_name("button").click()

    # Añadir elemento al carrito
    browser.find_element_by_id('need_input').send_keys(dict_info['item'])
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, "go")))
    browser.find_element_by_id('go').click()
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, "quantity")))
    quants = browser.find_element_by_id('quantity')
    quants.clear()
    quants.send_keys(dict_info['cantidad'])
    val = browser.find_element_by_id('supplier_item').get_attribute('value')
    cart_number = 'add_to_cart_' + str(val)
    browser.find_element_by_id(cart_number).click()
    browser.find_element_by_id('cart').click()

    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID,"pageHeader")))
    browser.find_element_by_id("requisition_header_justification").send_keys('Recordatorio evento')
    browser.find_element_by_id("requisition_header_department_id").send_keys(dict_info['Departamento'])
    browser.find_element_by_xpath('//*[@id="requisition_line_tbody"]/tr/td/div/div[3]/div/span/span[2]/a/img').click()

    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID,"picker_account_table")))
    browser.find_element_by_id("account_account_type_id").send_keys(dict_info['Chart of Accounts'])
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="account_segment_2_lv_id_chosen"]/a/span/ul')))

    campos = [dict_info['Cuenta Contable'],dict_info['Centro de Costos'],dict_info['SDC'],dict_info['Linea SDC']]
    for x in range(2,6):
        path = '//*[@id="account_segment_' + str(x) + '_lv_id_chosen"]/a/span/ul'
        browser.find_element_by_xpath(path).click()
        try:
            n = x+ 6
            xp = '/html/body/div[' + str(n) + ']/div/ul/li/input'
            browser.find_element_by_xpath(xp).send_keys(str(campos[x-2]))
            time.sleep(2)
            browser.find_element_by_xpath(xp).send_keys(Keys.TAB)
            time.sleep(2)
        except:
            n = x + 7
            xp = '/html/body/div[' + str(n) + ']/div/ul/li/input'
            browser.find_element_by_xpath(xp).send_keys(str(campos[x-2]))
            time.sleep(2)
            browser.find_element_by_xpath(xp).send_keys(Keys.TAB)
            time.sleep(2)

    browser.find_element_by_xpath('//*[@class="page_buttons_right"]/a').click()

    # Aprobadores
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,"//*[@id='approvers']/div[3]/div/div[2]/div")))
    sec_aprobadores = browser.find_element_by_id('approvers')
    lista_aprobadores = sec_aprobadores.find_elements_by_class_name('name')
    for l in lista_aprobadores:
        print (l.text)
    #browser.find_element_by_id('submit_for_approval_link').click()

def factura():
    ##### Inicio de session
    browser.get("https://multibank-sandbox.coupahost.com/sessions/new")
    nombre_usuario = browser.find_element_by_name("user[login]")
    nombre_usuario.clear()
    nombre_usuario.send_keys(credeMulti2.usuario)
    contrasena = browser.find_element_by_name("user[password]")
    contrasena.clear()
    contrasena.send_keys(credeMulti2.contrasena)
    browser.find_element_by_class_name("button").click()

    # Agregar la SDC al carrito
    browser.find_element_by_id('need_input').send_keys('Linea de SDC')
    browser.find_element_by_id('go').click()
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'supplier_item')))
    val = browser.find_element_by_id('supplier_item').get_attribute('value')
    cart_number = 'add_to_cart_' + str(val)
    browser.find_element_by_id(cart_number).click()
    browser.find_element_by_id(cart_number).click()
    browser.find_element_by_id('cart').click()

    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID,"pageHeader")))
    browser.find_element_by_id("requisition_header_justification").send_keys('Recordatorio evento')
    browser.find_element_by_xpath('//*[@id="requisition_line_tbody"]/tr/td/div/div[3]/div/span/span[2]/a/img').click()
    
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID,"picker_account_table")))
    browser.find_element_by_id("account_account_type_id").send_keys('SDC Mul')

    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="account_segment_1_lv_id_chosen"]/a/span/ul')))

    campos = ['5404801702000000','Comunicaci']
    for x in range(1,3):
        path = '//*[@id="account_segment_' + str(x) + '_lv_id_chosen"]/a/span/ul'
        browser.find_element_by_xpath(path).click()
        try:
            n = x+ 6
            xp = '/html/body/div[' + str(n) + ']/div/ul/li/input'
            browser.find_element_by_xpath(xp).send_keys(str(campos[x-1]))
            time.sleep(2)
            browser.find_element_by_xpath(xp).send_keys(Keys.TAB)
            time.sleep(2)
        except:
            n = x + 7
            xp = '/html/body/div[' + str(n) + ']/div/ul/li/input'
            browser.find_element_by_xpath(xp).send_keys(str(campos[x-1]))
            time.sleep(2)
            browser.find_element_by_xpath(xp).send_keys(Keys.TAB)
            time.sleep(2)

    browser.find_element_by_xpath('//*[@class="page_buttons_right"]/a').click()     
    #browser.find_element_by_id('submit_for_approval_link').click()

if __name__ == "__main__":
    orden_de_compra()
    #factura()
    #cadena()