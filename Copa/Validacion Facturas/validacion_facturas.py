##Validacion de facturas COPA

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
import time
import pandas as pd
import os
import credenciales

os.chdir('G:\Mi unidad\Data\Copa\Validacion facturas')

driver = webdriver.Chrome("C:\Python\Python37\chromedriver")
tabla_facturas = pd.DataFrame(columns=['Supplier'], index = [0])
comparacion = pd.DataFrame(columns=['Supplier'], index = [0])
with open("Columnas Copa.txt") as reader:
    for line in reader:
        a =  str(line).replace('\n','')
        tabla_facturas[a] = ""
        comparacion[a] = ""
tabla_facturas = pd.DataFrame(columns = ['Supplier',	'Invoice #',	'Invoice Date',	'Currency',	'Date of Invoice Received',	'Key reference for SAP',	'From',	'To',	'Line #',	'Description',	'Price',	'Total tax (Header)',	'Contract',	'Total invoice ',	'Attachments',], index = [0])
tabla_url = pd.DataFrame(columns = ['url'], index = [0])

insumo_UAT = pd.read_excel('Insumo UAT.xlsx', header = 0)

def funcion_uno():
    ###Inicio de sesion
    driver.get('https://copaair-test.coupahost.com/sessions/support_login')
    wait = WebDriverWait(driver, 10)
    nombre_usuario = driver.find_element_by_name("user[login]")
    nombre_usuario.clear()
    nombre_usuario.send_keys(credenciales.user)
    contrasena = driver.find_element_by_name("user[password]")
    contrasena.clear()
    contrasena.send_keys(credenciales.password)
    driver.find_element_by_class_name("button").click()
    
    #Buscar las facturas correspondientes
    driver.find_element_by_link_text('Invoices').click()
    driver.find_element_by_link_text('Invoice Lines').click()
    time.sleep(1)
    vista = driver.find_element_by_id('invoice_line_filter')
    vista.send_keys('UAT Creaci')
    time.sleep(2)
    # Aca comenzaría el ciclo
    cuadro = driver.find_element_by_id('sf_invoice_line')
    ix_output = 0
    for x in range(0,len(insumo_UAT)):
        item = insumo_UAT.loc[x,'Invoice #']
        try:
            cuadro.clear()
        except:
            time.sleep(2)
            cuadro = driver.find_element_by_id('sf_invoice_line')
            cuadro.clear()
        cuadro.send_keys(item)
        cuadro.send_keys(Keys.ENTER)
        prueba = 1
        encontrar = False
        tbody = driver.find_element_by_id('invoice_line_tbody')
        first_row = tbody.find_element_by_tag_name('tr')
        tds = first_row.find_elements_by_tag_name('td')
        while True and prueba< 10:
            if tds[0].text == 'Nothing matching your search was found.':
                print('No se encontro el item %s'.format(item))
                prueba = prueba + 1
            else:
                if tds[1].text == item:
                    try:
                        tabla_facturas.loc[ix_output,'Supplier'] = tds[0].text
                        tabla_facturas.loc[ix_output,'Invoice #'] =tds[1].text
                        tabla_facturas.loc[ix_output,'Invoice Date'] =tds[2].text
                        tabla_facturas.loc[ix_output,'Currency'] =tds[3].text
                        tabla_facturas.loc[ix_output,'Date of Invoice Received'] =tds[4].text
                        tabla_facturas.loc[ix_output,'Key reference for SAP'] =tds[5].text
                        tabla_facturas.loc[ix_output,'From'] =tds[6].text
                        tabla_facturas.loc[ix_output,'To'] =tds[7].text
                        tabla_facturas.loc[ix_output,'Line #'] =tds[8].text
                        tabla_facturas.loc[ix_output,'Description'] =tds[9].text
                        tabla_facturas.loc[ix_output,'Price'] =tds[10].text
                        tabla_facturas.loc[ix_output,'Total tax (Header)'] =tds[11].text
                        tabla_facturas.loc[ix_output,'Contract'] =tds[12].text
                    except:
                        time.sleep(3)
                        tabla_facturas.loc[ix_output,'Supplier'] = tds[0].text
                        tabla_facturas.loc[ix_output,'Invoice #'] =tds[1].text
                        tabla_facturas.loc[ix_output,'Invoice Date'] =tds[2].text
                        tabla_facturas.loc[ix_output,'Currency'] =tds[3].text
                        tabla_facturas.loc[ix_output,'Date of Invoice Received'] =tds[4].text
                        tabla_facturas.loc[ix_output,'Key reference for SAP'] =tds[5].text
                        tabla_facturas.loc[ix_output,'From'] =tds[6].text
                        tabla_facturas.loc[ix_output,'To'] =tds[7].text
                        tabla_facturas.loc[ix_output,'Line #'] =tds[8].text
                        tabla_facturas.loc[ix_output,'Description'] =tds[9].text
                        tabla_facturas.loc[ix_output,'Price'] =tds[10].text
                        tabla_facturas.loc[ix_output,'Total tax (Header)'] =tds[11].text
                        tabla_facturas.loc[ix_output,'Contract'] =tds[12].text
                    url = tds[1].find_element_by_tag_name('a').get_attribute('href')
                    driver.get(url)
                    time.sleep(3)
                    attachments_list = driver.find_element_by_class_name('attachments-list')
                    attachments = attachments_list.find_elements_by_tag_name('div')
                    line = ""
                    for ata in attachments:
                        if ata.text != "":
                            line = line + ata.text + ";"
                    line = line[:-1]
                    tabla_facturas.loc[ix_output,'Attachments'] = line
                    tabla_facturas.loc[ix_output,'Total invoice '] = driver.find_element_by_id('totalWithTaxes').text
                    driver.find_element_by_link_text('Invoice Lines').click()
                    encontrar = True
                    ix_output = ix_output + 1
                    break
                else:
                    time.sleep(1)
                    prueba = prueba + 1
        if not encontrar:
            print('No se encontro el item %s'.format(item))
    row_f, col_f = tabla_facturas.shape
    for x in range(0,row_f):
        comparacion.loc[x] = ''
        for y in range(0,col_f):
            if str(tabla_facturas.iloc[x,y]).strip() == str(insumo_UAT.iloc[x,y]).strip():
                comparacion.iloc[x,y] = 'OK'
            else:
                comparacion.iloc[x,y] = str(insumo_UAT.iloc[x,y]) + "-" + str(tabla_facturas.iloc[x,y])
    tabla_facturas.to_excel('Output Facturas.xlsx',header = True, encoding = 'utf-8', index=None)
    comparacion.to_excel('Resultado Comparacion.xlsx',header = True, encoding = 'utf-8', index=None)


    '''contiene = driver.find_element_by_xpath('//*[@id="invoice_header_adv_cond_w"]/div/div/select')
    contiene.send_keys('creado')
    time.sleep(3)
    id1 = 'conditions_' + str(contiene.get_attribute('id')).split('_')[1].strip() + '_created_by'
    driver.find_element_by_id(id1).send_keys('PIF')
    driver.find_element_by_xpath('//*[@class="table_condition_button"]/a[2]/img').click()
    time.sleep(1)
    contiene2 = driver.find_element_by_xpath('//*[@id="invoice_header_adv_cond_w"]/div/div[2]/select')
    contiene2.send_keys('Estad')
    time.sleep(3)
    driver.find_element_by_xpath('//*[@class="condition_clause"]/select/option[11]').click()
    driver.find_element_by_id('search_advanced_button_invoice_header').click()
    time.sleep(2)
    i=0
    for t in range(3):
        tbody = driver.find_element_by_id('invoice_header_tbody')
        rows = tbody.find_elements_by_tag_name('tr')
        for row in rows:
            link = row.find_element_by_tag_name("td").find_element_by_tag_name("a").get_attribute("href")
            tabla_url.loc[i,'url'] = link
            i=i+1
        try:
            driver.find_element_by_link_text("Siguiente").click()
        except:
            print ("Se obtuvieron todas las url")
            break
    
    for i in range(len(tabla_url)):
        driver.get(tabla_url.loc[i,'url'])
        wait.until(EC.element_to_be_clickable((By.ID, 'add_comment_link')))
        tabla_facturas.loc[i,'N de factura'] = driver.find_element_by_id('invoice_invoice_number').text
        tabla_facturas.loc[i,'Desde'] = driver.find_element_by_xpath('//*[@id="topHalf"]/div[1]/div[26]/time').text
        tabla_facturas.loc[i,'Hasta'] = driver.find_element_by_xpath('//*[@id="topHalf"]/div[1]/div[27]/time').text
        tabla_facturas.loc[i,'Suma Extended Price'] = driver.find_element_by_id('invoice_amount_line_price').text
        tabla_facturas.loc[i,'Impuestos'] = driver.find_element_by_xpath('//*[@class="tax_section"]/span[2]/span[3]/div/span').text
    
    tabla_facturas.to_excel('Resultado facturas.xlsx', header = True, index = None, encoding = 'utf-8')
    print ('Exportacion exitosa. Fin de la validacion')
    driver.close()'''

if __name__ == "__main__":
    funcion_uno()