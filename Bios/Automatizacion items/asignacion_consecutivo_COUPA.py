### Asignacion consecutivo COUPA

import pandas as pd
import os
import datetime
from selenium import webdriver

subir = True

# Para darle el nombre al archivo
now = datetime.datetime.now()
if len(str(now.day)) ==1:
    dia = '0' + str(now.day)
else:
    dia = str(now.day)

if len(str(now.month)) ==1:
    mes = '0' + str(now.month)
else:
    mes = str(now.month)

os.chdir(r'G:\Mi unidad\Data\Grupo Bios\Automatizacion_Items\asignacion consecutivo COUPA')
consecutivos_COUPA = pd.read_excel('ultimo_codigo_coupa.xlsx', header = 0, index_col = 0)
consecutivos_COUPA_AF = pd.read_excel('ultimo_codigo_coupa_AF.xlsx', header = 0, index_col = 0)
solicitudes_dia = pd.read_excel('solicitudes_del_dia.xlsx',header = 0)
carga_tabla_homologacion = pd.DataFrame(columns=['CoupaCode', "CompanyCode", 'itemAttribute', 'itemValue','UnidadNegocio', 'disponible'], index = [0])
tabla_coupa = pd.DataFrame(columns = ['IdCompania','CoupaCode', 'Descripcion', 'Unidad','NomGrupoWA','NomSubgrupoWA','cCOUPA', 'ReferenciaSIESA', 'TipoInv','UnidadMedida', 'UnidadNegocio'],index = [0])
tags_homologacion = pd.DataFrame(columns=['COUPA','SIESA'], index = [0])
carga_masiva = pd.DataFrame(columns=['Name*'], index = [0])
verificar = pd.DataFrame(columns=['COUPA'],index = [0])
nombre_archivo = []
with open("columnas_SIESA.txt") as reader:
    for line in reader:
        a =  str(line).replace('\n','')
        carga_masiva[a] = ""


def asignacion_codigo():
    ix_homologacion = 0
    ix_carga = 0
    for x in range(0,len(solicitudes_dia)):
        if pd.isna(solicitudes_dia.loc[x,'CodigoCoupa']):
            if (solicitudes_dia.loc[x,'ActivoFijo'] =="Si") | (solicitudes_dia.loc[x,'ActivoFijo'] =="SI"):
                carga_masiva.loc[ix_carga,"active*"] = 'No'
                codigo_asignar = consecutivos_COUPA_AF.loc[solicitudes_dia.loc[x,"NomSubgrupoWA"], 'codigo']
                codigo_siguiente = codigo_asignar + 1
                consecutivos_COUPA_AF.loc[solicitudes_dia.loc[x,"NomSubgrupoWA"], 'codigo'] = codigo_siguiente
                ceros = 4 - len(str(consecutivos_COUPA_AF.loc[solicitudes_dia.loc[x,"NomSubgrupoWA"], 'codigo']))
                codigo_string = str(codigo_siguiente)
                for i in range(0,ceros):
                    codigo_string = "0" + codigo_string
                carga_masiva.loc[ix_carga,"Supplier Part Num"] = str(consecutivos_COUPA_AF.loc[solicitudes_dia.loc[x,"NomSubgrupoWA"], 'prefijo']) + codigo_string
                solicitudes_dia.loc[x,"CodigoCoupa"] = str(consecutivos_COUPA_AF.loc[solicitudes_dia.loc[x,"NomSubgrupoWA"], 'prefijo']) + codigo_string
            else:
                carga_masiva.loc[ix_carga,"active*"] = 'Yes'
                codigo_asignar = consecutivos_COUPA.loc[solicitudes_dia.loc[x,"NomSubgrupoWA"], 'codigo']
                codigo_siguiente = codigo_asignar + 1
                consecutivos_COUPA.loc[solicitudes_dia.loc[x,"NomSubgrupoWA"], 'codigo'] = codigo_siguiente
                ceros = 4 - len(str(consecutivos_COUPA.loc[solicitudes_dia.loc[x,"NomSubgrupoWA"], 'codigo']))
                codigo_string = str(codigo_siguiente)
                for i in range(0,ceros):
                    codigo_string = "0" + codigo_string
                carga_masiva.loc[ix_carga,"Supplier Part Num"] = str(consecutivos_COUPA.loc[solicitudes_dia.loc[x,"NomSubgrupoWA"], 'prefijo']) + codigo_string
                solicitudes_dia.loc[x,"CodigoCoupa"] = str(consecutivos_COUPA.loc[solicitudes_dia.loc[x,"NomSubgrupoWA"], 'prefijo']) + codigo_string
                if (str(solicitudes_dia.loc[x,'NomSubgrupoWA']) == 'ADITIVOS') | (str(solicitudes_dia.loc[x,'NomSubgrupoWA']) == 'ADITIVOS MP'):
                    consecutivos_COUPA.loc['ADITIVOS', 'codigo'] = max(consecutivos_COUPA.loc['ADITIVOS', 'codigo'],consecutivos_COUPA.loc['ADITIVOS MP', 'codigo'])
                    consecutivos_COUPA.loc['ADITIVOS MP', 'codigo'] = max(consecutivos_COUPA.loc['ADITIVOS', 'codigo'],consecutivos_COUPA.loc['ADITIVOS MP', 'codigo'])

            carga_masiva.loc[ix_carga,"Name*"] = solicitudes_dia.loc[x,'Descripcion']
            carga_masiva.loc[ix_carga,"Description*"] = solicitudes_dia.loc[x,'Descripcion']
            carga_masiva.loc[ix_carga,"UOM code*"] = solicitudes_dia.loc[x,'Unidad']
            carga_masiva.loc[ix_carga,"Item Number"] = carga_masiva.loc[ix_carga,"Supplier Part Num"] 
            carga_masiva.loc[ix_carga,"Commodity"] = solicitudes_dia.loc[x,'NomSubgrupoWA']
            carga_masiva.loc[ix_carga,"Tags"] = solicitudes_dia.loc[x,'ReferenciaSIESA']
            carga_masiva.loc[ix_carga,"Supplier"] = '00000-Servicios BIOS'
            carga_masiva.loc[ix_carga,"Contract Number"] = '00000-Servicios BIOS'
            carga_masiva.loc[ix_carga,"Price"] = '0'
            carga_masiva.loc[ix_carga,"Currency"] = 'COP'
            carga_masiva.loc[ix_carga,"Supplier Part Num"]
            carga_masiva.loc[ix_carga,"Fixed asset"] = 'No'
            carga_masiva.loc[ix_carga,"Controlled Substance"] = 'No'
            carga_masiva.loc[ix_carga,"Storable"] = 'Yes'
            carga_masiva.loc[ix_carga,"Lot tracking"] = 'No'
            verificar.loc[ix_carga, 'COUPA'] = carga_masiva.loc[ix_carga,"Item Number"]
            ix_carga = ix_carga+1

        else:
            tags_homologacion.loc[ix_homologacion,'COUPA'] = solicitudes_dia.loc[x,'CodigoCoupa']
            tags_homologacion.loc[ix_homologacion,'SIESA'] = solicitudes_dia.loc[x,'ReferenciaSIESA']
            ix_homologacion = ix_homologacion + 1
    
    solicitudes_dia.to_excel('resultado asignacion.xlsx', header=True,encoding='utf-8')
    consecutivos_COUPA.to_excel('nuevo_codigo_coupa.xlsx', header = True, encoding = 'utf-8')
    consecutivos_COUPA_AF.to_excel('nuevo_codigo_coupa_AF.xlsx', header = True, encoding = 'utf-8')
    now = datetime.datetime.now()
    carga_masiva_test = carga_masiva.copy()
    carga_masiva_test["Contract Number"] = 'Servicios BIOS 00001'
    carga_masiva_test.to_csv('G:\Mi unidad\Data\Grupo Bios\Administración de ítems\Servicios BIOS\Carga Masiva Test.csv',index=None, header = True, encoding = 'utf-8')
    nombre2 = 'G:\Mi unidad\Data\Grupo Bios\Administración de ítems\Servicios BIOS\Servicios BIOS '+ dia + mes + str(now.year) + ".csv"
    nombre_archivo.append(nombre2)
    carga_masiva.to_csv(nombre2, index=None, header = True, encoding = 'utf-8')
    verificar.to_excel(r'G:\Mi unidad\Data\Grupo Bios\Administración de ítems\Tags\verificar.xlsx', index = None, header = True, encoding = 'utf-8')
    tags_homologacion.to_excel('G:\Mi unidad\Data\Grupo Bios\Administración de ítems\Tags\homologacion.xlsx', header = True, encoding = 'utf-8')

def carga_tabla_homologacion_funcion():
    for x in range(0,len(solicitudes_dia)):
        carga_tabla_homologacion.loc[x,'CoupaCode'] = solicitudes_dia.loc[x,'CodigoCoupa']
        tabla_coupa.loc[x,'CoupaCode'] = solicitudes_dia.loc[x,'CodigoCoupa']
        tabla_coupa.loc[x,'cCOUPA']=tabla_coupa.loc[x,'CoupaCode']
        if solicitudes_dia.loc[x,'IdCompania'] == 1:
            carga_tabla_homologacion.loc[x,"CompanyCode"] = "001"
            tabla_coupa.loc[x,'IdCompania'] = "001"
        elif solicitudes_dia.loc[x,'IdCompania'] == 20:
            carga_tabla_homologacion.loc[x,"CompanyCode"] = "020"
            tabla_coupa.loc[x,'IdCompania'] = "020"
        else:
            carga_tabla_homologacion.loc[x,"CompanyCode"] = solicitudes_dia.loc[x,'IdCompania']
            tabla_coupa.loc[x,'IdCompania'] = solicitudes_dia.loc[x,'IdCompania']
        carga_tabla_homologacion.loc[x,'itemAttribute'] = 1
        carga_tabla_homologacion.loc[x,'itemValue'] = solicitudes_dia.loc[x,'ReferenciaSIESA']
        tabla_coupa.loc[x,'ReferenciaSIESA'] = solicitudes_dia.loc[x,'ReferenciaSIESA']
        if solicitudes_dia.loc[x,'UnidadNegocio'] == 99:
            carga_tabla_homologacion.loc[x,'UnidadNegocio'] = 99
            tabla_coupa.loc[x,'UnidadNegocio'] = 99
        elif solicitudes_dia.loc[x,'UnidadNegocio'] == 999:
            carga_tabla_homologacion.loc[x,'UnidadNegocio'] = 999
            tabla_coupa.loc[x,'UnidadNegocio'] = 999
        else:
            carga_tabla_homologacion.loc[x,'UnidadNegocio'] = "0" + str(solicitudes_dia.loc[x,'UnidadNegocio'])
            tabla_coupa.loc[x,'UnidadNegocio'] = "0" + str(solicitudes_dia.loc[x,'UnidadNegocio'])
        carga_tabla_homologacion.loc[x,'disponible'] = 1
        
        tabla_coupa.loc[x,'Descripcion'] = solicitudes_dia.loc[x,'Descripcion']
        tabla_coupa.loc[x,'Unidad'] = solicitudes_dia.loc[x,'Unidad']
        tabla_coupa.loc[x,'NomGrupoWA'] = solicitudes_dia.loc[x,'NomGrupoWA'] 
        tabla_coupa.loc[x,'NomSubgrupoWA'] = solicitudes_dia.loc[x,'NomSubgrupoWA']
        tabla_coupa.loc[x,'TipoInv']= solicitudes_dia.loc[x,'TipoInventario']
        tabla_coupa.loc[x,'UnidadMedida']= solicitudes_dia.loc[x,'Unidad']
        
    now = datetime.datetime.now()
    nombre = 'G:\Mi unidad\Data\Grupo Bios\Administración de ítems\Carga Tabla Homologación\PIF '+ dia + mes + str(now.year) + ".csv"
    carga_tabla_homologacion.to_csv(nombre,index = None, header = False, encoding = 'utf-8')
    tabla_coupa.to_excel('tabla_coupa.xlsx', index = None, header = True, encoding="utf-8")
    


if __name__ == "__main__":
    asignacion_codigo()
    carga_tabla_homologacion_funcion()
    if subir:
        browser_test = webdriver.Chrome("C:\Python\Python37\chromedriver")
        browser_test.get("https://grupobios-test.coupahost.com/")
        nombre_usuario = browser_test.find_element_by_name("user[login]")
        nombre_usuario.clear()
        nombre_usuario.send_keys('jbuitrago')
        contrasena = browser_test.find_element_by_name("user[password]")
        contrasena.clear()
        contrasena.send_keys('valleyball1')
        browser_test.find_element_by_class_name("button").click()
        browser_test.find_element_by_link_text('Items').click()
        browser_test.find_element_by_xpath('//*[@id="item_data_table_form_search"]/div[1]/table/tbody/tr/td[1]/a[2]/span').click()
        browser_test.find_element_by_id("data_source_file").send_keys('G:\Mi unidad\Data\Grupo Bios\Administración de ítems\Servicios BIOS\Carga Masiva Test.csv')
        browser_test.find_element_by_xpath('//*[@id="csv_upload"]/li[4]/button/span').click()

        browser_prod = webdriver.Chrome("C:\Python\Python37\chromedriver")
        browser_prod.get("https://grupobios.coupahost.com/")
        nombre_usuario = browser_prod.find_element_by_name("user[login]")
        nombre_usuario.clear()
        nombre_usuario.send_keys('ParametaData')
        contrasena = browser_prod.find_element_by_name("user[password]")
        contrasena.clear()
        contrasena.send_keys('Parametadata*')
        browser_prod.find_element_by_class_name("button").click()
        browser_prod.find_element_by_link_text('Items').click()
        browser_prod.find_element_by_xpath('//*[@id="item_data_table_form_search"]/div[1]/table/tbody/tr/td[1]/a[2]/span').click()
        browser_prod.find_element_by_id("data_source_file").send_keys(nombre_archivo[0])
        browser_prod.find_element_by_xpath('//*[@id="csv_upload"]/li[4]/button/span').click()
    