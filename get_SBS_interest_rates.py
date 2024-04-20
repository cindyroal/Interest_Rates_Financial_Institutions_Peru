# -*- coding: utf-8 -*-

Author: Cindy Rojas


from selenium import webdriver
from selenium.webdriver.support.ui import Select 
import datetime
from datetime import datetime as dt, timedelta, date
import pandas as pd
from pandas.tseries.offsets import BDay,BMonthEnd
import time
import calendar
import locale
import numpy as np
from dateutil.relativedelta import *
import os
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import SessionNotCreatedException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


options = Options()
options.headless = True
options.add_experimental_option("excludeSwitches", ["enable-logging"])

#To begin the scrapping
locale.setlocale(locale.LC_ALL, 'Spanish_Peru.1252')


#Put the root where the executable chromedriver is located (chromedriver.exe)
path_driver = r'C:\Users\cindy\Downloads\chromedriver_win32\chromedriver.exe'

#Put the root and name of the new file that contains the dates to scrap
path_file_fecha_consulta = r"C:\Users\cindy\Dropbox (IDB RES)\Fintech\data\raw\SBS\Interest_rates_by_FIKind"

grupo_path_page = ((1,"https://www.sbs.gob.pe/app/pp/EstadisticasSAEEPortal/Paginas/TIActivaTipoCreditoEmpresa.aspx?tip=B"),
                   (2,'https://www.sbs.gob.pe/app/pp/EstadisticasSAEEPortal/Paginas/TIActivaTipoCreditoEmpresa.aspx?tip=F'),
                   (3,"https://www.sbs.gob.pe/app/pp/EstadisticasSAEEPortal/Paginas/TIActivaTipoCreditoEmpresa.aspx?tip=C"),
                   (4,"https://www.sbs.gob.pe/app/pp/EstadisticasSAEEPortal/Paginas/TIActivaTipoCreditoEmpresa.aspx?tip=R"))


dates=['2014-08-29', '2014-09-30', '2014-10-31', '2016-11-28',
                '2014-12-31', '2015-01-30', '2015-02-27', '2015-03-31',
                '2015-04-30', '2015-05-29', '2015-06-30', '2015-07-31',
                '2015-08-31', '2015-09-30', '2015-10-30', '2015-11-30',
                '2015-12-31']

"""
dates=['2016-01-29', '2016-02-29', '2016-03-31', '2016-04-29',
                '2016-05-31', '2016-06-30', '2016-07-29', '2016-08-31',
                '2016-09-30', '2016-10-31', '2016-11-30', '2016-12-30',
                '2017-01-31', '2017-02-28', '2017-03-31', '2017-04-28',
                '2017-05-31', '2017-06-30', '2017-07-31', '2017-08-31',
                '2017-09-29', '2017-10-31', '2017-11-30', '2017-12-29',
                '2018-01-31', '2018-02-28', '2018-03-30', '2018-04-30',
                '2018-05-31', '2018-06-29', '2018-07-31', '2018-08-31',
                '2018-09-28', '2018-10-31', '2018-11-30', '2018-12-31',
                '2019-01-31', '2019-02-28', '2019-03-29', '2019-04-30',
                '2019-05-31', '2019-06-28', '2019-07-31', '2019-08-30',
                '2019-09-30', '2019-10-31', '2019-11-29', '2019-12-31',
                '2020-01-31', '2020-02-28', '2020-03-31', '2020-04-30',
                '2020-05-29', '2020-06-30', '2020-07-31', '2020-08-31',
                '2020-09-30', '2020-10-30', '2020-11-30', '2020-12-31',
                '2021-01-29', '2021-02-26', '2021-03-31', '2021-04-30',
                '2021-05-31', '2021-06-30', '2021-07-30', '2021-08-31',
                '2021-09-30', '2021-10-29', '2021-11-30', '2021-12-31']
"""

#Set date list
dates2=[]

for d in dates:
    fe = d[-2:] + "/" + d[5:7] + "/" + d[:4]
    dates2.append(fe)


l_tasasInteres_error = []
df_diario = pd.DataFrame({'Fecha':[],'Grupo':[],'Entidad':[],'Categoria':[],'SubCategoria':[],'Moneda':[],'Tasa':[]})
#
sesion_err = 0

print('Ini...')
for x in dates2:
    
    fecha_fin = datetime.datetime.strptime(x, '%d/%m/%Y')
    mes = fecha_fin.strftime('%B').capitalize()
    anio = fecha_fin.strftime('%Y')
    
    
    for grupo,path_page in grupo_path_page: 
        try:
                        
            driver = webdriver.Chrome(executable_path=r'C:/Users/cindy/Downloads/chromedriver_win32 (1)/chromedriver.exe')            
            driver.get(path_page)    
                    
            
            df_moneda = pd.DataFrame()
            df_mn = pd.DataFrame()
            df_me = pd.DataFrame()
            
            
            driver_cate = driver.find_element_by_id("ctl00_cphContent_rpgActualMn_OT")
            
            elems = WebDriverWait(driver_cate,10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".rpgRowHeaderField")))
                
            subcategoria = [elem.text for elem in elems]
            
            if grupo in [1,2] :                
                
                time.sleep(3)
                
                tasas_input = driver.find_element_by_xpath("//*[@id='ctl00_cphContent_rdpDate_dateInput']")
                tasas_input.clear()
                tasas_input.send_keys(x)
                
                time.sleep(3)
                
                buttonConsultar = driver.find_element_by_xpath("//*[@id='ctl00_cphContent_btnConsultar']")
                
                buttonConsultar.click()
                
                time.sleep(10)
                       
                df_mn=pd.read_html(driver.find_element_by_id("ctl00_cphContent_rpgActualMn_OT").get_attribute('outerHTML'))[1]
                
                buttonConsultarME = driver.find_element_by_xpath("//*[@id='ctl00_cphContent_lbtnMex']")
                
                buttonConsultarME.click()
                
                time.sleep(7)
                
                df_me=pd.read_html(driver.find_element_by_id("ctl00_cphContent_rpgActualMex_OT").get_attribute('outerHTML'))[1]
                
                time.sleep(5)        
            
            else :

                time.sleep(3)
                
                mes = fecha_fin.strftime('%B').capitalize()
                anio = fecha_fin.strftime('%Y')            
                
                driver.find_element_by_id("ctl00_cphContent_rMes").click()
                select = driver.find_element_by_id("ctl00_cphContent_rMes_DropDown")
                
                time.sleep(3)
                
                select = select.find_element_by_class_name("rddlPopup")
                select = select.find_element_by_class_name("rddlList")
                
                
                select = select.find_element_by_xpath("//li[text()='{0}']".format(mes))
                select.click() 
                
                time.sleep(3)
                
                driver.find_element_by_id("ctl00_cphContent_rAnio").click()
                select = driver.find_element_by_id("ctl00_cphContent_rAnio_DropDown")
                
                time.sleep(3)
                
                select = select.find_element_by_class_name("rddlPopup")
                select = select.find_element_by_class_name("rddlList")
                
                
                select = select.find_element_by_xpath("//li[text()='{0}']".format(anio))
                select.click() 
                                
                time.sleep(3)
                
                buttonConsultar = driver.find_element_by_xpath("//*[@id='ctl00_cphContent_btnConsultaMensual']")
                
                buttonConsultar.click()
                
                time.sleep(10)
                
                df_mn=pd.read_html(driver.find_element_by_id("ctl00_cphContent_rpgActualMn_OT").get_attribute('outerHTML'))[1]
                
                buttonConsultarME = driver.find_element_by_xpath("//*[@id='ctl00_cphContent_lbtnMex']")
                
                buttonConsultarME.click()
                
                time.sleep(7)
                
                df_me=pd.read_html(driver.find_element_by_id("ctl00_cphContent_rpgActualMex_OT").get_attribute('outerHTML'))[1]
                
                time.sleep(5)
  
            
            prueba = df_mn
            df_mn['SubCategoria'] = subcategoria
            df_mn = df_mn.set_index('SubCategoria')
            df_mn = df_mn.stack().to_frame()            
            df_mn = df_mn.reset_index()
            df_mn.columns = ['SubCategoria','Entidad','Tasa']
            df_mn.insert(loc=2, column='Moneda', value=['' for i in range(df_mn.shape[0])])            
            df_mn['Moneda'] = 1
            
            df_me['SubCategoria'] = subcategoria
            df_me = df_me.set_index('SubCategoria')
            df_me = df_me.stack().to_frame()            
            df_me = df_me.reset_index()
            df_me.columns = ['SubCategoria','Entidad','Tasa']
            df_me.insert(loc=1, column='Moneda', value=['' for i in range(df_me.shape[0])])
            df_me['Moneda'] = 2
            
            df_moneda = df_mn[['Entidad','SubCategoria','Moneda','Tasa']].append(df_me[['Entidad','SubCategoria','Moneda','Tasa']])
            df_moneda.insert(loc=0, column='Fecha', value=['' for i in range(df_moneda.shape[0])])
            df_moneda.insert(loc=1, column='Grupo', value=['' for i in range(df_moneda.shape[0])])
            df_moneda['Fecha'] = x
            df_moneda['Grupo'] = grupo
            
            df_diario = df_diario.append(df_moneda)
            
            driver.close()
                
        except SessionNotCreatedException as e:
            sesion_err = 1
            l_tasasInteres_error.append([x,grupo,e])

            break
        except Exception as e :
             l_tasasInteres_error.append([x,grupo,e])
             driver.close()
             pass         
          
    fecha_fin = fecha_fin - BDay(1)

if len(l_tasasInteres_error) > 0:
    df_error = pd.DataFrame.from_records(l_tasasInteres_error)
    df_error.columns = ['Fecha','Grupo','Error']
    df_error["Grupo"] = np.where(df_error["Grupo"] == 1, "Bancos",
                         np.where(df_error["Grupo"] == 2, "Financieras",
                         np.where(df_error["Grupo"] == 3, "Cajas municipales",
                         np.where(df_error["Grupo"] == 4, "Cajas rurales",np.nan))))
    l_tasasInteres_error = []

else : 
    df_error = pd.DataFrame()
    
#This will let us know which dates were not find in the query. 
df_error.to_excel(path_file_fecha_consulta + 'TasasInteresPromedio_log.xlsx', index=False)


df_diario["Grupo"] = np.where(df_diario["Grupo"] == 1, "Bancos",
                               np.where(df_diario["Grupo"] == 2, "Financieras",
                               np.where(df_diario["Grupo"] == 3, "Cajas municipales",
                               np.where(df_diario["Grupo"] == 4, "Cajas rurales",np.nan))))

df_diario["Categoria"] = np.where(df_diario["SubCategoria"] == 'Corporativos', df_diario["SubCategoria"],
                               np.where(df_diario["SubCategoria"] == 'Grandes Empresas', df_diario["SubCategoria"],
                               np.where(df_diario["SubCategoria"] == 'Medianas Empresas', df_diario["SubCategoria"],
                               np.where(df_diario["SubCategoria"] == 'Pequeñas Empresas', df_diario["SubCategoria"],
                               np.where(df_diario["SubCategoria"] == 'Microempresas', df_diario["SubCategoria"],
                               np.where(df_diario["SubCategoria"] == 'Consumo', df_diario["SubCategoria"], 
                               np.where(df_diario["SubCategoria"] == 'Hipotecarios', df_diario["SubCategoria"],np.nan)))))))

df_diario = df_diario.fillna(method="ffill")
df_diario["Moneda"] = np.where(df_diario["Moneda"] == 1, "MN","ME")
df_diario['Tasa'] = df_diario['Tasa'].astype(str).replace('s.i.',0).replace("-",0).astype(float)
df_diario = df_diario.reset_index(drop=True)
df_diario = df_diario[['Fecha','Grupo','Entidad','Categoria','SubCategoria','Moneda','Tasa']]

#Final dataset
df_diario.to_excel(path_file_fecha_consulta + "TasasInteresPromedio_2014_Dec2015.xlsx",index=False)
df_diario.to_csv(path_file_fecha_consulta + "TasasInteresPromedio_2014_Dec2015.csv",index=False)

"""
df_diario.to_excel(path_file_fecha_consulta + "TasasInteresPromedio_2016_Dec2021.xlsx",index=False)
df_diario.to_csv(path_file_fecha_consulta + "TasasInteresPromedio_2016_Dec2021.csv",index=False)
"""
print('fin..')