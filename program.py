import time
import os
import re
import warnings
import traceback
import datetime
import shutil
from math import ceil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys #pip install selenium 
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
warnings.filterwarnings('ignore')

import undetected_chromedriver as uc #pip install undetected-chromedriver 










try:
    print('Reading input data       ', end = '\r')
    #pegando dados dos arquivos txt
    with open('login.txt', 'r', encoding='utf8') as file:
        login = file.read()
    with open('password.txt', 'r', encoding='utf8') as file:
        password = file.read()
    with open('timeout.txt', 'r', encoding='utf8') as file:
        timeout = file.read()
    for i in [' ', '\t', '\n']:
        login = login.replace(i, '')
        password = password.replace(i, '')
        timeout = timeout.replace(i, '')
    timeout = float(timeout)



    #lendo XLSX de inputs
    for diretorio, subpastas, arquivos in os.walk('Excel'):
        for file in arquivos:
            if file.count('.xls') != 0 or file.count('.xlsx') != 0:
                arquivo_XLSX = 'Excel/' + file
    df_inputs = pd.read_excel(arquivo_XLSX)
    n_rows = len(df_inputs[df_inputs.columns[0]])
    df = df_inputs


    tipooo = input('Hacer flujo normal (1) o hacer todo desde "crear guia" (2)? ')

    #agrupando trackings
    tracking_obj = {}
    tracking_list = []
    for i in range(0, n_rows):
        tracking = str(df_inputs[df_inputs.columns[7]][i])
        fecha = str(df_inputs[df_inputs.columns[12]][i])
        Status = str(df_inputs[df_inputs.columns[20]][i])
        #print(tracking)
            
        if Status != 'OK2' and fecha != 'nan' and Status != 'Tracking no encontrado':
            if tracking not in tracking_list:
                tracking_obj[tracking] = [i]
                tracking_list.append(tracking)
            else:
                tracking_obj[tracking].append(i)
    for j in range(0, len(tracking_obj)):
        indexes = tracking_obj[tracking_list[j]]
        print(f'{tracking_list[j]} - {str([x + 2 for x in indexes])}')
    


    
    #opening the web site
    print('Opening the website       ', end = '\r')
    options = Options()   
    options.add_argument("--start-maximized")
    options.add_argument("--incognito")
    options.add_argument("--disable-popup-blocking")

    driver = uc.Chrome(options=options)#para acessar no modo undetected
    dpp = f'{os.getcwd()}\\Downloads'
    params = {
        "behavior": "allow",
        "downloadPath": dpp
    }
    driver.execute_cdp_cmd("Page.setDownloadBehavior", params)
    
    driver.maximize_window()
    wait = WebDriverWait(driver, 60)
    wait_faster = WebDriverWait(driver, 2)
    wait_fast = WebDriverWait(driver, 0.1)
    




    #logging in
    driver.get('localhost')

    field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="username"]'))).send_keys(login)
    time.sleep(0.30)

    field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="password"]'))).send_keys(password)
    time.sleep(0.30)

    field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="cl-wrapper"]/center/div/div[1]/div[2]/form/div[2]/div/button'))).click()
    time.sleep(0.30)


    
    

    for j in range(0, len(tracking_obj)):
        indexes = tracking_obj[tracking_list[j]]
        n_este_tracking = 0
        tracking_encontado = True
        for i in indexes:
            n_este_tracking += 1
            venta = str(df_inputs[df_inputs.columns[2]][i])
            producto = str(df_inputs[df_inputs.columns[3]][i])
            destinatario = str(df_inputs[df_inputs.columns[4]][i])
            n_pedido = str(df_inputs[df_inputs.columns[6]][i])
            tracking = str(df_inputs[df_inputs.columns[7]][i])
            valor = str(df_inputs[df_inputs.columns[8]][i]).replace(',', '.')
            peso_exacto = str(df_inputs[df_inputs.columns[10]][i])
            fecha = str(df_inputs[df_inputs.columns[12]][i])
            desc = str(df_inputs[df_inputs.columns[16]][i])
            if len(desc) > 250:
                desc = desc[:249]
            partida_arranceraria = str(df_inputs[df_inputs.columns[17]][i])
            direcion = str(df_inputs[df_inputs.columns[18]][i])
            ciudad = str(df_inputs[df_inputs.columns[19]][i])
            Status = str(df_inputs[df_inputs.columns[20]][i])
            pais = str(df_inputs[df_inputs.columns[21]][i])
            estado = str(df_inputs[df_inputs.columns[22]][i])
            print(f'Processing tracking {j + 1}/{len(tracking_obj)} ({tracking}) - line {i + 2}/{str([x + 2 for x in indexes])} (pedido {n_pedido})')


            try:
                #driver.get('')
                if str(tipooo) == '2':
                    n_este_tracking = 2###################
                
                if n_este_tracking == 1:
                    driver.get('localhost')
                    time.sleep(5)

                    busqueda = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="globalfilter"]'))).send_keys(tracking)
                    time.sleep(12)


                    #verificando si el tracking se encuentra
                    while True:
                        try:
                            msg = wait_fast.until(EC.presence_of_element_located((By.XPATH, '//*[@id="receiptTable"]/tbody/tr/td/div'))).text
                            if msg.count('No se encuentran resultados') > 0:
                                tracking_encontado = False
                                break
                        except:
                            pass

                        try:
                            check = wait_fast.until(EC.presence_of_element_located((By.XPATH, '//*[@id="receiptTable"]/tbody/tr/td[1]/div[1]/div/ins'))).click()
                            time.sleep(0.30)
                            tracking_encontado = True

                            desdoblar = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="pcont"]/div[2]/div/div[2]/div[1]/div[1]/div/div[2]/button[2]'))).click()
                            time.sleep(0.30)
                            crear_guia = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="pcont"]/div[2]/div/div[2]/div[1]/div[1]/div/div[2]/ul/li[1]/button'))).click()
                            time.sleep(0.30)
                            break
                        except:
                            pass
                
                if tracking_encontado == False:
                    Stat = 'Tracking no encontrado'
                if tracking_encontado == True:
                    if n_este_tracking > 1:
                        driver.get('localhost')
                        time.sleep(2)

                        remitente = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-drop"]/div/input'))).send_keys(Keys.ESCAPE)
                        time.sleep(0.30)
                    
                    invoice = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type_contain"]')))
                    invoice.clear()
                    time.sleep(0.30)
                    invoice.send_keys(tracking)
                    time.sleep(0.30)

                    partida_arranceraria_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type_pa"]'))).send_keys(partida_arranceraria)
                    time.sleep(0.30)

                    COD = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="divcod"]/h4/label/div[1]/ins'))).click()
                    time.sleep(0.30)


                    #remitente
                    remitente = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="s2id_search_shipper"]/a/span[1]'))).click()
                    time.sleep(0.30)
                    remitente = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-drop"]/div/input'))).send_keys('AMAZON')
                    time.sleep(0.30)
                    remitente = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-drop"]/ul/li/div'))).click()
                    time.sleep(0.30)



                    #destinatario
                    destinatario_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="s2id_search_receiver"]/a/span[1]'))).click()
                    time.sleep(0.30)
                    destinatario_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-drop"]/div/input'))).send_keys(destinatario)
                    time.sleep(0.30)

                    
                    while True:
                        try:
                            msg = wait_fast.until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-drop"]/ul/li'))).text
                            if msg.count('No se encuentran resultados') > 0:
                                print(f'\tDestinatario no existia ({destinatario})')
                                destinatario_existia = False
                                break
                        except:
                            pass

                        try:
                            destinatario_c = wait_fast.until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-drop"]/ul/li/div'))).click()
                            time.sleep(0.30)
                            destinatario_existia = True
                            break
                        except:
                            pass

                            
                    if destinatario_existia ==  False:
                        crear = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type"]/div[2]/div[1]/div[2]/div/div[2]/button[2]')))
                        driver.execute_script("arguments[0].click();", crear)#click falso
                        time.sleep(2)
                        
                        nombre = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[3]/div[10]/div/div/div[2]/div[3]/div[1]/input')))
                        nombre = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type_name_addr"]'))).send_keys(Keys.ESCAPE)
                        time.sleep(0.30)
                        nombre = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type_name_addr"]'))).send_keys(destinatario.split(' ', 1)[0])
                        time.sleep(0.30)

                        apellido = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type_lastname_addr"]'))).send_keys(destinatario.split(' ', 1)[1])
                        time.sleep(0.30)

                        direcion_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type_direccion_addr"]'))).send_keys(direcion)
                        time.sleep(0.30)

                        '''
                        ciudad_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="s2id_cityaddr"]/a/span[1]'))).click()
                        time.sleep(0.30)
                        ciudad_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-drop"]/div/input'))).send_keys(ciudad)
                        time.sleep(0.30)
                        ciudad_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-drop"]/ul/li[1]/div'))).click()
                        time.sleep(2)
                        '''

                        no_encuentro = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="addaddr"]/div/div/div[2]/div[7]/div[2]/div/ins'))).click()
                        time.sleep(2)

                        pais_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="s2id_guide_type_country_addr"]/a/span[1]'))).click()
                        time.sleep(0.30)
                        pais_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-drop"]/div/input'))).send_keys(pais)
                        time.sleep(0.30)
                        pais_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-drop"]/ul/li[1]/div'))).click()
                        time.sleep(2)

                        estado_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="s2id_guide_type_state_addr"]/a/span[1]'))).click()
                        time.sleep(0.30)
                        estado_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-drop"]/div/input'))).send_keys(estado)
                        time.sleep(0.30)
                        estado_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-drop"]/ul/li[1]/div'))).click()
                        time.sleep(2)

                        ciudad_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type_cityname_addr"]')))
                        ciudad_c.clear()
                        time.sleep(0.30)
                        ciudad_c.send_keys(ciudad)
                        time.sleep(2)


                        #time.sleep(1111)###########

                        AGREGAR_destinatario = wait_fast.until(EC.presence_of_element_located((By.XPATH, '//*[@id="gotoaddaddr"]'))).click()
                        time.sleep(2)                
                        
                        print('\tDestinatario fue creado')




                    tipo_caja = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type_packages_0_epacktype"]/option[2]'))).click()
                    time.sleep(0.30)
                    
                    venta_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type_packages_0_tracking"]')))
                    venta_c.clear()
                    time.sleep(0.30)
                    venta_c.send_keys(venta)
                    time.sleep(0.30)

                    desc_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type_packages_0_description"]'))).send_keys(desc)
                    time.sleep(0.30)

                    valor_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type_packages_0_value"]'))).send_keys(valor)
                    time.sleep(0.30)

                    peso_exacto_c = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type_packages_0_weight"]')))
                    peso_exacto_c.clear()
                    time.sleep(0.30)
                    peso_exacto_c.send_keys(peso_exacto)
                    time.sleep(0.30)

                    tipo_producto = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="s2id_guide_type_packages_0_kindproduct"]/a/span[1]'))).click()
                    time.sleep(0.30)
                    tipo_producto = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-drop"]/ul/li[1]/div'))).click()
                    time.sleep(0.30)

                    asignar_tarifa = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="thetariff"]/div[1]/div[1]/button'))).click()
                    time.sleep(0.30)

                    while True:
                        try:
                            courier_18 = wait_faster.until(EC.presence_of_element_located((By.XPATH, '//*[@id="tariffTable"]/tbody/tr[1]'))).click()
                            break
                        except:
                            pass

                    while True:
                        try:
                            tarifa = wait_faster.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type_tariffname"]'))).get_attribute('value')
                            if tarifa == 'Courier 1.8 X LB SIN MINIMA':
                                break
                        except:
                            pass

                        try:
                            time.sleep(1)
                            courier_18 = wait_faster.until(EC.presence_of_element_located((By.XPATH, '//*[@id="tariffTable"]/tbody/tr[1]'))).click()
                        except:
                            pass

                        #try:
                        #    error_seguir = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="mod-error"]/div/div/div[3]/div/button')))
                        #    driver.execute_script("arguments[0].click();", error_seguir)#click falso
                        #    time.sleep(2)
                        #    break
                        #except:
                        #    pass

                    CREAR = wait_fast.until(EC.presence_of_element_located((By.XPATH, '//*[@id="guide_type_submit"]')))
                    driver.execute_script("arguments[0].click();", CREAR)#click falso
                    time.sleep(0.30)

                    #salvando resultado
                    ID_guia_creada = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="pcont"]/div[2]/div[1]/div[1]/div/div[1]/h3'))).text
                    df.iat[i, 11] = ID_guia_creada



                    #subir imagenes
                    if os.path.exists(f'IMG/{producto}.jpg'):
                        print('\tColocando imagen')
                        input_img = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="image"]'))).send_keys(f'{os.getcwd()}\\IMG\\{producto}.jpg')
                        time.sleep(7)
                        print('\tImagen colocada')
                    else:
                        print('\tLa imagen no fue encontrada')





                    print(f'\tCreado: {ID_guia_creada}')
                    Stat = 'OK2'
            except:
                Stat = 'error'
                print(traceback.format_exc())########
                time.sleep(10)


    
            #results
            print(f'\t{Stat}')
            df.iat[i, 20] = Stat
            df.to_excel(arquivo_XLSX, 'Sheet1', index=False)
            if Stat != 'Tracking no encontrado':
                time.sleep(timeout)
except:
    print('\n\tSome error occurred... ')
    print(traceback.format_exc())

end = input('\n\nProgram finished! Press ENTER to close')
